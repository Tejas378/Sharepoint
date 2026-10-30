import { DIGEST_PLACEHOLDER } from "./request";
import { isDigestRejection } from "./digest";
import type { DigestProvider } from "./digest";
import { preview, redactHeaders, redactText } from "./redact";
import {
  AttemptRecord,
  ExecutionRecord,
  HttpMethod,
  HttpRequestSpec,
  ResponseClass,
  Transport,
  TransportResponse,
} from "./types";

const SAFE_TO_RETRY: HttpMethod[] = ["GET", "HEAD", "OPTIONS"];

export interface ExecutorOptions {
  /** Extra attempts after a digest rejection. Default 1. */
  maxDigestRetries?: number;
  /** Extra attempts for throttling / transient failures. Default 2. */
  maxTransientRetries?: number;
  /** Retry transient failures for writes. Off by default (non-idempotent). */
  retryTransientForWrites?: boolean;
  /** Preview length stored in the execution record. Default 2000. */
  previewLength?: number;
  sleep?: (ms: number) => Promise<void>;
}

export interface ExecuteOptions {
  /** Skip digest injection (e.g. when deliberately testing an invalid digest). */
  skipDigestInjection?: boolean;
  /** Override the configured retry limits for a single execution. */
  maxDigestRetries?: number;
  maxTransientRetries?: number;
}

/**
 * Sends exactly one logical request, with bounded retries for digest rejection
 * and transient failures. It never retries destructive operations unless the
 * caller explicitly opts in.
 */
export class RequestExecutor {
  private readonly transport: Transport;
  private readonly digest: DigestProvider | null;
  private readonly options: ExecutorOptions;
  private readonly sleep: (ms: number) => Promise<void>;

  constructor(transport: Transport, digest?: DigestProvider | null, options?: ExecutorOptions) {
    if (typeof transport !== "function") {
      throw new Error("RequestExecutor requires a transport function.");
    }
    this.transport = transport;
    this.digest = digest || null;
    this.options = options || {};
    this.sleep =
      this.options.sleep ||
      function (ms: number): Promise<void> {
        return new Promise<void>(function (resolve) {
          setTimeout(resolve, ms);
        });
      };
  }

  public async execute(spec: HttpRequestSpec, options?: ExecuteOptions): Promise<ExecutionRecord> {
    const opts = options || {};
    const maxDigestRetries =
      opts.maxDigestRetries !== undefined
        ? opts.maxDigestRetries
        : this.options.maxDigestRetries !== undefined
        ? this.options.maxDigestRetries
        : 1;
    const maxTransientRetries =
      opts.maxTransientRetries !== undefined
        ? opts.maxTransientRetries
        : this.options.maxTransientRetries !== undefined
        ? this.options.maxTransientRetries
        : 2;

    const attempts: AttemptRecord[] = [];
    const startedAt = new Date().toISOString();
    let digestRetries = 0;
    let transientRetries = 0;
    let current = spec;
    let lastResponse: TransportResponse | null = null;
    let lastClass: ResponseClass = "unexpected";
    let lastError = "";

    try {
      if (!opts.skipDigestInjection && spec.requiresDigest) {
        current = await this.injectDigest(spec);
      }
    } catch (error) {
      return this.networkFailure(spec, startedAt, attempts, error);
    }

    while (true) {
      const attemptNumber = attempts.length + 1;
      const attemptStart = nowMs();
      let response: TransportResponse;
      try {
        response = await this.transport(current);
      } catch (error) {
        const duration = nowMs() - attemptStart;
        lastClass = "network";
        lastError = error instanceof Error ? error.message : String(error);
        attempts.push({
          attempt: attemptNumber,
          status: 0,
          responseClass: "network",
          durationMs: duration,
          digestRefreshed: false,
        });
        if (this.canRetryTransient(current, transientRetries, maxTransientRetries, "network")) {
          transientRetries++;
          await this.sleep(backoffMs(transientRetries, undefined));
          continue;
        }
        break;
      }

      const duration = nowMs() - attemptStart;
      lastResponse = response;
      const digestRejected = isDigestRejection(response.status, response.bodyText);
      lastClass = classifyResponse(response.status, response.bodyText);
      attempts.push({
        attempt: attemptNumber,
        status: response.status,
        responseClass: lastClass,
        durationMs: duration,
        digestRefreshed: false,
      });

      if (digestRejected && this.digest && digestRetries < maxDigestRetries) {
        // The rejected request did not apply, so retrying is safe even for writes.
        digestRetries++;
        this.digest.invalidate();
        attempts[attempts.length - 1].digestRefreshed = true;
        try {
          const refreshed = await this.digest.get(true);
          current = replaceDigestHeader(current, refreshed);
        } catch (error) {
          lastError = error instanceof Error ? error.message : String(error);
          break;
        }
        continue;
      }

      if (
        (lastClass === "throttled" || lastClass === "transient") &&
        this.canRetryTransient(current, transientRetries, maxTransientRetries, lastClass)
      ) {
        transientRetries++;
        await this.sleep(backoffMs(transientRetries, response.headers["retry-after"]));
        continue;
      }

      break;
    }

    const totalDuration = attempts.reduce(function (sum, attempt) {
      return sum + attempt.durationMs;
    }, 0);

    return {
      endpointId: spec.endpointId,
      method: spec.intendedMethod,
      url: spec.url,
      status: lastResponse ? lastResponse.status : 0,
      responseClass: lastClass,
      durationMs: totalDuration,
      attempts: attempts,
      responsePreview: lastResponse ? preview(redactText(lastResponse.bodyText), this.options.previewLength) : "",
      responseHeaders: lastResponse ? redactHeaders(lastResponse.headers) : {},
      startedAt: startedAt,
      error: lastError || undefined,
    };
  }

  private async injectDigest(spec: HttpRequestSpec): Promise<HttpRequestSpec> {
    if (!this.digest) {
      return spec;
    }
    const existing = findHeader(spec.headers, "X-RequestDigest");
    if (existing && existing !== DIGEST_PLACEHOLDER && existing !== "*") {
      return spec;
    }
    const digest = await this.digest.get(false);
    return replaceDigestHeader(spec, digest);
  }

  private canRetryTransient(
    spec: HttpRequestSpec,
    used: number,
    limit: number,
    reason: ResponseClass
  ): boolean {
    if (used >= limit) {
      return false;
    }
    if (SAFE_TO_RETRY.indexOf(spec.method) >= 0) {
      return true;
    }
    if (this.options.retryTransientForWrites) {
      return true;
    }
    // Network errors can happen before the server received the request, but we
    // cannot prove that. Only explicitly allowed retries are performed.
    return reason === "network" && !spec.destructive;
  }

  private networkFailure(
    spec: HttpRequestSpec,
    startedAt: string,
    attempts: AttemptRecord[],
    error: unknown
  ): ExecutionRecord {
    return {
      endpointId: spec.endpointId,
      method: spec.intendedMethod,
      url: spec.url,
      status: 0,
      responseClass: "network",
      durationMs: 0,
      attempts: attempts,
      responsePreview: "",
      responseHeaders: {},
      startedAt: startedAt,
      error: redactText(error instanceof Error ? error.message : String(error)),
    };
  }
}

function replaceDigestHeader(spec: HttpRequestSpec, digest: string): HttpRequestSpec {
  const headers: Record<string, string> = {};
  const names = Object.keys(spec.headers);
  for (let i = 0; i < names.length; i++) {
    headers[names[i]] = spec.headers[names[i]];
  }
  headers["X-RequestDigest"] = digest;
  return {
    method: spec.method,
    intendedMethod: spec.intendedMethod,
    url: spec.url,
    headers: headers,
    body: spec.body,
    endpointId: spec.endpointId,
    family: spec.family,
    requiresDigest: spec.requiresDigest,
    destructive: spec.destructive,
    buildNotes: spec.buildNotes,
  };
}

function findHeader(headers: Record<string, string>, name: string): string | undefined {
  const names = Object.keys(headers);
  for (let i = 0; i < names.length; i++) {
    if (names[i].toLowerCase() === name.toLowerCase()) {
      return headers[names[i]];
    }
  }
  return undefined;
}

/** Maps status + body to a category used by tests and evidence triage. */
export function classifyResponse(status: number, body: string): ResponseClass {
  if (status === 0) {
    return "network";
  }
  if (status >= 200 && status < 300) {
    return "success";
  }
  if (status === 400 || status === 422) {
    return "validation";
  }
  if (status === 401) {
    return "authn";
  }
  if (status === 403) {
    return "authz";
  }
  if (status === 404 || status === 410) {
    return "not-found";
  }
  if (status === 409 || status === 412) {
    return "conflict";
  }
  if (status === 429) {
    return "throttled";
  }
  if (status === 502 || status === 503 || status === 504) {
    return "transient";
  }
  if (status >= 500) {
    return "unexpected";
  }
  return "unexpected";
}

function backoffMs(retry: number, retryAfter: string | undefined): number {
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (!isNaN(seconds) && seconds >= 0) {
      return Math.min(seconds * 1000, 30000);
    }
    const date = Date.parse(retryAfter);
    if (!isNaN(date)) {
      return Math.max(0, Math.min(date - Date.now(), 30000));
    }
  }
  return Math.min(500 * Math.pow(2, retry - 1), 8000);
}

function nowMs(): number {
  return typeof performance !== "undefined" && performance && typeof performance.now === "function"
    ? performance.now()
    : Date.now();
}
