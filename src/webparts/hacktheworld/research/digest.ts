/**
 * Centralised SharePoint request-digest handling.
 *
 * Design constraints:
 *  - one in-flight refresh at a time (concurrent writes share one contextinfo call)
 *  - the digest is cached until shortly before FormDigestTimeoutSeconds
 *  - a rejected digest invalidates the cache; callers decide how many times to retry
 *  - the digest value is never logged or exported
 */

export interface DigestResponse {
  digest: string;
  /** FormDigestTimeoutSeconds from contextinfo, when available. */
  expiresInSeconds?: number;
}

export type DigestFetcher = () => Promise<DigestResponse>;

export interface DigestOptions {
  fetchDigest: DigestFetcher;
  now?: () => number;
  /** Fallback TTL when the server does not return one. Default 1800s. */
  defaultTtlSeconds?: number;
  /** Refresh this many seconds before the server-side expiry. Default 120s. */
  safetyMarginSeconds?: number;
  /** Minimum cache time to avoid a refresh storm with tiny TTLs. Default 30s. */
  minimumTtlSeconds?: number;
}

export interface DigestSnapshot {
  cached: boolean;
  expiresAt: number | null;
  refreshes: number;
  pending: boolean;
}

export class DigestProvider {
  private readonly options: DigestOptions;
  private readonly now: () => number;
  private value: string | null = null;
  private expiresAt = 0;
  private pending: Promise<string> | null = null;
  private pendingId = 0;
  private refreshes = 0;

  constructor(options: DigestOptions) {
    if (!options || typeof options.fetchDigest !== "function") {
      throw new Error("DigestProvider requires a fetchDigest function.");
    }
    this.options = options;
    this.now = options.now || function () {
      return Date.now();
    };
  }

  /** Returns a digest, refreshing it when expired or when forceRefresh is true. */
  public get(forceRefresh?: boolean): Promise<string> {
    const now = this.now();
    if (!forceRefresh && this.value && now < this.expiresAt) {
      return Promise.resolve(this.value);
    }
    if (this.pending) {
      return this.pending;
    }

    const requestId = ++this.pendingId;
    this.refreshes++;
    const promise = this.options.fetchDigest().then(
      (response: DigestResponse): string => {
        if (!response || !response.digest) {
          throw new Error("contextinfo did not return a FormDigestValue.");
        }
        // If invalidate() ran while this request was in flight, never cache the
        // stale value; the caller that forced the refresh will start a new one.
        if (requestId === this.pendingId) {
          const ttlSeconds =
            response.expiresInSeconds !== undefined && response.expiresInSeconds > 0
              ? response.expiresInSeconds
              : this.options.defaultTtlSeconds !== undefined
              ? this.options.defaultTtlSeconds
              : 1800;
          const margin =
            this.options.safetyMarginSeconds !== undefined ? this.options.safetyMarginSeconds : 120;
          const minimum =
            this.options.minimumTtlSeconds !== undefined ? this.options.minimumTtlSeconds : 30;
          const effectiveSeconds = Math.max(ttlSeconds - margin, minimum);
          this.value = response.digest;
          this.expiresAt = this.now() + effectiveSeconds * 1000;
        }
        return response.digest;
      }
    );

    const tracked = promise.then(
      (digest: string): string => {
        if (requestId === this.pendingId) {
          this.pending = null;
        }
        return digest;
      },
      (error: unknown): never => {
        if (requestId === this.pendingId) {
          this.pending = null;
        }
        throw error;
      }
    );

    this.pending = tracked;
    return tracked;
  }

  /** Drops the cached value and abandons any in-flight refresh result. */
  public invalidate(): void {
    this.value = null;
    this.expiresAt = 0;
    this.pendingId++;
    this.pending = null;
  }

  /** Safe diagnostics: never includes the digest itself. */
  public snapshot(): DigestSnapshot {
    return {
      cached: !!this.value && this.now() < this.expiresAt,
      expiresAt: this.value ? this.expiresAt : null,
      refreshes: this.refreshes,
      pending: !!this.pending,
    };
  }
}

/**
 * Parses both verbose (d.GetContextWebInformation) and nometadata
 * (GetContextWebInformation) contextinfo responses.
 */
export function parseContextInfoResponse(json: unknown): DigestResponse {
  if (!json || typeof json !== "object") {
    throw new Error("contextinfo response is not a JSON object.");
  }
  const root = json as Record<string, unknown>;
  const container =
    root.d && typeof root.d === "object"
      ? (root.d as Record<string, unknown>)
      : root;
  const info =
    container.GetContextWebInformation && typeof container.GetContextWebInformation === "object"
      ? (container.GetContextWebInformation as Record<string, unknown>)
      : container;

  const value = info.FormDigestValue;
  if (typeof value !== "string" || value.length === 0) {
    throw new Error("contextinfo response is missing FormDigestValue.");
  }
  const timeout = info.FormDigestTimeoutSeconds;
  return {
    digest: value,
    expiresInSeconds: typeof timeout === "number" ? timeout : undefined,
  };
}

/**
 * Detects SharePoint digest rejections so they are not mistaken for real
 * authorization failures. Both the classic and modern error texts are covered.
 */
export function isDigestRejection(status: number, body: string): boolean {
  if (status !== 403 && status !== 400 && status !== 401) {
    return false;
  }
  if (!body) {
    return false;
  }
  const text = body.toLowerCase();
  return (
    text.indexOf("security validation") >= 0 ||
    text.indexOf("-2130575251") >= 0 ||
    text.indexOf("form digest") >= 0 ||
    text.indexOf("x-requestdigest") >= 0 ||
    text.indexOf("requestdigest") >= 0
  );
}
