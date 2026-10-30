import { DigestProvider } from "../../src/webparts/hacktheworld/research/digest";
import { buildRequest } from "../../src/webparts/hacktheworld/research/request";
import {
  classifyResponse,
  RequestExecutor,
} from "../../src/webparts/hacktheworld/research/executor";
import {
  HttpRequestSpec,
  ResponseClass,
  TransportResponse,
} from "../../src/webparts/hacktheworld/research/types";
import { assertEquals, assertTrue, group, test } from "./harness";

const SITE = "https://contoso.sharepoint.com/sites/team";
const LIST_ID = "4a81ba4f-84ed-4e91-a731-9a82c0ede37e";

interface RecordedCall {
  spec: HttpRequestSpec;
}

function response(status: number, bodyText: string, headers?: Record<string, string>): TransportResponse {
  return { status: status, bodyText: bodyText, headers: headers || {} };
}

/** A transport that replays a fixed list of outcomes and records every call. */
function scriptedTransport(outcomes: Array<TransportResponse | Error>): {
  transport: (spec: HttpRequestSpec) => Promise<TransportResponse>;
  calls: RecordedCall[];
} {
  const calls: RecordedCall[] = [];
  let index = 0;
  return {
    calls: calls,
    transport: function (spec: HttpRequestSpec): Promise<TransportResponse> {
      calls.push({ spec: spec });
      const outcome = outcomes[Math.min(index, outcomes.length - 1)];
      index++;
      if (outcome instanceof Error) {
        return Promise.reject(outcome);
      }
      return Promise.resolve(outcome);
    },
  };
}

function noSleep(): { sleep: (ms: number) => Promise<void>; delays: number[] } {
  const delays: number[] = [];
  return {
    delays: delays,
    sleep: function (ms: number): Promise<void> {
      delays.push(ms);
      return Promise.resolve();
    },
  };
}

export function registerExecutorTests(): void {
  group("executor: response classification");

  const CASES: Array<[number, ResponseClass]> = [
    [0, "network"],
    [200, "success"],
    [204, "success"],
    [400, "validation"],
    [422, "validation"],
    [401, "authn"],
    [403, "authz"],
    [404, "not-found"],
    [410, "not-found"],
    [409, "conflict"],
    [412, "conflict"],
    [429, "throttled"],
    [502, "transient"],
    [503, "transient"],
    [504, "transient"],
    [500, "unexpected"],
    [418, "unexpected"],
  ];
  CASES.forEach(function (entry) {
    test("status " + entry[0] + " classifies as " + entry[1], function () {
      assertEquals(classifyResponse(entry[0], ""), entry[1]);
    });
  });

  group("executor: transient retries");

  test("GET retries a 503 and returns the eventual success", async function () {
    const script = scriptedTransport([response(503, "busy"), response(200, "{\"ok\":true}")]);
    const timer = noSleep();
    const executor = new RequestExecutor(script.transport, null, { sleep: timer.sleep });
    const record = await executor.execute(buildRequest("lists.collection.get", { siteUrl: SITE }));
    assertEquals(record.status, 200);
    assertEquals(record.responseClass, "success");
    assertEquals(record.attempts.length, 2);
    assertEquals(script.calls.length, 2);
    assertEquals(record.attempts[0].status, 503);
  });

  test("retry-after in seconds bounds the backoff delay", async function () {
    const script = scriptedTransport([response(429, "slow down", { "retry-after": "2" }), response(200, "")]);
    const timer = noSleep();
    const executor = new RequestExecutor(script.transport, null, { sleep: timer.sleep });
    const record = await executor.execute(buildRequest("lists.collection.get", { siteUrl: SITE }));
    assertEquals(record.status, 200);
    assertEquals(timer.delays[0], 2000);
  });

  test("retries are bounded by maxTransientRetries", async function () {
    const script = scriptedTransport([response(503, "busy")]);
    const timer = noSleep();
    const executor = new RequestExecutor(script.transport, null, {
      sleep: timer.sleep,
      maxTransientRetries: 2,
    });
    const record = await executor.execute(buildRequest("lists.collection.get", { siteUrl: SITE }));
    assertEquals(record.status, 503);
    assertEquals(record.attempts.length, 3);
    assertEquals(timer.delays.length, 2);
  });

  test("DELETE is never retried on a transient failure", async function () {
    const script = scriptedTransport([response(503, "busy")]);
    const executor = new RequestExecutor(script.transport, null, { sleep: noSleep().sleep });
    const record = await executor.execute(
      buildRequest("items.delete", {
        siteUrl: SITE,
        path: { listId: LIST_ID, itemId: 2 },
        digest: "0xABC",
      })
    );
    assertEquals(record.status, 503);
    assertEquals(record.attempts.length, 1);
    assertEquals(script.calls.length, 1);
  });

  test("writes are retried only when explicitly opted in", async function () {
    const script = scriptedTransport([response(503, "busy"), response(204, "")]);
    const executor = new RequestExecutor(script.transport, null, {
      sleep: noSleep().sleep,
      retryTransientForWrites: true,
    });
    const record = await executor.execute(
      buildRequest("items.update", {
        siteUrl: SITE,
        path: { listId: LIST_ID, itemId: 2 },
        body: { Title: "probe" },
        digest: "0xABC",
      })
    );
    assertEquals(record.status, 204);
    assertEquals(record.attempts.length, 2);
  });

  group("executor: network failures");

  test("a network error on a safe read is retried", async function () {
    const script = scriptedTransport([new Error("ECONNRESET"), response(200, "ok")]);
    const executor = new RequestExecutor(script.transport, null, { sleep: noSleep().sleep });
    const record = await executor.execute(buildRequest("lists.collection.get", { siteUrl: SITE }));
    assertEquals(record.status, 200);
    assertEquals(record.attempts.length, 2);
    assertEquals(record.attempts[0].responseClass, "network");
  });

  test("a network error on a destructive write is not retried", async function () {
    const script = scriptedTransport([new Error("ECONNRESET")]);
    const executor = new RequestExecutor(script.transport, null, { sleep: noSleep().sleep });
    const record = await executor.execute(
      buildRequest("items.delete", {
        siteUrl: SITE,
        path: { listId: LIST_ID, itemId: 2 },
        digest: "0xABC",
      })
    );
    assertEquals(record.status, 0);
    assertEquals(record.responseClass, "network");
    assertEquals(record.attempts.length, 1);
    assertTrue(!!record.error && record.error.indexOf("ECONNRESET") >= 0, "error must be preserved");
  });

  test("an exhausted network failure is reported, not thrown", async function () {
    const script = scriptedTransport([new Error("ENOTFOUND")]);
    const executor = new RequestExecutor(script.transport, null, {
      sleep: noSleep().sleep,
      maxTransientRetries: 1,
    });
    const record = await executor.execute(buildRequest("lists.collection.get", { siteUrl: SITE }));
    assertEquals(record.status, 0);
    assertEquals(record.responseClass, "network");
    assertEquals(record.attempts.length, 2);
    assertTrue(!!record.error, "expected an error message");
  });

  group("executor: digest handling");

  test("a digest rejection refreshes the digest and retries exactly once", async function () {
    let fetches = 0;
    const digests = new DigestProvider({
      fetchDigest: function () {
        fetches++;
        return Promise.resolve({ digest: "digest-" + fetches, expiresInSeconds: 3600 });
      },
    });
    const script = scriptedTransport([
      response(403, "The security validation for this page is invalid."),
      response(200, "{\"ok\":true}"),
    ]);
    const executor = new RequestExecutor(script.transport, digests, { sleep: noSleep().sleep });
    const record = await executor.execute(
      buildRequest("items.update", {
        siteUrl: SITE,
        path: { listId: LIST_ID, itemId: 2 },
        body: { Title: "probe" },
      })
    );
    assertEquals(fetches, 2);
    assertEquals(record.status, 200);
    assertEquals(record.attempts.length, 2);
    assertEquals(record.attempts[0].digestRefreshed, true);
    assertEquals(script.calls[0].spec.headers["X-RequestDigest"], "digest-1");
    assertEquals(script.calls[1].spec.headers["X-RequestDigest"], "digest-2");
  });

  test("a digest rejection without a provider is surfaced as an authz response", async function () {
    const script = scriptedTransport([response(403, "security validation failed")]);
    const executor = new RequestExecutor(script.transport, null, { sleep: noSleep().sleep });
    const record = await executor.execute(
      buildRequest("items.update", {
        siteUrl: SITE,
        path: { listId: LIST_ID, itemId: 2 },
        body: { Title: "probe" },
      })
    );
    assertEquals(record.status, 403);
    assertEquals(record.responseClass, "authz");
    assertEquals(record.attempts.length, 1);
  });

  test("maxDigestRetries=0 disables the refresh even with a provider", async function () {
    const digests = new DigestProvider({
      fetchDigest: function () {
        return Promise.resolve({ digest: "digest", expiresInSeconds: 3600 });
      },
    });
    const script = scriptedTransport([response(403, "security validation failed")]);
    const executor = new RequestExecutor(script.transport, digests, { sleep: noSleep().sleep });
    const record = await executor.execute(
      buildRequest("items.update", {
        siteUrl: SITE,
        path: { listId: LIST_ID, itemId: 2 },
        body: { Title: "probe" },
      }),
      { maxDigestRetries: 0 }
    );
    assertEquals(record.status, 403);
    assertEquals(record.attempts.length, 1);
    assertEquals(record.attempts[0].digestRefreshed, false);
  });

  test("a placeholder digest is replaced before the first attempt", async function () {
    const digests = new DigestProvider({
      fetchDigest: function () {
        return Promise.resolve({ digest: "real-digest", expiresInSeconds: 3600 });
      },
    });
    const script = scriptedTransport([response(204, "")]);
    const executor = new RequestExecutor(script.transport, digests, { sleep: noSleep().sleep });
    await executor.execute(
      buildRequest("items.delete", {
        siteUrl: SITE,
        path: { listId: LIST_ID, itemId: 2 },
      })
    );
    assertEquals(script.calls[0].spec.headers["X-RequestDigest"], "real-digest");
  });

  test("skipDigestInjection leaves the placeholder intact", async function () {
    const digests = new DigestProvider({
      fetchDigest: function () {
        return Promise.resolve({ digest: "real-digest", expiresInSeconds: 3600 });
      },
    });
    const script = scriptedTransport([response(403, "security validation failed")]);
    const executor = new RequestExecutor(script.transport, digests, { sleep: noSleep().sleep });
    await executor.execute(
      buildRequest("items.delete", {
        siteUrl: SITE,
        path: { listId: LIST_ID, itemId: 2 },
      }),
      { skipDigestInjection: true }
    );
    assertEquals(script.calls[0].spec.headers["X-RequestDigest"], "{{REQUEST_DIGEST}}");
  });

  test("an intentionally invalid digest is not overwritten", async function () {
    const digests = new DigestProvider({
      fetchDigest: function () {
        return Promise.resolve({ digest: "real-digest", expiresInSeconds: 3600 });
      },
    });
    const script = scriptedTransport([response(403, "security validation failed")]);
    const executor = new RequestExecutor(script.transport, digests, { sleep: noSleep().sleep });
    await executor.execute(
      buildRequest("items.delete", {
        siteUrl: SITE,
        path: { listId: LIST_ID, itemId: 2 },
        digest: "intentionally-invalid",
      })
    );
    assertEquals(script.calls[0].spec.headers["X-RequestDigest"], "intentionally-invalid");
  });

  test("a digest fetch failure is reported as a network failure", async function () {
    const digests = new DigestProvider({
      fetchDigest: function () {
        return Promise.reject(new Error("contextinfo unavailable"));
      },
    });
    const script = scriptedTransport([response(204, "")]);
    const executor = new RequestExecutor(script.transport, digests, { sleep: noSleep().sleep });
    const record = await executor.execute(
      buildRequest("items.delete", {
        siteUrl: SITE,
        path: { listId: LIST_ID, itemId: 2 },
      })
    );
    assertEquals(record.status, 0);
    assertEquals(record.responseClass, "network");
    assertEquals(script.calls.length, 0);
    assertTrue(!!record.error && record.error.indexOf("contextinfo") >= 0, "error must explain the digest failure");
  });

  group("executor: evidence capture");

  test("the response preview redacts credentials and truncates", async function () {
    const body = "{\"access_token\":\"SECRET-TOKEN\",\"Authorization\":\"Bearer abc.def.ghi\"}";
    const script = scriptedTransport([
      response(200, body, { Authorization: "Bearer abc.def.ghi", "X-Trace": "keep-me" }),
    ]);
    const executor = new RequestExecutor(script.transport, null, {
      sleep: noSleep().sleep,
      previewLength: 500,
    });
    const record = await executor.execute(buildRequest("lists.collection.get", { siteUrl: SITE }));
    assertTrue(record.responsePreview.indexOf("SECRET-TOKEN") < 0, "token must not be stored in the preview");
    assertTrue(record.responsePreview.indexOf("abc.def.ghi") < 0, "bearer token must not be stored");
    assertEquals(record.responseHeaders["Authorization"], "[REDACTED]");
    assertEquals(record.responseHeaders["X-Trace"], "keep-me");
  });

  test("the preview is cut at previewLength", async function () {
    const long = "x".repeat(1000);
    const script = scriptedTransport([response(200, long)]);
    const executor = new RequestExecutor(script.transport, null, {
      sleep: noSleep().sleep,
      previewLength: 50,
    });
    const record = await executor.execute(buildRequest("lists.collection.get", { siteUrl: SITE }));
    assertTrue(record.responsePreview.indexOf("chars truncated") >= 0, "expected a truncation marker");
    assertTrue(record.responsePreview.length < 200, "preview must stay small");
  });

  test("the execution record keeps the endpoint id, intended method and URL", async function () {
    const script = scriptedTransport([response(200, "")]);
    const executor = new RequestExecutor(script.transport, null, { sleep: noSleep().sleep });
    const spec = buildRequest("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
      body: { Title: "probe" },
      method: "MERGE",
      digest: "0xABC",
    });
    const record = await executor.execute(spec);
    assertEquals(record.endpointId, "items.update");
    assertEquals(record.method, "MERGE");
    assertEquals(record.url, spec.url);
    assertEquals(record.startedAt.length > 0, true);
  });

  test("the constructor rejects a non-function transport", function () {
    let threw = false;
    try {
      new RequestExecutor(null as never);
    } catch (error) {
      threw = true;
    }
    assertTrue(threw, "RequestExecutor must require a transport function");
  });
}
