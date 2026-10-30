import {
  DigestProvider,
  DigestResponse,
  isDigestRejection,
  parseContextInfoResponse,
} from "../../src/webparts/hacktheworld/research/digest";
import { assertEquals, assertThrows, assertTrue, group, test } from "./harness";

function clock(start: number): { now: () => number; advance: (ms: number) => void } {
  let value = start;
  return {
    now: function () {
      return value;
    },
    advance: function (ms: number) {
      value += ms;
    },
  };
}

export function registerDigestTests(): void {
  group("digest: contextinfo parsing");

  test("parses the verbose d.GetContextWebInformation envelope", function () {
    const parsed = parseContextInfoResponse({
      d: {
        GetContextWebInformation: {
          FormDigestValue: "0xABC",
          FormDigestTimeoutSeconds: 1800,
        },
      },
    });
    assertEquals(parsed.digest, "0xABC");
    assertEquals(parsed.expiresInSeconds, 1800);
  });

  test("parses the nometadata envelope", function () {
    const parsed = parseContextInfoResponse({
      GetContextWebInformation: { FormDigestValue: "0xDEF", FormDigestTimeoutSeconds: 900 },
    });
    assertEquals(parsed.digest, "0xDEF");
    assertEquals(parsed.expiresInSeconds, 900);
  });

  test("parses a response with no envelope at all", function () {
    const parsed = parseContextInfoResponse({ FormDigestValue: "0x123" });
    assertEquals(parsed.digest, "0x123");
    assertEquals(parsed.expiresInSeconds, undefined);
  });

  test("rejects a non-object response", function () {
    assertThrows(function () {
      parseContextInfoResponse("{");
    }, "not a JSON object");
  });

  test("rejects a response without FormDigestValue", function () {
    assertThrows(function () {
      parseContextInfoResponse({ d: { GetContextWebInformation: { FormDigestTimeoutSeconds: 10 } } });
    }, "FormDigestValue");
  });

  group("digest: caching and expiry");

  test("caches a digest until the safety margin expires", async function () {
    const time = clock(0);
    let fetches = 0;
    const provider = new DigestProvider({
      now: time.now,
      fetchDigest: function () {
        fetches++;
        return Promise.resolve({ digest: "digest-" + fetches, expiresInSeconds: 1800 });
      },
    });

    assertEquals(await provider.get(false), "digest-1");
    assertEquals(provider.snapshot().cached, true);
    assertEquals(provider.snapshot().expiresAt, 1680 * 1000);

    time.advance(1679 * 1000);
    assertEquals(await provider.get(false), "digest-1");
    assertEquals(fetches, 1);

    time.advance(2 * 1000);
    assertEquals(provider.snapshot().cached, false);
    assertEquals(await provider.get(false), "digest-2");
    assertEquals(fetches, 2);
  });

  test("uses the default TTL when the server omits it", async function () {
    const time = clock(0);
    const provider = new DigestProvider({
      now: time.now,
      defaultTtlSeconds: 600,
      fetchDigest: function () {
        return Promise.resolve({ digest: "d" });
      },
    });
    await provider.get(false);
    assertEquals(provider.snapshot().expiresAt, (600 - 120) * 1000);
  });

  test("never caches for less than minimumTtlSeconds", async function () {
    const time = clock(0);
    const provider = new DigestProvider({
      now: time.now,
      minimumTtlSeconds: 300,
      fetchDigest: function () {
        return Promise.resolve({ digest: "d", expiresInSeconds: 10 });
      },
    });
    await provider.get(false);
    assertEquals(provider.snapshot().expiresAt, 300 * 1000);
  });

  test("forceRefresh bypasses a valid cache", async function () {
    let fetches = 0;
    const provider = new DigestProvider({
      now: clock(0).now,
      fetchDigest: function () {
        fetches++;
        return Promise.resolve({ digest: "digest-" + fetches, expiresInSeconds: 1800 });
      },
    });
    await provider.get(false);
    const refreshed = await provider.get(true);
    assertEquals(refreshed, "digest-2");
    assertEquals(fetches, 2);
  });

  test("concurrent callers share a single in-flight refresh", async function () {
    let fetches = 0;
    let release: (value: DigestResponse) => void = function () {
      /* replaced below */
    };
    const provider = new DigestProvider({
      now: clock(0).now,
      fetchDigest: function () {
        fetches++;
        return new Promise<DigestResponse>(function (resolve) {
          release = resolve;
        });
      },
    });

    const first = provider.get(false);
    const second = provider.get(false);
    assertEquals(fetches, 1);
    assertTrue(provider.snapshot().pending, "a refresh should be in flight");

    release({ digest: "shared", expiresInSeconds: 1800 });
    assertEquals(await first, "shared");
    assertEquals(await second, "shared");
    assertEquals(fetches, 1);
    assertEquals(provider.snapshot().pending, false);
  });

  test("invalidate drops the cached value", async function () {
    let fetches = 0;
    const provider = new DigestProvider({
      now: clock(0).now,
      fetchDigest: function () {
        fetches++;
        return Promise.resolve({ digest: "digest-" + fetches, expiresInSeconds: 1800 });
      },
    });
    await provider.get(false);
    provider.invalidate();
    assertEquals(provider.snapshot().cached, false);
    assertEquals(provider.snapshot().expiresAt, null);
    assertEquals(await provider.get(false), "digest-2");
  });

  test("invalidate while a refresh is in flight never caches the stale value", async function () {
    let fetches = 0;
    let release: (value: DigestResponse) => void = function () {
      /* replaced below */
    };
    const provider = new DigestProvider({
      now: clock(0).now,
      fetchDigest: function () {
        fetches++;
        return new Promise<DigestResponse>(function (resolve) {
          release = resolve;
        });
      },
    });

    const pending = provider.get(false);
    provider.invalidate();
    release({ digest: "stale", expiresInSeconds: 1800 });
    assertEquals(await pending, "stale");
    assertEquals(provider.snapshot().cached, false);

    const next = provider.get(false);
    assertEquals(fetches, 2, "a new refresh must start after invalidation");
    release({ digest: "fresh", expiresInSeconds: 1800 });
    assertEquals(await next, "fresh");
    assertEquals(provider.snapshot().cached, true);
  });

  test("a failed refresh clears pending so a later call can succeed", async function () {
    let attempt = 0;
    const provider = new DigestProvider({
      now: clock(0).now,
      fetchDigest: function () {
        attempt++;
        return attempt === 1
          ? Promise.reject(new Error("contextinfo down"))
          : Promise.resolve({ digest: "recovered", expiresInSeconds: 1800 });
      },
    });

    let firstError = "";
    try {
      await provider.get(false);
    } catch (error) {
      firstError = error instanceof Error ? error.message : String(error);
    }
    assertTrue(firstError.indexOf("contextinfo down") >= 0, "the fetch error must propagate");
    assertEquals(provider.snapshot().pending, false);
    assertEquals(await provider.get(false), "recovered");
  });

  test("a response without a digest is rejected", async function () {
    const provider = new DigestProvider({
      now: clock(0).now,
      fetchDigest: function () {
        return Promise.resolve({ digest: "" });
      },
    });
    let message = "";
    try {
      await provider.get(false);
    } catch (error) {
      message = error instanceof Error ? error.message : String(error);
    }
    assertTrue(message.indexOf("FormDigestValue") >= 0, "the error must explain the missing digest");
    assertEquals(provider.snapshot().cached, false);
  });

  test("the diagnostics snapshot never contains the digest", async function () {
    const provider = new DigestProvider({
      now: clock(0).now,
      fetchDigest: function () {
        return Promise.resolve({ digest: "super-secret-digest-value", expiresInSeconds: 1800 });
      },
    });
    await provider.get(false);
    const serialised = JSON.stringify(provider.snapshot());
    assertTrue(serialised.indexOf("super-secret-digest-value") < 0, "snapshot leaked the digest");
    assertEquals(provider.snapshot().refreshes, 1);
  });

  test("the constructor requires a fetch function", function () {
    assertThrows(function () {
      new DigestProvider({} as never);
    }, "fetchDigest");
  });

  group("digest: rejection detection");

  test("recognises the classic security-validation text", function () {
    assertTrue(isDigestRejection(403, "The security validation for this page is invalid."), "text not detected");
  });

  test("recognises the error code and header name", function () {
    assertTrue(isDigestRejection(400, "error -2130575251"), "error code not detected");
    assertTrue(isDigestRejection(401, "missing X-RequestDigest"), "header name not detected");
    assertTrue(isDigestRejection(403, "Invalid form digest"), "form digest text not detected");
  });

  test("does not treat unrelated failures as digest rejections", function () {
    assertTrue(!isDigestRejection(403, "Access denied."), "plain authorization failure is not a digest issue");
    assertTrue(!isDigestRejection(500, "security validation"), "500 is not a digest rejection");
    assertTrue(!isDigestRejection(403, ""), "an empty body cannot prove a digest rejection");
  });

  test("a valid request duration is irrelevant to rejection detection", function () {
    assertTrue(!isDigestRejection(200, "security validation"), "success is never a rejection");
  });
}
