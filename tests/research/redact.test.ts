import {
  isSecretKey,
  preview,
  REDACTED,
  redactHeaders,
  redactJson,
  redactText,
  redactUrl,
} from "../../src/webparts/hacktheworld/research/redact";
import { assertEquals, assertTrue, group, test } from "./harness";

export function registerRedactTests(): void {
  group("redact: key classification");

  ["password", "passwd", "pwd", "client_secret", "clientSecret", "access_token", "refresh_token", "bearer", "apiKey", "api_key", "authorization", "cookie", "digest", "saml", "assertion", "signature", "sig"].forEach(
    function (key) {
      test("'" + key + "' is treated as secret", function () {
        assertTrue(isSecretKey(key), key + " should be redacted");
      });
    }
  );

  ["$select", "$filter", "$skiptoken", "skiptoken", "ChangeToken", "viewToken", "FormDigestTimeoutSeconds", "Title", "Id", "resource"].forEach(
    function (key) {
      test("'" + key + "' is not treated as secret", function () {
        assertTrue(!isSecretKey(key), key + " must stay readable");
      });
    }
  );

  test("empty keys are not secret", function () {
    assertTrue(!isSecretKey(""), "empty key must not trigger redaction");
  });

  group("redact: headers");

  test("credential headers are replaced but ordinary headers survive", function () {
    const redacted = redactHeaders({
      Authorization: "Bearer token",
      Cookie: "FedAuth=abc",
      "Set-Cookie": "FedAuth=abc",
      "X-RequestDigest": "0x123",
      "X-Api-Key": "key",
      "Content-Type": "application/json",
      ETag: '"1"',
    });
    assertEquals(redacted["Authorization"], REDACTED);
    assertEquals(redacted["Cookie"], REDACTED);
    assertEquals(redacted["Set-Cookie"], REDACTED);
    assertEquals(redacted["X-RequestDigest"], REDACTED);
    assertEquals(redacted["X-Api-Key"], REDACTED);
    assertEquals(redacted["Content-Type"], "application/json");
    assertEquals(redacted["ETag"], '"1"');
  });

  test("header matching is case-insensitive", function () {
    const redacted = redactHeaders({ aUtHoRiZaTiOn: "secret" });
    assertEquals(redacted["aUtHoRiZaTiOn"], REDACTED);
  });

  test("the input object is not mutated", function () {
    const input = { Authorization: "secret" };
    redactHeaders(input);
    assertEquals(input["Authorization"], "secret");
  });

  group("redact: URLs");

  test("credential query values are redacted", function () {
    const url = redactUrl(
      "https://contoso.sharepoint.com/_api/web?access_token=abc&password=p%40ss&$select=Id&$skiptoken=Paged%3DTRUE"
    );
    assertTrue(url.indexOf("abc") < 0, "access token must be removed");
    assertTrue(url.indexOf("p%40ss") < 0, "password must be removed");
    assertTrue(url.indexOf("$select=Id") >= 0, "$select must stay readable");
    assertTrue(url.indexOf("$skiptoken=Paged%3DTRUE") >= 0, "$skiptoken must stay readable");
  });

  test("the fragment survives redaction", function () {
    const url = redactUrl("https://contoso.sharepoint.com/x?password=abc#frag");
    assertEquals(url, "https://contoso.sharepoint.com/x?password=" + REDACTED + "#frag");
  });

  test("a URL without a query is returned unchanged", function () {
    const url = "https://contoso.sharepoint.com/sites/team";
    assertEquals(redactUrl(url), url);
  });

  test("a malformed percent-encoded name does not throw", function () {
    const url = redactUrl("https://contoso.sharepoint.com/x?%E0%A4%A=1");
    assertTrue(url.length > 0, "malformed input must still produce a value");
  });

  test("parameters without a value are preserved", function () {
    assertEquals(redactUrl("https://x/y?flag&$top=1"), "https://x/y?flag&$top=1");
  });

  group("redact: JSON bodies");

  test("nested credential keys are redacted recursively", function () {
    const redacted = redactJson({
      Title: "probe",
      nested: { client_secret: "abc", list: [{ access_token: "def" }, { Id: 1 }] },
    }) as Record<string, any>;
    assertEquals(redacted["Title"], "probe");
    assertEquals(redacted["nested"]["client_secret"], REDACTED);
    assertEquals(redacted["nested"]["list"][0]["access_token"], REDACTED);
    assertEquals(redacted["nested"]["list"][1]["Id"], 1);
  });

  test("primitive and null values pass through", function () {
    assertEquals(redactJson(null), null);
    assertEquals(redactJson(42), 42);
    assertEquals(redactJson("text"), "text");
    assertEquals(redactJson([1, 2]), [1, 2]);
  });

  test("depth is bounded so hostile nesting cannot recurse forever", function () {
    let deep: Record<string, unknown> = { password: "top-secret" };
    for (let i = 0; i < 20; i++) {
      deep = { child: deep };
    }
    const serialised = JSON.stringify(redactJson(deep));
    assertTrue(serialised.length > 0, "deep structure must still serialise");
    assertTrue(serialised.indexOf("top-secret") >= 0, "beyond the depth cap values are returned unchanged by design");
  });

  group("redact: free text");

  test("bearer tokens are removed from text", function () {
    const text = redactText("Authorization: Bearer abc.def.ghi");
    assertTrue(text.indexOf("abc.def.ghi") < 0, "bearer token must be removed");
    assertTrue(text.indexOf(REDACTED) >= 0, "expected a redaction marker");
  });

  test("credential-looking JSON fields are removed from text", function () {
    const text = redactText('{"password":"hunter2","access_token":"abc"}');
    assertTrue(text.indexOf("hunter2") < 0, "password must be removed");
    assertTrue(text.indexOf('"abc"') < 0, "access token must be removed");
  });

  test("a request digest header in text is removed", function () {
    const text = redactText("X-RequestDigest: 0xABC123 trailing");
    assertTrue(text.indexOf("0xABC123") < 0, "digest must be removed");
  });

  test("empty text is returned as-is", function () {
    assertEquals(redactText(""), "");
  });

  group("redact: preview truncation");

  test("an empty body previews as an empty string", function () {
    assertEquals(preview(""), "");
  });

  test("a body at or below the limit is returned unchanged", function () {
    assertEquals(preview("abc", 3), "abc");
    assertEquals(preview("abcdef", 6), "abcdef");
  });

  test("a longer body is truncated with a marker and the omitted size", function () {
    const text = preview("abcdefghij", 4);
    assertEquals(text.indexOf("abcd"), 0);
    assertTrue(text.indexOf("6 chars truncated") >= 0, "expected the omitted character count");
  });

  test("the default preview length is 2000 characters", function () {
    const text = preview("x".repeat(2500));
    assertTrue(text.length < 2100, "default preview must stay near 2000 characters");
    assertTrue(text.indexOf("500 chars truncated") >= 0, "expected 500 omitted characters");
  });
}
