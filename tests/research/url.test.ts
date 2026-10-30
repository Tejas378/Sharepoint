import {
  encodePathSegment,
  encodeQueryName,
  encodeQueryValue,
  getOrigin,
  getPathAndQuery,
  isCompactGuid,
  isGuid,
  joinUrl,
  odataGuid,
  odataString,
  resolveTemplate,
  serializeQuery,
  toQueryPairs,
  trimTrailingSlash,
} from "../../src/webparts/hacktheworld/research/url";
import { assertEquals, assertThrows, assertTrue, group, test } from "./harness";

export function registerUrlTests(): void {
  group("url: joining and origin handling");

  test("trimTrailingSlash removes a single trailing slash only", function () {
    assertEquals(trimTrailingSlash("https://contoso.sharepoint.com/"), "https://contoso.sharepoint.com");
    assertEquals(trimTrailingSlash("https://contoso.sharepoint.com"), "https://contoso.sharepoint.com");
    assertEquals(trimTrailingSlash(""), "");
  });

  test("joinUrl joins site and path without double slashes", function () {
    assertEquals(
      joinUrl("https://contoso.sharepoint.com/sites/team/", "/_api/web"),
      "https://contoso.sharepoint.com/sites/team/_api/web"
    );
    assertEquals(joinUrl("", "/_api/web"), "/_api/web");
    assertEquals(joinUrl("https://contoso.sharepoint.com", ""), "https://contoso.sharepoint.com");
  });

  test("joinUrl keeps absolute paths absolute", function () {
    assertEquals(joinUrl("https://a.sharepoint.com", "https://b.sharepoint.com/x"), "https://b.sharepoint.com/x");
  });

  test("getOrigin and getPathAndQuery split a request line", function () {
    assertEquals(getOrigin("https://contoso.sharepoint.com/sites/team/_api/web?$top=1"), "https://contoso.sharepoint.com");
    assertEquals(
      getPathAndQuery("https://contoso.sharepoint.com/sites/team/_api/web?$top=1"),
      "/sites/team/_api/web?$top=1"
    );
    assertEquals(getPathAndQuery("/_api/web"), "/_api/web");
  });

  group("url: OData literals");

  test("odataString escapes embedded single quotes", function () {
    assertEquals(odataString("O'Brien"), "'O''Brien'");
    assertEquals(odataString("plain"), "'plain'");
  });

  test("odataGuid normalises braces and case", function () {
    assertEquals(
      odataGuid("{4A81BA4F-84ED-4E91-A731-9A82C0EDE37E}"),
      "guid'4a81ba4f-84ed-4e91-a731-9a82c0ede37e'"
    );
  });

  test("odataGuid rejects non-GUID input instead of building a broken URL", function () {
    assertThrows(function () {
      odataGuid("not-a-guid");
    }, "Expected a GUID");
  });

  test("GUID helpers accept and reject the documented shapes", function () {
    assertTrue(isGuid("4a81ba4f-84ed-4e91-a731-9a82c0ede37e"), "dashed GUID should be valid");
    assertTrue(!isGuid("4a81ba4f84ed4e91a7319a82c0ede37e"), "compact GUID is not a dashed GUID");
    assertTrue(isCompactGuid("4a81ba4f84ed4e91a7319a82c0ede37e"), "compact GUID should be valid");
  });

  group("url: query serialisation");

  test("encodeQueryValue keeps OData punctuation readable", function () {
    assertEquals(encodeQueryValue("a b'c,d(e):f"), "a%20b'c,d(e):f");
    assertEquals(encodeQueryValue("<View/>"), "%3CView%2F%3E");
  });

  test("encodeQueryName keeps $ and @ readable", function () {
    assertEquals(encodeQueryName("$select"), "$select");
    assertEquals(encodeQueryName("@viewXml"), "@viewXml");
  });

  test("encodePathSegment encodes slashes inside a name", function () {
    assertEquals(encodePathSegment("a/b.txt"), "a%2Fb.txt");
  });

  test("serializeQuery joins list options with commas and repeats others", function () {
    assertEquals(serializeQuery([{ name: "$select", value: ["Id", "Title"] }]), "$select=Id,Title");
    assertEquals(serializeQuery([{ name: "p", value: [1, 2] }]), "p=1&p=2");
    assertEquals(serializeQuery([{ name: "$top", value: 10 }]), "$top=10");
  });

  test("serializeQuery can emit raw values for malformed-input reproduction", function () {
    assertEquals(
      serializeQuery([{ name: "$filter", value: "Title eq 'x'", raw: true }]),
      "$filter=Title eq 'x'"
    );
  });

  test("toQueryPairs preserves insertion order", function () {
    const pairs = toQueryPairs({ b: 1, a: 2 });
    assertEquals(pairs[0].name, "b");
    assertEquals(pairs[1].name, "a");
  });

  group("url: template resolution");

  test("resolveTemplate applies guid, odata, path and int modifiers", function () {
    assertEquals(
      resolveTemplate("/_api/web/lists({listId:guid})", { listId: "4A81BA4F-84ED-4E91-A731-9A82C0EDE37E" }, "t")
        .path,
      "/_api/web/lists(guid'4a81ba4f-84ed-4e91-a731-9a82c0ede37e')"
    );
    assertEquals(resolveTemplate("/x('{t:odata}')", { t: "O'Brien" }, "t").path, "/x(''O''Brien'')");
    assertEquals(resolveTemplate("/x/{p:path}", { p: "a b" }, "t").path, "/x/a%20b");
    assertEquals(resolveTemplate("/x({i:int})", { i: 42 }, "t").path, "/x(42)");
  });

  test("resolveTemplate reports the missing placeholder and endpoint id", function () {
    assertThrows(function () {
      resolveTemplate("/_api/web/lists({listId:guid})", {}, "lists.getById");
    }, "Missing path placeholder \\{listId\\}.*lists.getById");
  });

  test("resolveTemplate rejects a non-integer where :int is declared", function () {
    assertThrows(function () {
      resolveTemplate("/x({itemId:int})", { itemId: "1 or 1=1" }, "items.getById");
    }, "expected an integer");
  });

  test("resolveTemplate rejects unknown modifiers", function () {
    assertThrows(function () {
      resolveTemplate("/x({v:bogus})", { v: "1" }, "t");
    }, "Unknown placeholder modifier");
  });

  test("resolveTemplate only reports placeholders that were actually used", function () {
    const result = resolveTemplate("/x/{a}", { a: "1", b: "2" }, "t");
    assertEquals(result.used.a, true);
    assertEquals(result.used.b, undefined);
  });
}
