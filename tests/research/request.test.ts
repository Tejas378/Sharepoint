import { buildListItemRequest, buildRequest, DIGEST_PLACEHOLDER } from "../../src/webparts/hacktheworld/research/request";
import { assertEquals, assertThrows, assertTrue, assertContains, group, test } from "./harness";

const SITE = "https://contoso.sharepoint.com/sites/team";
const LIST_ID = "4a81ba4f-84ed-4e91-a731-9a82c0ede37e";

export function registerRequestTests(): void {
  group("request: URL construction");

  test("builds an absolute URL by joining siteUrl and template", function () {
    const spec = buildRequest("lists.collection.get", { siteUrl: SITE });
    assertEquals(spec.url, SITE + "/_api/web/lists");
    assertEquals(spec.method, "GET");
    assertEquals(spec.intendedMethod, "GET");
  });

  test("substitutes and validates path placeholders", function () {
    const spec = buildListItemRequest("items.getById", SITE, LIST_ID, 7);
    assertEquals(spec.url, SITE + "/_api/web/lists(guid'" + LIST_ID + "')/items(7)");
  });

  test("rejects a non-integer item id for an :int placeholder", function () {
    assertThrows(function () {
      buildRequest("items.getById", { siteUrl: SITE, path: { listId: LIST_ID, itemId: "1 or 1=1" } });
    }, "expected an integer");
  });

  test("throws for an unknown endpoint id and lists the known ones", function () {
    assertThrows(function () {
      buildRequest("does.not.exist", { siteUrl: SITE });
    }, "Unknown endpoint 'does.not.exist'");
  });

  test("uses the declared method when none is supplied", function () {
    const spec = buildRequest("items.update", { siteUrl: SITE, path: { listId: LIST_ID, itemId: 1 } });
    assertEquals(spec.intendedMethod, "PATCH");
    assertEquals(spec.method, "PATCH");
  });

  group("request: query handling");

  test("serialises a query map and percent-encodes values", function () {
    const spec = buildRequest("lists.collection.get", {
      siteUrl: SITE,
      query: { "$select": ["Id", "Title"], "$top": 5 },
    });
    assertEquals(spec.url, SITE + "/_api/web/lists?$select=Id,Title&$top=5");
  });

  test("preserves duplicate parameters through queryPairs", function () {
    const spec = buildRequest("lists.collection.get", {
      siteUrl: SITE,
      queryPairs: [
        { name: "$top", value: 1 },
        { name: "$top", value: 2 },
      ],
    });
    assertContains(spec.url, "$top=1&$top=2", "duplicate parameters should be preserved");
  });

  test("rawQuery is appended verbatim for malformed-input reproduction", function () {
    const spec = buildRequest("lists.collection.get", { siteUrl: SITE, rawQuery: "$$filter=bad" });
    assertEquals(spec.url, SITE + "/_api/web/lists?$$filter=bad");
  });

  test("encodeQueryValues=false leaves the value untouched", function () {
    const spec = buildRequest("lists.collection.get", {
      siteUrl: SITE,
      query: { "$filter": "Title eq 'a b'" },
      encodeQueryValues: false,
    });
    assertContains(spec.url, "Title eq 'a b'", "raw filter should not be encoded");
  });

  group("request: headers and bodies");

  test("GET carries no digest and no If-Match", function () {
    const spec = buildRequest("lists.collection.get", { siteUrl: SITE });
    assertEquals(spec.headers["X-RequestDigest"], undefined);
    assertEquals(spec.headers["If-Match"], undefined);
    assertEquals(spec.headers["Accept"], "application/json;odata=verbose");
  });

  test("writes carry the digest placeholder and a default If-Match", function () {
    const spec = buildRequest("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 1 },
      body: { Title: "x" },
    });
    assertEquals(spec.headers["X-RequestDigest"], DIGEST_PLACEHOLDER);
    assertEquals(spec.headers["If-Match"], "*");
    assertEquals(spec.headers["Content-Type"], "application/json;odata=verbose");
    assertEquals(spec.body, JSON.stringify({ Title: "x" }));
  });

  test("a supplied digest replaces the placeholder", function () {
    const spec = buildRequest("items.delete", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 1 },
      digest: "0xABC",
    });
    assertEquals(spec.headers["X-RequestDigest"], "0xABC");
    assertEquals(spec.body, undefined);
  });

  test("digestPlaceholder=false sends * and explains why", function () {
    const spec = buildRequest("items.delete", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 1 },
      digestPlaceholder: false,
    });
    assertEquals(spec.headers["X-RequestDigest"], "*");
    assertTrue(
      spec.buildNotes.join(" ").indexOf("SharePoint will normally reject this write") >= 0,
      "build notes should explain the '*' digest"
    );
  });

  test("If-Match can be overridden and If-None-Match is supported", function () {
    const spec = buildRequest("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 1 },
      body: { Title: "x" },
      ifMatch: '"1"',
      ifNoneMatch: '"2"',
    });
    assertEquals(spec.headers["If-Match"], '"1"');
    assertEquals(spec.headers["If-None-Match"], '"2"');
  });

  test("MERGE is sent as POST with X-HTTP-Method by default", function () {
    const spec = buildRequest("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 1 },
      body: { Title: "x" },
      method: "MERGE",
    });
    assertEquals(spec.intendedMethod, "MERGE");
    assertEquals(spec.method, "POST");
    assertEquals(spec.headers["X-HTTP-Method"], "MERGE");
  });

  test("useMethodOverride=false sends the literal MERGE verb", function () {
    const spec = buildRequest("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 1 },
      method: "MERGE",
      useMethodOverride: false,
    });
    assertEquals(spec.method, "MERGE");
    assertEquals(spec.headers["X-HTTP-Method"], undefined);
  });

  test("odataMode changes Accept and Content-Type together", function () {
    const spec = buildRequest("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 1 },
      body: { Title: "x" },
      odataMode: "nometadata",
    });
    assertEquals(spec.headers["Accept"], "application/json;odata=nometadata");
    assertEquals(spec.headers["Content-Type"], "application/json;odata=nometadata");
  });

  test("a string body is passed through unchanged (malformed JSON testing)", function () {
    const spec = buildRequest("items.collection.add", {
      siteUrl: SITE,
      path: { listId: LIST_ID },
      body: "{",
    });
    assertEquals(spec.body, "{");
  });

  test("an undeclared method is allowed but flagged as a probe", function () {
    const spec = buildRequest("lists.collection.get", { siteUrl: SITE, method: "DELETE" });
    assertEquals(spec.intendedMethod, "DELETE");
    assertContains(spec.buildNotes.join(" "), "not declared", "probe note should be present");
  });

  test("extra headers are merged and can override defaults", function () {
    const spec = buildRequest("lists.collection.get", {
      siteUrl: SITE,
      headers: { Accept: "text/html", "X-Probe": "1" },
    });
    assertEquals(spec.headers["Accept"], "text/html");
    assertEquals(spec.headers["X-Probe"], "1");
  });

  test("unused path values are reported instead of silently dropped", function () {
    const spec = buildRequest("lists.collection.get", { siteUrl: SITE, path: { listId: LIST_ID } });
    assertContains(spec.buildNotes.join(" "), "not used by the template", "unused path note expected");
  });
}
