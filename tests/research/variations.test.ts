import { buildRequest } from "../../src/webparts/hacktheworld/research/request";
import { getEndpoint } from "../../src/webparts/hacktheworld/research/registry";
import { generateVariations } from "../../src/webparts/hacktheworld/research/variations";
import { assertEquals, assertTrue, group, test } from "./harness";

const SITE = "https://contoso.sharepoint.com/sites/team";
const LIST_ID = "4a81ba4f-84ed-4e91-a731-9a82c0ede37e";

function idsOf(variations: Array<{ id: string }>): string[] {
  return variations.map(function (variation) {
    return variation.id;
  });
}

function countMatching(variations: Array<{ id: string }>, fragment: string): number {
  return variations.filter(function (variation) {
    return variation.id.indexOf(fragment) >= 0;
  }).length;
}

export function registerVariationTests(): void {
  group("variations: baseline");

  test("the baseline is always the first variation and is documented", function () {
    const variations = generateVariations("lists.collection.get", { siteUrl: SITE });
    assertEquals(variations[0].id, "lists.collection.get.baseline");
    assertEquals(variations[0].classification, "documented");
    assertEquals(variations[0].changes, []);
  });

  test("the baseline request is identical to buildRequest", function () {
    const variations = generateVariations("items.getById", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
    });
    const expected = buildRequest("items.getById", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
    });
    assertEquals(variations[0].request, expected);
    assertEquals(variations[0].request.endpointId, "items.getById");
  });

  test("every variation carries an id, rationale and expected observations", function () {
    const variations = generateVariations("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
      body: { Title: "probe" },
    });
    assertTrue(variations.length > 1, "expected mutations in addition to the baseline");
    for (let i = 0; i < variations.length; i++) {
      const variation = variations[i];
      assertTrue(variation.id.length > 0, "variation " + i + " has no id");
      assertTrue(variation.description.length > 0, variation.id + " has no description");
      assertTrue(variation.rationale.length > 0, variation.id + " has no rationale");
      assertTrue(
        !!variation.expectedObservations && variation.expectedObservations.length > 0,
        variation.id + " has no expected observations"
      );
      assertEquals(variation.request.endpointId, "items.update");
    }
  });

  test("variation ids are unique for an endpoint", function () {
    const variations = generateVariations("lists.collection.get", { siteUrl: SITE });
    const seen: Record<string, boolean> = {};
    const ids = idsOf(variations);
    for (let i = 0; i < ids.length; i++) {
      assertTrue(!seen[ids[i]], "duplicate variation id: " + ids[i]);
      seen[ids[i]] = true;
    }
  });

  group("variations: method probes");

  test("an undeclared verb is generated as a controlled probe", function () {
    const variations = generateVariations("lists.collection.get", { siteUrl: SITE });
    const probe = variations.filter(function (variation) {
      return variation.id === "lists.collection.get.method-probe-post";
    })[0];
    assertTrue(!!probe, "expected a POST method probe for a GET endpoint");
    assertEquals(probe.classification, "controlled-probe");
    assertEquals(probe.request.intendedMethod, "POST");
    assertEquals(probe.request.method, "POST");
  });

  test("all eight verbs are probed at most once", function () {
    const variations = generateVariations("lists.collection.get", { siteUrl: SITE }, { maxPerCategory: 20 });
    const methods = variations
      .filter(function (variation) {
        return variation.id.indexOf(".method-probe-") >= 0;
      })
      .map(function (variation) {
        return variation.request.intendedMethod;
      });
    const seen: Record<string, boolean> = {};
    for (let i = 0; i < methods.length; i++) {
      assertTrue(!seen[methods[i]], "duplicate method probe: " + methods[i]);
      seen[methods[i]] = true;
    }
    assertTrue(methods.length >= 6, "expected probes for the verbs not declared by the endpoint");
  });

  test("includeMethodProbes=false removes every method mutation", function () {
    const variations = generateVariations("lists.collection.get", { siteUrl: SITE }, { includeMethodProbes: false });
    assertEquals(countMatching(variations, ".method-"), 0);
    assertEquals(variations.length, 1 + countMatching(variations, ".query-") + countMatching(variations, ".path-"));
  });

  test("declared write verbs are generated as documented variations", function () {
    const variations = generateVariations("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
      body: { Title: "probe" },
    });
    const declared = variations.filter(function (variation) {
      return variation.id.indexOf(".method-declared-") >= 0;
    });
    assertTrue(declared.length > 0, "items.update declares more than one write verb");
  });

  group("variations: query probes");

  test("documented query options and malformed probes are both generated", function () {
    const variations = generateVariations("lists.collection.get", { siteUrl: SITE });
    assertTrue(countMatching(variations, ".query-documented-") > 0, "expected documented query variants");
    assertTrue(countMatching(variations, ".query-probe-") > 0, "expected query probes");
    assertTrue(countMatching(variations, ".query-duplicate") === 1, "expected one duplicate-parameter probe");
  });

  test("tautology filters are sent as controlled probes", function () {
    const variations = generateVariations("lists.collection.get", { siteUrl: SITE });
    const tautology = variations.filter(function (variation) {
      return variation.request.url.indexOf("1+eq+1") >= 0 || variation.request.url.indexOf("1%20eq%201") >= 0;
    });
    assertTrue(tautology.length > 0, "expected at least one tautology probe");
  });

  test("includeQueryProbes=false removes query variants", function () {
    const variations = generateVariations("lists.collection.get", { siteUrl: SITE }, { includeQueryProbes: false });
    assertEquals(countMatching(variations, ".query-"), 0);
  });

  test("maxPerCategory bounds the generated probes", function () {
    const bounded = generateVariations("lists.collection.get", { siteUrl: SITE }, { maxPerCategory: 1 });
    assertEquals(countMatching(bounded, ".query-probe-"), 1);
    assertEquals(countMatching(bounded, ".method-probe-"), 1);
  });

  test("an endpoint without query options gets no query variants", function () {
    const variations = generateVariations("rest.contextinfo", { siteUrl: SITE });
    assertEquals(countMatching(variations, ".query-"), 0);
  });

  group("variations: body probes");

  test("body probes include omitted and empty bodies for write endpoints", function () {
    const variations = generateVariations("items.collection.add", {
      siteUrl: SITE,
      path: { listId: LIST_ID },
      body: { Title: "probe" },
    });
    assertTrue(countMatching(variations, ".body-omitted") === 1, "expected an omitted-body probe");
    assertTrue(countMatching(variations, ".body-empty-object") === 1, "expected an empty-object probe");
    const omitted = variations.filter(function (variation) {
      return variation.id === "items.collection.add.body-omitted";
    })[0];
    assertEquals(omitted.request.body, undefined);
  });

  test("protected properties produce mass-assignment probes", function () {
    const variations = generateVariations("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
      body: { Title: "probe" },
    });
    assertTrue(
      countMatching(variations, ".body-protected-") > 0,
      "items.update declares protected properties, so probes must exist"
    );
  });

  test("includeBodyProbes=false removes body variants", function () {
    const variations = generateVariations(
      "items.collection.add",
      { siteUrl: SITE, path: { listId: LIST_ID }, body: { Title: "probe" } },
      { includeBodyProbes: false }
    );
    assertEquals(countMatching(variations, ".body-"), 0);
  });

  test("an endpoint without a request body gets no body variants", function () {
    const variations = generateVariations("lists.collection.get", { siteUrl: SITE });
    assertEquals(countMatching(variations, ".body-"), 0);
  });

  group("variations: header probes");

  test("digest and ETag probes are generated for a conditional write", function () {
    const variations = generateVariations("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
      body: { Title: "probe" },
    });
    ["header-invalid-digest", "header-missing-digest", "header-if-match-mismatch", "header-if-none-match"].forEach(
      function (fragment) {
        assertTrue(countMatching(variations, fragment) === 1, "expected " + fragment);
      }
    );
  });

  test("the missing-digest variation actually removes the header", function () {
    const variations = generateVariations("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
      body: { Title: "probe" },
    });
    const missing = variations.filter(function (variation) {
      return variation.id === "items.update.header-missing-digest";
    })[0];
    assertEquals(missing.request.headers["X-RequestDigest"], undefined);
    assertEquals(missing.request.requiresDigest, false);
  });

  test("the invalid-digest variation sends a value that is not the placeholder", function () {
    const variations = generateVariations("items.update", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
      body: { Title: "probe" },
    });
    const invalid = variations.filter(function (variation) {
      return variation.id === "items.update.header-invalid-digest";
    })[0];
    assertTrue(
      invalid.request.headers["X-RequestDigest"] !== "{{REQUEST_DIGEST}}",
      "the invalid digest must not be replaced by the executor"
    );
  });

  test("a read-only endpoint has no digest probes", function () {
    const variations = generateVariations("lists.collection.get", { siteUrl: SITE });
    assertEquals(countMatching(variations, "header-invalid-digest"), 0);
    assertEquals(countMatching(variations, "header-missing-digest"), 0);
  });

  group("variations: identifier probes");

  test("integer path parameters get boundary probes", function () {
    const variations = generateVariations("items.getById", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
    });
    assertTrue(countMatching(variations, ".path-itemid-") === 4, "expected four integer boundary probes");
    const zero = variations.filter(function (variation) {
      return variation.id === "items.getById.path-itemid-0";
    })[0];
    assertTrue(!!zero, "expected the zero-id probe");
    assertTrue(zero.request.url.indexOf("/items(0)") >= 0, "zero probe must change the URL");
  });

  test("GUID path parameters get foreign-GUID probes", function () {
    const variations = generateVariations("lists.getById", {
      siteUrl: SITE,
      path: { listId: LIST_ID },
    });
    assertTrue(countMatching(variations, ".path-listid-") === 3, "expected three GUID probes");
    const foreign = variations.filter(function (variation) {
      return variation.id === "lists.getById.path-listid-1";
    })[0];
    assertTrue(
      foreign.request.url.indexOf("ffffffff-ffff-ffff-ffff-ffffffffffff") >= 0,
      "the foreign GUID must reach the URL"
    );
  });

  test("includeIdentifierProbes=false removes path variants", function () {
    const variations = generateVariations(
      "items.getById",
      { siteUrl: SITE, path: { listId: LIST_ID, itemId: 2 } },
      { includeIdentifierProbes: false }
    );
    assertEquals(countMatching(variations, ".path-"), 0);
  });

  test("the baseline is unaffected by probe switches", function () {
    const base = { siteUrl: SITE, path: { listId: LIST_ID, itemId: 2 } };
    const full = generateVariations("items.getById", base);
    const minimal = generateVariations("items.getById", base, {
      includeMethodProbes: false,
      includeQueryProbes: false,
      includeBodyProbes: false,
      includeHeaderProbes: false,
      includeIdentifierProbes: false,
    });
    assertEquals(minimal.length, 1);
    assertEquals(minimal[0], full[0]);
  });

  test("unknown endpoint ids throw instead of generating nothing", function () {
    let threw = false;
    try {
      generateVariations("does.not.exist", { siteUrl: SITE });
    } catch (error) {
      threw = true;
      assertTrue(
        String(error instanceof Error ? error.message : error).indexOf("Unknown endpoint") >= 0,
        "error must mention the unknown endpoint"
      );
    }
    assertTrue(threw, "generateVariations must reject unknown endpoint ids");
  });

  test("every generated request targets the endpoint's own URL template", function () {
    const endpoint = getEndpoint("items.getById");
    const variations = generateVariations("items.getById", {
      siteUrl: SITE,
      path: { listId: LIST_ID, itemId: 2 },
    });
    for (let i = 0; i < variations.length; i++) {
      const url = variations[i].request.url.split("?")[0];
      assertTrue(
        url.indexOf("/_api/web/lists(") >= 0 && url.indexOf("/items(") >= 0,
        variations[i].id + " escaped the endpoint template: " + url
      );
    }
    assertEquals(endpoint.module, "items");
  });
}
