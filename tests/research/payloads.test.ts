import {
  PAYLOAD_FAMILIES,
  generatePayloads,
  generatePayloadsForFamilies,
} from "../../src/webparts/hacktheworld/research/payloads";
import { PayloadFamily } from "../../src/webparts/hacktheworld/research/types";
import { assertEquals, assertTrue, group, test } from "./harness";

const FAMILIES: PayloadFamily[] = [
  "idor",
  "mass-assignment",
  "odata-injection",
  "xss",
  "sqli",
  "ssrf",
  "path-traversal",
  "method-confusion",
  "business-logic",
  "input-validation",
  "authz",
];

export function registerPayloadTests(): void {
  group("payloads: family metadata");

  test("every family has a description, sink and safety note", function () {
    FAMILIES.forEach(function (family) {
      const info = PAYLOAD_FAMILIES[family];
      assertTrue(!!info, family + " has no family info");
      assertTrue(info.title.length > 0, family + " has no title");
      assertTrue(info.suspectedSink.length > 0, family + " has no suspected sink");
      assertTrue(info.safety.length > 0, family + " has no safety note");
    });
  });

  group("payloads: generation");

  FAMILIES.forEach(function (family) {
    test(family + " starts with a benign control and has unique ids", function () {
      const payloads = generatePayloads(family, { parameter: "Title", location: "body" });
      assertTrue(payloads.length > 1, family + " should generate more than the control");
      assertTrue(payloads[0].control, family + " first payload must be the control");
      assertEquals(payloads[0].family, family);
      const seen: Record<string, boolean> = {};
      for (let i = 0; i < payloads.length; i++) {
        const payload = payloads[i];
        assertTrue(!seen[payload.id], "duplicate payload id " + payload.id);
        seen[payload.id] = true;
        assertTrue(payload.rationale.length > 0, payload.id + " has no rationale");
        assertTrue(payload.expectedObservations.length > 0, payload.id + " has no expected observations");
        assertTrue(payload.parameter.length > 0, payload.id + " has no target parameter");
      }
    });
  });

  test("payloads are JSON-serialisable (no functions or cycles)", function () {
    FAMILIES.forEach(function (family) {
      const payloads = generatePayloads(family, {});
      const text = JSON.stringify(payloads);
      assertTrue(text.length > 2, family + " serialised to nothing");
    });
  });

  test("the control value honours the supplied baseline", function () {
    const payloads = generatePayloads("idor", { parameter: "itemId", baselineValue: 42 });
    assertEquals(payloads[0].value, 42);
  });

  test("context overrides the target parameter and location", function () {
    const payloads = generatePayloads("xss", { parameter: "Description", location: "query" });
    assertEquals(payloads[0].parameter, "Description");
    assertEquals(payloads[0].location, "query");
  });

  test("defaults are used when no context is provided", function () {
    const payloads = generatePayloads("idor");
    assertEquals(payloads[0].parameter, "itemId");
    assertEquals(payloads[0].location, "path");
  });

  test("an unknown family yields no payloads instead of throwing", function () {
    assertEquals(generatePayloads("not-a-family" as PayloadFamily, {}), []);
  });

  test("combined families are de-duplicated by id", function () {
    const combined = generatePayloadsForFamilies(["xss", "idor", "xss"], { parameter: "Title" });
    const seen: Record<string, boolean> = {};
    for (let i = 0; i < combined.length; i++) {
      assertTrue(!seen[combined[i].id], "duplicate id after combination: " + combined[i].id);
      seen[combined[i].id] = true;
    }
    assertTrue(combined.length > 2, "expected payloads from both families");
  });

  test("destructive payloads are explicitly flagged", function () {
    const payloads = generatePayloadsForFamilies(FAMILIES, { parameter: "Title" });
    for (let i = 0; i < payloads.length; i++) {
      assertTrue(
        typeof payloads[i].destructive === "boolean",
        payloads[i].id + " destructive flag must be boolean"
      );
    }
  });
}
