import {
  applyEvidenceToCoverage,
  emptyEvidenceLog,
  evidenceFromExecution,
  EvidenceLog,
  parseEvidenceLog,
  summarizeEvidence,
} from "../../src/webparts/hacktheworld/research/evidence";
import {
  CoverageReport,
  ExecutionRecord,
} from "../../src/webparts/hacktheworld/research/types";
import { assertEquals, assertThrows, assertTrue, group, test } from "./harness";

function log(entries: unknown[]): EvidenceLog {
  return parseEvidenceLog({ schemaVersion: 1, entries: entries }, "test log");
}

function entry(overrides: Record<string, unknown>): Record<string, unknown> {
  return {
    endpointId: "lists.collection.get",
    method: "GET",
    outcome: "executed",
    statusCode: 200,
    executedAt: "2026-10-30T10:00:00.000Z",
    ...overrides,
  };
}

function report(): CoverageReport {
  return {
    generatedAt: "2026-10-30T00:00:00.000Z",
    endpoints: { total: 2, implemented: 2, generated: 2, executed: 0, verified: 0 },
    families: { "sharepoint-rest": 2 },
    modules: {
      lists: { endpoints: 2, verified: 0, methods: ["GET"], payloadFamilies: ["idor"] },
    },
    methods: { GET: 2 },
    payloadFamilies: { idor: 2 },
    statuses: { "not-run": 2, executed: 0, interesting: 0, inconclusive: 0, verified: 0 },
    rows: [
      {
        id: "lists.collection.get",
        module: "lists",
        methods: ["GET"],
        status: { implemented: true, generated: true, executed: false, verified: false },
        payloadFamilies: ["idor"],
      },
      {
        id: "lists.item.get",
        module: "lists",
        methods: ["GET"],
        status: { implemented: true, generated: true, executed: false, verified: false },
        payloadFamilies: ["idor"],
      },
    ],
    gaps: [],
    limitations: [
      "No live SharePoint session was used: executed=0 and verified=0 are accurate, not a bug.",
      "Generated means a URL was constructed from documented examples; it does not prove the endpoint works.",
    ],
  };
}

export function registerEvidenceTests(): void {
  group("evidence: parsing");

  test("an empty log is valid and summarises to zero", function () {
    const parsed = parseEvidenceLog({ schemaVersion: 1, entries: [] });
    assertEquals(parsed.entries.length, 0);
    const summary = summarizeEvidence(parsed);
    assertEquals(summary.totalEntries, 0);
    assertEquals(summary.executedEndpointIds, []);
    assertEquals(summary.verifiedEndpointIds, []);
  });

  test("an unsupported schema version is rejected", function () {
    assertThrows(function () {
      parseEvidenceLog({ schemaVersion: 2, entries: [] });
    }, "unsupported schemaVersion");
  });

  test("an unknown outcome is rejected instead of silently counted", function () {
    assertThrows(function () {
      log([entry({ outcome: "probably-vulnerable" })]);
    }, "outcome must be one of");
  });

  test("status codes must be HTTP integers or null", function () {
    assertThrows(function () {
      log([entry({ statusCode: 999 })]);
    }, "integer HTTP status");
    assertThrows(function () {
      log([entry({ statusCode: "200" })]);
    }, "integer HTTP status");
    const parsed = log([entry({ statusCode: null })]);
    assertEquals(parsed.entries[0].statusCode, null);
  });

  test("timestamps must parse as dates", function () {
    assertThrows(function () {
      log([entry({ executedAt: "yesterday" })]);
    }, "ISO-8601");
  });

  test("missing endpointId and method are rejected", function () {
    assertThrows(function () {
      log([entry({ endpointId: "" })]);
    }, "endpointId");
    assertThrows(function () {
      log([entry({ method: "" })]);
    }, "method");
  });

  test("URLs and free text are redacted on parse", function () {
    const parsed = log([
      entry({
        requestUrl: "https://contoso.sharepoint.com/_api/web?digest=SECRET&$top=1",
        notes: "Cookie: FedAuth=abc; Authorization: Bearer abc.def; password=hunter2",
      }),
    ]);
    const stored = parsed.entries[0];
    assertTrue(
      (stored.requestUrl || "").indexOf("SECRET") < 0,
      "digest query value must be redacted"
    );
    assertTrue(
      (stored.requestUrl || "").indexOf("$top=1") >= 0,
      "non-secret query values must be kept"
    );
    assertTrue((stored.notes || "").indexOf("hunter2") < 0, "password must be redacted");
    assertTrue((stored.notes || "").indexOf("abc.def") < 0, "bearer token must be redacted");
  });

  test("the method is normalised to uppercase", function () {
    const parsed = log([entry({ method: "post" })]);
    assertEquals(parsed.entries[0].method, "POST");
  });

  group("evidence: summarising");

  test("executed and verified are derived per endpoint", function () {
    const summary = summarizeEvidence(
      log([
        entry({ endpointId: "lists.collection.get", outcome: "executed", statusCode: 200 }),
        entry({
          endpointId: "lists.collection.get",
          outcome: "verified",
          statusCode: 200,
          executedAt: "2026-10-30T11:00:00.000Z",
        }),
        entry({ endpointId: "lists.item.get", outcome: "inconclusive", statusCode: 403 }),
      ])
    );
    assertEquals(summary.totalEntries, 3);
    assertEquals(summary.executedEndpointIds, ["lists.collection.get", "lists.item.get"]);
    assertEquals(summary.verifiedEndpointIds, ["lists.collection.get"]);
    assertEquals(summary.byEndpoint["lists.collection.get"].lastOutcome, "verified");
    assertEquals(summary.byEndpoint["lists.collection.get"].statusCodes, [200]);
    assertEquals(summary.byEndpoint["lists.item.get"].statusCodes, [403]);
    assertEquals(summary.totals.verified, 1);
    assertEquals(summary.totals.inconclusive, 1);
  });

  test("a not-run record does not count as execution", function () {
    const summary = summarizeEvidence(
      log([entry({ endpointId: "lists.item.get", outcome: "not-run", statusCode: null })])
    );
    assertEquals(summary.executedEndpointIds, []);
    assertEquals(summary.verifiedEndpointIds, []);
    assertEquals(summary.totals["not-run"], 1);
  });

  group("evidence: coverage integration");

  test("evidence marks rows and totals without changing generation status", function () {
    const reportValue = report();
    const updated = applyEvidenceToCoverage(
      reportValue,
      log([
        entry({ endpointId: "lists.collection.get", outcome: "executed" }),
        entry({
          endpointId: "lists.item.get",
          outcome: "verified",
          executedAt: "2026-10-30T12:00:00.000Z",
        }),
      ])
    );
    assertEquals(updated.endpoints.executed, 2);
    assertEquals(updated.endpoints.verified, 1);
    assertEquals(updated.rows[0].status.executed, true);
    assertEquals(updated.rows[0].status.verified, false);
    assertEquals(updated.rows[1].status.verified, true);
    assertEquals(updated.rows[1].status.generated, true);
    assertEquals(updated.modules.lists.verified, 1);
    assertEquals(updated.statuses["not-run"], 0);
  });

  test("unknown endpoint ids do not affect coverage", function () {
    const updated = applyEvidenceToCoverage(
      report(),
      log([entry({ endpointId: "lists.does.not.exist", outcome: "verified" })])
    );
    assertEquals(updated.endpoints.executed, 0);
    assertEquals(updated.endpoints.verified, 0);
    assertEquals(updated.rows[0].status.executed, false);
  });

  test("an empty log leaves zero counts but says why", function () {
    const updated = applyEvidenceToCoverage(report(), emptyEvidenceLog());
    assertEquals(updated.endpoints.executed, 0);
    assertEquals(updated.endpoints.verified, 0);
    assertTrue(
      updated.limitations[0].indexOf("No evidence entries") === 0,
      "limitation should explain that zero is accurate"
    );
  });

  group("evidence: execution records");

  test("evidenceFromExecution redacts a digest URL and response preview", function () {
    const record: ExecutionRecord = {
      endpointId: "lists.item.update",
      method: "POST",
      url: "https://contoso.sharepoint.com/_api/web/lists('x')?digest=SECRET",
      status: 200,
      responseClass: "success",
      durationMs: 42,
      attempts: [],
      responsePreview: "Cookie: FedAuth=abc; ok",
      responseHeaders: {},
      startedAt: "2026-10-30T10:00:00.000Z",
    };
    const stored = evidenceFromExecution(record, "executed", "password=hunter2");
    assertEquals(stored.statusCode, 200);
    assertEquals(stored.executedAt, record.startedAt);
    assertTrue((stored.requestUrl || "").indexOf("SECRET") < 0, "digest must be redacted");
    assertTrue((stored.responsePreview || "").indexOf("FedAuth=abc") < 0, "cookie must be redacted");
    assertTrue((stored.notes || "").indexOf("hunter2") < 0, "password must be redacted");
  });

  test("evidenceFromExecution uses null for a network failure with no status", function () {
    const record: ExecutionRecord = {
      endpointId: "lists.collection.get",
      method: "GET",
      url: "https://contoso.sharepoint.com/_api/web/lists",
      status: 0,
      responseClass: "network",
      durationMs: 5,
      attempts: [],
      responsePreview: "",
      responseHeaders: {},
      startedAt: "2026-10-30T10:00:00.000Z",
      error: "offline",
    };
    assertEquals(evidenceFromExecution(record).statusCode, null);
  });
}
