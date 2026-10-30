/**
 * Evidence log: the only place where "executed" and "verified" may come from.
 *
 * The coverage report must never invent execution. A record is added here only
 * after a request was actually sent (or a manual result was actually observed).
 * Records are plain JSON so they can be committed, diffed and reviewed.
 *
 * Everything that is persisted is redacted first: URLs go through redactUrl()
 * and free text through redactText(), so a digest, cookie or bearer token that
 * slipped into a note cannot be committed.
 *
 * This module is SPFx-free and has no I/O so it can be unit tested in Node.
 */
import { preview, redactText, redactUrl } from "./redact";
import {
  CoverageReport,
  CoverageEndpointRow,
  ExecutionRecord,
  ExecutionStatus,
} from "./types";

export const EVIDENCE_SCHEMA_VERSION = 1;

/** One observed request / manual observation. */
export interface EvidenceEntry {
  /** Registry id, e.g. "lists.collection.get". */
  endpointId: string;
  /** Method that was actually sent, uppercased. */
  method: string;
  /** Manual lifecycle state at the time the record was written. */
  outcome: ExecutionStatus;
  /** HTTP status, or null when no response was received. */
  statusCode: number | null;
  /** ISO-8601 timestamp of the observation. */
  executedAt: string;
  durationMs?: number;
  operator?: string;
  /** Free-form reproduction notes. Redacted before persisting. */
  notes?: string;
  /** Request URL with credential-looking query values redacted. */
  requestUrl?: string;
  /** Truncated, redacted response body for triage. */
  responsePreview?: string;
}

export interface EvidenceLog {
  schemaVersion: number;
  entries: EvidenceEntry[];
}

export interface EvidenceEndpointSummary {
  endpointId: string;
  /** Number of entries for this endpoint, including "not-run". */
  records: number;
  lastOutcome: ExecutionStatus;
  lastExecutedAt: string;
  outcomes: Record<ExecutionStatus, number>;
  statusCodes: number[];
}

export interface EvidenceSummary {
  totalEntries: number;
  /** Endpoint ids with at least one outcome other than "not-run". */
  executedEndpointIds: string[];
  /** Endpoint ids with at least one "verified" outcome. */
  verifiedEndpointIds: string[];
  byEndpoint: Record<string, EvidenceEndpointSummary>;
  totals: Record<ExecutionStatus, number>;
}

const OUTCOMES: ExecutionStatus[] = [
  "not-run",
  "executed",
  "interesting",
  "inconclusive",
  "verified",
];

export function emptyEvidenceLog(): EvidenceLog {
  return { schemaVersion: EVIDENCE_SCHEMA_VERSION, entries: [] };
}

export function isExecutionStatus(value: unknown): value is ExecutionStatus {
  return typeof value === "string" && OUTCOMES.indexOf(value as ExecutionStatus) >= 0;
}

function emptyOutcomes(): Record<ExecutionStatus, number> {
  return { "not-run": 0, executed: 0, interesting: 0, inconclusive: 0, verified: 0 };
}

function requireString(value: unknown, field: string, source: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(source + ": " + field + " must be a non-empty string.");
  }
  return value;
}

function optionalString(
  value: unknown,
  field: string,
  source: string
): string | undefined {
  if (value === undefined || value === null) {
    return undefined;
  }
  if (typeof value !== "string") {
    throw new Error(source + ": " + field + " must be a string when present.");
  }
  return value;
}

function readStatusCode(value: unknown, field: string, source: string): number | null {
  if (value === undefined || value === null) {
    return null;
  }
  if (
    typeof value !== "number" ||
    Math.floor(value) !== value ||
    value < 100 ||
    value > 599
  ) {
    throw new Error(source + ": " + field + " must be an integer HTTP status (100-599) or null.");
  }
  return value;
}

function readDuration(value: unknown, field: string, source: string): number | undefined {
  if (value === undefined || value === null) {
    return undefined;
  }
  if (typeof value !== "number" || value < 0 || !isFinite(value)) {
    throw new Error(source + ": " + field + " must be a non-negative number when present.");
  }
  return value;
}

function readTimestamp(value: unknown, field: string, source: string): string {
  const text = requireString(value, field, source);
  if (isNaN(Date.parse(text))) {
    throw new Error(source + ": " + field + " must be an ISO-8601 timestamp.");
  }
  return text;
}

/**
 * Validates and normalises an evidence log.
 *
 * Unknown top-level fields are ignored so newer tools can add metadata without
 * breaking older ones. Unknown entry fields are dropped, never silently kept:
 * a typo in "outcome" is an error, not a record that counts as "executed".
 */
export function parseEvidenceLog(value: unknown, source: string = "evidence log"): EvidenceLog {
  if (!value || typeof value !== "object" || value instanceof Array) {
    throw new Error(source + ": expected an object with schemaVersion and entries.");
  }
  const root = value as Record<string, unknown>;
  if (root.schemaVersion !== EVIDENCE_SCHEMA_VERSION) {
    throw new Error(
      source + ": unsupported schemaVersion " + String(root.schemaVersion) +
        " (expected " + EVIDENCE_SCHEMA_VERSION + ")."
    );
  }
  if (!(root.entries instanceof Array)) {
    throw new Error(source + ": entries must be an array.");
  }

  const entries: EvidenceEntry[] = [];
  for (let i = 0; i < root.entries.length; i++) {
    const at = source + ": entries[" + i + "]";
    const raw = root.entries[i];
    if (!raw || typeof raw !== "object" || raw instanceof Array) {
      throw new Error(at + " must be an object.");
    }
    const input = raw as Record<string, unknown>;
    const outcome = input.outcome;
    if (!isExecutionStatus(outcome)) {
      throw new Error(
        at + ".outcome must be one of " + OUTCOMES.join(", ") + " (received " +
          String(outcome) + ")."
      );
    }
    const entry: EvidenceEntry = {
      endpointId: requireString(input.endpointId, "endpointId", at),
      method: requireString(input.method, "method", at).toUpperCase(),
      outcome: outcome,
      statusCode: readStatusCode(input.statusCode, "statusCode", at),
      executedAt: readTimestamp(input.executedAt, "executedAt", at),
    };

    const durationMs = readDuration(input.durationMs, "durationMs", at);
    if (durationMs !== undefined) {
      entry.durationMs = durationMs;
    }
    const operator = optionalString(input.operator, "operator", at);
    if (operator !== undefined) {
      entry.operator = redactText(operator);
    }
    const notes = optionalString(input.notes, "notes", at);
    if (notes !== undefined) {
      entry.notes = redactText(notes);
    }
    const requestUrl = optionalString(input.requestUrl, "requestUrl", at);
    if (requestUrl !== undefined) {
      entry.requestUrl = redactUrl(requestUrl);
    }
    const responsePreview = optionalString(input.responsePreview, "responsePreview", at);
    if (responsePreview !== undefined) {
      entry.responsePreview = preview(redactText(responsePreview));
    }
    entries.push(entry);
  }

  return { schemaVersion: EVIDENCE_SCHEMA_VERSION, entries: entries };
}

/**
 * Builds an evidence entry from a RequestExecutor result. The caller still
 * decides the manual outcome; this only guarantees that what is stored is
 * redacted and complete.
 */
export function evidenceFromExecution(
  record: ExecutionRecord,
  outcome: ExecutionStatus = "executed",
  notes?: string
): EvidenceEntry {
  if (!isExecutionStatus(outcome)) {
    throw new Error("Unsupported evidence outcome: " + String(outcome));
  }
  const entry: EvidenceEntry = {
    endpointId: record.endpointId,
    method: record.method,
    outcome: outcome,
    statusCode: record.status > 0 ? record.status : null,
    executedAt: record.startedAt,
    durationMs: record.durationMs,
    requestUrl: redactUrl(record.url),
    responsePreview: preview(redactText(record.responsePreview)),
  };
  if (notes) {
    entry.notes = redactText(notes);
  }
  return entry;
}

/** Aggregates records without changing them. */
export function summarizeEvidence(log: EvidenceLog): EvidenceSummary {
  const byEndpoint: Record<string, EvidenceEndpointSummary> = {};
  const totals = emptyOutcomes();
  const executed: Record<string, boolean> = {};
  const verified: Record<string, boolean> = {};

  for (let i = 0; i < log.entries.length; i++) {
    const entry = log.entries[i];
    totals[entry.outcome]++;
    if (entry.outcome !== "not-run") {
      executed[entry.endpointId] = true;
    }
    if (entry.outcome === "verified") {
      verified[entry.endpointId] = true;
    }

    let summary = byEndpoint[entry.endpointId];
    if (!summary) {
      summary = {
        endpointId: entry.endpointId,
        records: 0,
        lastOutcome: "not-run",
        lastExecutedAt: "",
        outcomes: emptyOutcomes(),
        statusCodes: [],
      };
      byEndpoint[entry.endpointId] = summary;
    }
    summary.records++;
    summary.outcomes[entry.outcome]++;
    if (entry.statusCode !== null && summary.statusCodes.indexOf(entry.statusCode) < 0) {
      summary.statusCodes.push(entry.statusCode);
    }
    // ISO-8601 strings with the same offset sort lexicographically. Entries from
    // different offsets are still compared as dates to stay correct.
    if (
      summary.lastExecutedAt === "" ||
      Date.parse(entry.executedAt) >= Date.parse(summary.lastExecutedAt)
    ) {
      summary.lastExecutedAt = entry.executedAt;
      summary.lastOutcome = entry.outcome;
    }
  }

  return {
    totalEntries: log.entries.length,
    executedEndpointIds: Object.keys(executed).sort(),
    verifiedEndpointIds: Object.keys(verified).sort(),
    byEndpoint: byEndpoint,
    totals: totals,
  };
}

function toLookup(ids: string[]): Record<string, boolean> {
  const lookup: Record<string, boolean> = {};
  for (let i = 0; i < ids.length; i++) {
    lookup[ids[i]] = true;
  }
  return lookup;
}

/**
 * Returns a copy of the coverage report where executed/verified are derived
 * from the evidence log instead of being hardcoded to false.
 *
 * "not-run" in the status block means "endpoint has no execution evidence",
 * not "no record was written". That keeps the report useful after partial runs.
 */
export function applyEvidenceToCoverage(
  report: CoverageReport,
  log: EvidenceLog
): CoverageReport {
  const summary = summarizeEvidence(log);
  const executedSet = toLookup(summary.executedEndpointIds);
  const verifiedSet = toLookup(summary.verifiedEndpointIds);

  const rows: CoverageEndpointRow[] = [];
  const modules: CoverageReport["modules"] = {};
  const moduleNames = Object.keys(report.modules);
  for (let m = 0; m < moduleNames.length; m++) {
    const original = report.modules[moduleNames[m]];
    modules[moduleNames[m]] = {
      endpoints: original.endpoints,
      verified: 0,
      methods: original.methods.slice(),
      payloadFamilies: original.payloadFamilies.slice(),
    };
  }

  for (let i = 0; i < report.rows.length; i++) {
    const row = report.rows[i];
    const executed = executedSet[row.id] === true;
    const verified = verifiedSet[row.id] === true;
    rows.push({
      id: row.id,
      module: row.module,
      methods: row.methods.slice(),
      status: {
        implemented: row.status.implemented,
        generated: row.status.generated,
        executed: executed,
        verified: verified,
      },
      payloadFamilies: row.payloadFamilies.slice(),
    });
    if (verified && modules[row.module]) {
      modules[row.module].verified++;
    }
  }

  const statuses: Record<ExecutionStatus, number> = {
    "not-run": report.rows.length - summary.executedEndpointIds.length,
    executed: summary.totals.executed,
    interesting: summary.totals.interesting,
    inconclusive: summary.totals.inconclusive,
    verified: summary.totals.verified,
  };
  if (statuses["not-run"] < 0) {
    statuses["not-run"] = 0;
  }

  const limitations: string[] = [];
  for (let i = 0; i < report.limitations.length; i++) {
    if (report.limitations[i].indexOf("No live SharePoint session") === 0) {
      continue;
    }
    limitations.push(report.limitations[i]);
  }
  limitations.unshift(
    summary.totalEntries === 0
      ? "No evidence entries: executed=0 and verified=0 are accurate, not a bug."
      : "Executed/verified come from " + summary.totalEntries +
          " evidence entr" + (summary.totalEntries === 1 ? "y" : "ies") +
          "; generated still only proves construction."
  );

  return {
    generatedAt: report.generatedAt,
    endpoints: {
      total: report.endpoints.total,
      implemented: report.endpoints.implemented,
      generated: report.endpoints.generated,
      executed: summary.executedEndpointIds.length,
      verified: summary.verifiedEndpointIds.length,
    },
    families: report.families,
    modules: modules,
    methods: report.methods,
    payloadFamilies: report.payloadFamilies,
    statuses: statuses,
    rows: rows,
    gaps: report.gaps.slice(),
    limitations: limitations,
  };
}
