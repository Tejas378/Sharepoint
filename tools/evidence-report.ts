/**
 * Evidence report.
 *
 * Reads docs/evidence/evidence.json (or a file given with --file) and prints
 * what was actually executed. It never sends traffic and never invents a
 * result: a missing file is reported as "no evidence", not as a failure.
 *
 * Run with:
 *   npm run research:evidence
 *   npm run research:evidence -- --file docs/evidence/evidence.json --json
 */
import * as fs from "fs";
import * as path from "path";
import {
  emptyEvidenceLog,
  EvidenceLog,
  parseEvidenceLog,
  summarizeEvidence,
} from "../src/webparts/hacktheworld/research/evidence";
import { ENDPOINTS } from "../src/webparts/hacktheworld/research/registry";
import { ExecutionStatus } from "../src/webparts/hacktheworld/research/types";

const DEFAULT_EVIDENCE_FILE = path.join(process.cwd(), "docs", "evidence", "evidence.json");
const OUTCOME_ORDER: ExecutionStatus[] = [
  "not-run",
  "executed",
  "interesting",
  "inconclusive",
  "verified",
];

function readLog(file: string): { log: EvidenceLog; found: boolean } {
  if (!fs.existsSync(file)) {
    return { log: emptyEvidenceLog(), found: false };
  }
  const text = fs.readFileSync(file, "utf8");
  return { log: parseEvidenceLog(JSON.parse(text), file), found: true };
}

function knownEndpointIds(): Record<string, boolean> {
  const ids: Record<string, boolean> = {};
  for (let i = 0; i < ENDPOINTS.length; i++) {
    ids[ENDPOINTS[i].id] = true;
  }
  return ids;
}

function markdown(file: string, found: boolean, log: EvidenceLog): string {
  const summary = summarizeEvidence(log);
  const known = knownEndpointIds();
  const lines: string[] = [];
  lines.push("# Evidence report");
  lines.push("");
  lines.push("File: `" + file + "`" + (found ? "" : " (not found)"));
  lines.push("");
  lines.push("| Outcome | Entries |");
  lines.push("| --- | --- |");
  for (let i = 0; i < OUTCOME_ORDER.length; i++) {
    const outcome = OUTCOME_ORDER[i];
    lines.push("| " + outcome + " | " + summary.totals[outcome] + " |");
  }
  lines.push("");
  lines.push(
    "Executed endpoints: " + summary.executedEndpointIds.length +
      " | Verified endpoints: " + summary.verifiedEndpointIds.length
  );
  lines.push("");
  lines.push("| Endpoint | Records | Last outcome | Last executed | Status codes |");
  lines.push("| --- | --- | --- | --- | --- |");
  const ids = Object.keys(summary.byEndpoint).sort();
  for (let i = 0; i < ids.length; i++) {
    const item = summary.byEndpoint[ids[i]];
    lines.push(
      "| " + item.endpointId + " | " + item.records + " | " + item.lastOutcome + " | " +
        item.lastExecutedAt + " | " +
        (item.statusCodes.length > 0 ? item.statusCodes.join(", ") : "-") + " |"
    );
  }
  if (ids.length === 0) {
    lines.push("| (none) | 0 | - | - | - |");
  }
  lines.push("");

  const unknown = ids.filter(function (id) {
    return !known[id];
  });
  if (unknown.length > 0) {
    lines.push("## Unknown endpoint ids");
    lines.push("");
    lines.push(
      "These entries do not match the registry, so they do not affect coverage. " +
        "Fix the id or add the endpoint before relying on them."
    );
    lines.push("");
    for (let i = 0; i < unknown.length; i++) {
      lines.push("- " + unknown[i]);
    }
    lines.push("");
  }
  return lines.join("\n");
}

function main(): void {
  if (require.main !== module) {
    return;
  }
  const args = process.argv.slice(2);
  let file = DEFAULT_EVIDENCE_FILE;
  let asJson = false;
  for (let i = 0; i < args.length; i++) {
    if (args[i] === "--json") {
      asJson = true;
    } else if (args[i] === "--file" && args[i + 1]) {
      file = path.resolve(args[i + 1]);
      i++;
    }
  }

  let read: { log: EvidenceLog; found: boolean };
  try {
    read = readLog(file);
  } catch (error) {
    console.error(
      "Invalid evidence log: " + (error instanceof Error ? error.message : String(error))
    );
    process.exitCode = 1;
    return;
  }

  const summary = summarizeEvidence(read.log);
  const known = knownEndpointIds();
  const unknown = Object.keys(summary.byEndpoint).filter(function (id) {
    return !known[id];
  });

  if (asJson) {
    console.log(
      JSON.stringify(
        {
          file: file,
          found: read.found,
          totalEntries: summary.totalEntries,
          executedEndpointIds: summary.executedEndpointIds,
          verifiedEndpointIds: summary.verifiedEndpointIds,
          totals: summary.totals,
          unknownEndpointIds: unknown,
        },
        null,
        2
      )
    );
  } else {
    console.log(markdown(file, read.found, read.log));
  }
}

main();
