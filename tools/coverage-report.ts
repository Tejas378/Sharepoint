/**
 * Coverage report generator.
 *
 * Reads the endpoint registry and reports what the toolkit can actually do.
 * It deliberately does NOT claim execution: there is no live-tenant result
 * store in this repository, so `executed` and `verified` are always false here.
 *
 * Status definitions (see PROJECT_CONTEXT.md):
 *   implemented - an EndpointDefinition exists in the registry
 *   generated   - buildRequest() produced a complete URL from documented examples
 *   executed    - a real request was sent (requires a transport + live session)
 *   verified    - behaviour was confirmed manually (requires execution evidence)
 *
 * Run with:
 *   npm run research:coverage
 */
import { ENDPOINTS } from "../src/webparts/hacktheworld/research/registry";
import { buildRequest } from "../src/webparts/hacktheworld/research/request";
import {
  CoverageEndpointRow,
  CoverageReport,
  EndpointDefinition,
  HttpMethod,
  PayloadFamily,
} from "../src/webparts/hacktheworld/research/types";

const PLACEHOLDER_SITE = "https://research.invalid/sites/toolkit";

function documentedPath(endpoint: EndpointDefinition): {
  path: Record<string, string | number>;
  complete: boolean;
} {
  const path: Record<string, string | number> = {};
  let complete = true;
  for (let i = 0; i < endpoint.parameters.length; i++) {
    const parameter = endpoint.parameters[i];
    if (parameter.location !== "path") {
      continue;
    }
    if (parameter.example === undefined) {
      complete = false;
      continue;
    }
    path[parameter.name] =
      typeof parameter.example === "boolean" ? String(parameter.example) : parameter.example;
  }
  return { path: path, complete: complete };
}

function canGenerate(endpoint: EndpointDefinition): boolean {
  const input = documentedPath(endpoint);
  if (!input.complete) {
    return false;
  }
  try {
    buildRequest(endpoint.id, { siteUrl: PLACEHOLDER_SITE, path: input.path });
    return true;
  } catch (error) {
    return false;
  }
}

function unique<T>(values: T[]): T[] {
  const seen: Record<string, boolean> = {};
  const out: T[] = [];
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i]);
    if (!seen[key]) {
      seen[key] = true;
      out.push(values[i]);
    }
  }
  return out;
}

export function buildCoverageReport(): CoverageReport {
  const rows: CoverageEndpointRow[] = [];
  const modules: CoverageReport["modules"] = {};
  const families: Record<string, number> = {};
  const methods: Record<string, number> = {};
  const payloadFamilies: Record<string, number> = {};
  const gaps: string[] = [];

  for (let i = 0; i < ENDPOINTS.length; i++) {
    const endpoint = ENDPOINTS[i];
    const generated = canGenerate(endpoint);
    rows.push({
      id: endpoint.id,
      module: endpoint.module,
      methods: endpoint.methods,
      status: {
        implemented: true,
        generated: generated,
        executed: false,
        verified: false,
      },
      payloadFamilies: endpoint.payloadFamilies,
    });

    if (!modules[endpoint.module]) {
      modules[endpoint.module] = {
        endpoints: 0,
        verified: 0,
        methods: [],
        payloadFamilies: [],
      };
    }
    const module = modules[endpoint.module];
    module.endpoints++;
    module.methods = unique(module.methods.concat(endpoint.methods));
    module.payloadFamilies = unique(module.payloadFamilies.concat(endpoint.payloadFamilies));

    for (let m = 0; m < endpoint.methods.length; m++) {
      methods[endpoint.methods[m]] = (methods[endpoint.methods[m]] || 0) + 1;
    }
    for (let p = 0; p < endpoint.payloadFamilies.length; p++) {
      payloadFamilies[endpoint.payloadFamilies[p]] =
        (payloadFamilies[endpoint.payloadFamilies[p]] || 0) + 1;
    }
    families[endpoint.family] = (families[endpoint.family] || 0) + 1;

    if (!generated) {
      gaps.push(
        endpoint.id +
          ": no complete documented example for every path placeholder, so a baseline cannot be generated automatically."
      );
    }
    if (endpoint.destructive) {
      gaps.push(endpoint.id + ": destructive; blocked by safety.allowDestructive=false.");
    }
  }

  // Modules that the prompt explicitly calls out but the registry does not yet
  // cover. Keep this list honest; remove entries as modules are added.
  ["taxonomy", "workflow", "webhooks (tenant-level)", "microsoft-graph users/groups"].forEach(
    function (name) {
      if (!modules[name]) {
        gaps.push("Module not covered yet: " + name + ".");
      }
    }
  );

  return {
    generatedAt: new Date().toISOString(),
    endpoints: {
      total: rows.length,
      implemented: rows.length,
      generated: rows.filter(function (row) {
        return row.status.generated;
      }).length,
      executed: 0,
      verified: 0,
    },
    families: families,
    modules: modules,
    methods: methods,
    payloadFamilies: payloadFamilies,
    statuses: {
      "not-run": rows.length,
      executed: 0,
      interesting: 0,
      inconclusive: 0,
      verified: 0,
    },
    rows: rows,
    gaps: gaps,
    limitations: [
      "No live SharePoint session was used: executed=0 and verified=0 are accurate, not a bug.",
      "Generated means a URL was constructed from documented examples; it does not prove the endpoint works.",
      "HEAD, OPTIONS, PUT and MERGE are generated as variations, not as declared endpoint methods.",
      "Microsoft Graph requires a separate bearer token and is not exercised by the SharePoint cookie transport.",
      "Coverage counts implementations, not vulnerabilities. No finding is confirmed without manual evidence.",
    ],
  };
}

function markdown(report: CoverageReport): string {
  const lines: string[] = [];
  lines.push("# Endpoint coverage report");
  lines.push("");
  lines.push("Generated: " + report.generatedAt);
  lines.push("");
  lines.push("| Metric | Value |");
  lines.push("| --- | --- |");
  lines.push("| Registered endpoints | " + report.endpoints.total + " |");
  lines.push("| Implemented | " + report.endpoints.implemented + " |");
  lines.push("| Baseline generated | " + report.endpoints.generated + " |");
  lines.push("| Executed | " + report.endpoints.executed + " |");
  lines.push("| Verified | " + report.endpoints.verified + " |");
  lines.push("");
  lines.push("## Modules");
  lines.push("");
  lines.push("| Module | Endpoints | Methods | Payload families |");
  lines.push("| --- | --- | --- | --- |");
  const moduleNames = Object.keys(report.modules).sort();
  for (let i = 0; i < moduleNames.length; i++) {
    const name = moduleNames[i];
    const module = report.modules[name];
    lines.push(
      "| " + name + " | " + module.endpoints + " | " + module.methods.join(", ") + " | " +
        module.payloadFamilies.join(", ") + " |"
    );
  }
  lines.push("");
  lines.push("## Declared methods");
  lines.push("");
  const methodNames = Object.keys(report.methods).sort();
  for (let i = 0; i < methodNames.length; i++) {
    lines.push("- " + methodNames[i] + ": " + report.methods[methodNames[i]] + " endpoint(s)");
  }
  lines.push("");
  lines.push("## Gaps");
  lines.push("");
  for (let i = 0; i < report.gaps.length; i++) {
    lines.push("- " + report.gaps[i]);
  }
  lines.push("");
  lines.push("## Limitations");
  lines.push("");
  for (let i = 0; i < report.limitations.length; i++) {
    lines.push("- " + report.limitations[i]);
  }
  lines.push("");
  return lines.join("\n");
}

if (require.main === module) {
  const report = buildCoverageReport();
  const arg = process.argv[2];
  if (arg === "--json") {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(markdown(report));
  }
}
