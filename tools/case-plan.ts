/**
 * Generate a Burp-ready security case plan from the endpoint registry.
 *
 * This tool never sends traffic. It applies the same payload families and
 * request mutations as the in-browser runner, then writes:
 *
 *   <out>/index.md          Reviewable case table + safety notes
 *   <out>/plan.json         Structured plan (request specs included)
 *   <out>/requests/*.http   One raw HTTP request per file
 *   <out>/requests.http     All requests in one review file
 *
 * Usage:
 *   npm run research:plan -- --list
 *   npm run research:plan -- --endpoint items.getById --family idor
 *   npm run research:plan -- --module lists --family authz --max 100
 *   npm run research:plan -- --all --out .tmp/cases
 *
 * Flags:
 *   --endpoint <id>          Endpoint id (repeatable).
 *   --module <name>          Every endpoint in a module (repeatable).
 *   --family <name>          Payload family (repeatable; default: endpoint families).
 *   --all                    Plan every registered endpoint.
 *   --max <n>                Global case cap (default 1000).
 *   --per-endpoint <n>       Per-endpoint case cap (default 100).
 *   --out <dir>              Output directory (default .tmp/cases).
 *   --include-destructive    Include write / method-override cases.
 *   --variations             Also include method/query/body/header/ID variations.
 *   --no-payloads            Skip the generic payload library.
 *   --no-endpoint-cases      Skip endpoint-derived cases.
 *   --no-baseline            Skip the baseline control case.
 *   --path name=value        Override a path placeholder (repeatable).
 *   --query name=value       Override a query value (repeatable).
 *   --site <url>             Override the configured site URL.
 *   --json                   Print the plan summary as JSON.
 *   --list                   List modules, endpoint counts and payload families.
 */
import * as fs from "fs";
import * as path from "path";
import { loadTestConfig } from "../src/webparts/hacktheworld/config/loader";
import { listEndpoints, listModules } from "../src/webparts/hacktheworld/research/registry";
import { PAYLOAD_FAMILIES } from "../src/webparts/hacktheworld/research/payloads";
import { buildCasePlan, CasePlan, CasePlanOptions, PlannedCase, renderPlanMarkdown } from "../src/webparts/hacktheworld/research/casePlan";
import { toRawHttp } from "../src/webparts/hacktheworld/research/http";
import { PayloadFamily, QueryValue } from "../src/webparts/hacktheworld/research/types";

interface CliOptions {
  endpointIds: string[];
  modules: string[];
  families: PayloadFamily[];
  all: boolean;
  max: number;
  perEndpoint: number;
  out: string;
  includeDestructive: boolean;
  variations: boolean;
  noPayloads: boolean;
  noEndpointCases: boolean;
  noBaseline: boolean;
  path: Record<string, string | number>;
  query: Record<string, QueryValue>;
  site?: string;
  json: boolean;
  list: boolean;
}

function splitAssignment(value: string): { name: string; value: string } {
  const index = value.indexOf("=");
  if (index <= 0) {
    throw new Error("Expected name=value, received: " + value);
  }
  return { name: value.substring(0, index), value: value.substring(index + 1) };
}

function coerce(value: string): QueryValue {
  if (value === "true") {
    return true;
  }
  if (value === "false") {
    return false;
  }
  if (/^-?\d+$/.test(value)) {
    return Number(value);
  }
  return value;
}

function parseArgs(argv: string[]): CliOptions {
  const options: CliOptions = {
    endpointIds: [],
    modules: [],
    families: [],
    all: false,
    max: 1000,
    perEndpoint: 100,
    out: path.join(process.cwd(), ".tmp", "cases"),
    includeDestructive: false,
    variations: false,
    noPayloads: false,
    noEndpointCases: false,
    noBaseline: false,
    path: {},
    query: {},
    json: false,
    list: false,
  };

  for (let i = 0; i < argv.length; i++) {
    const flag = argv[i];
    if (flag === "--all") {
      options.all = true;
      continue;
    }
    if (flag === "--include-destructive") {
      options.includeDestructive = true;
      continue;
    }
    if (flag === "--variations") {
      options.variations = true;
      continue;
    }
    if (flag === "--no-payloads") {
      options.noPayloads = true;
      continue;
    }
    if (flag === "--no-endpoint-cases") {
      options.noEndpointCases = true;
      continue;
    }
    if (flag === "--no-baseline") {
      options.noBaseline = true;
      continue;
    }
    if (flag === "--json") {
      options.json = true;
      continue;
    }
    if (flag === "--list") {
      options.list = true;
      continue;
    }

    const value = argv[++i];
    if (value === undefined) {
      throw new Error("Missing value for " + flag);
    }
    if (flag === "--endpoint") {
      options.endpointIds.push(value);
    } else if (flag === "--module") {
      options.modules.push(value);
    } else if (flag === "--family") {
      if (!PAYLOAD_FAMILIES[value as PayloadFamily]) {
        throw new Error(
          "Unknown payload family '" + value + "'. Known: " + Object.keys(PAYLOAD_FAMILIES).join(", ")
        );
      }
      options.families.push(value as PayloadFamily);
    } else if (flag === "--max") {
      options.max = positiveInt(value, flag);
    } else if (flag === "--per-endpoint") {
      options.perEndpoint = positiveInt(value, flag);
    } else if (flag === "--out") {
      options.out = path.resolve(value);
    } else if (flag === "--path") {
      const pair = splitAssignment(value);
      options.path[pair.name] = coerce(pair.value) as string | number;
    } else if (flag === "--query") {
      const pair = splitAssignment(value);
      options.query[pair.name] = coerce(pair.value);
    } else if (flag === "--site") {
      options.site = value;
    } else {
      throw new Error("Unknown flag: " + flag);
    }
  }

  return options;
}

function positiveInt(value: string, flag: string): number {
  const parsed = Number(value);
  if (!isFinite(parsed) || Math.floor(parsed) !== parsed || parsed <= 0) {
    throw new Error(flag + " must be a positive integer.");
  }
  return parsed;
}

function selectedEndpointIds(options: CliOptions): string[] {
  const all = listEndpoints();
  const ids: string[] = [];
  const seen: Record<string, boolean> = {};
  const push = function (id: string): void {
    if (!seen[id]) {
      seen[id] = true;
      ids.push(id);
    }
  };

  options.endpointIds.forEach(push);
  options.modules.forEach(function (moduleName) {
    const matches = all.filter(function (endpoint) {
      return endpoint.module === moduleName;
    });
    if (matches.length === 0) {
      throw new Error("No endpoints for module '" + moduleName + "'. Modules: " + listModules().join(", "));
    }
    matches.forEach(function (endpoint) {
      push(endpoint.id);
    });
  });
  if (options.all) {
    all.forEach(function (endpoint) {
      push(endpoint.id);
    });
  }
  if (ids.length === 0) {
    throw new Error(
      "Select endpoints with --endpoint, --module or --all (or use --list to inspect coverage)."
    );
  }
  return ids;
}

function safeFileSlug(value: string, maxLength: number = 100): string {
  const slug = value.toLowerCase().replace(/[^a-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "");
  return (slug || "case").substring(0, maxLength);
}

function writePlan(plan: CasePlan, outDir: string): { files: number; indexPath: string } {
  const requestsDir = path.join(outDir, "requests");
  fs.mkdirSync(requestsDir, { recursive: true });

  let index = 0;
  const combined: string[] = [];
  for (let e = 0; e < plan.endpoints.length; e++) {
    const endpoint = plan.endpoints[e];
    for (let c = 0; c < endpoint.cases.length; c++) {
      const item: PlannedCase = endpoint.cases[c];
      index++;
      const fileName = String(index).padStart(4, "0") + "-" + safeFileSlug(item.caseId) + ".http";
      const raw = toRawHttp(item.request);
      fs.writeFileSync(path.join(requestsDir, fileName), raw + "\r\n", "utf8");

      combined.push("### " + item.caseId);
      combined.push("# endpoint: " + item.endpointId);
      combined.push("# family: " + item.family + " · source: " + item.source + " · classification: " + item.classification);
      combined.push("# case: " + item.name);
      combined.push("# rationale: " + item.rationale);
      for (let n = 0; n < item.expectedObservations.length; n++) {
        combined.push("# observe: " + item.expectedObservations[n]);
      }
      combined.push("# NOT SENT. Digest stays as {{REQUEST_DIGEST}}.");
      combined.push("");
      combined.push(raw);
      combined.push("");
    }
  }

  fs.writeFileSync(path.join(outDir, "requests.http"), combined.join("\r\n"), "utf8");
  fs.writeFileSync(
    path.join(outDir, "plan.json"),
    JSON.stringify(
      {
        plan: plan,
        files: plan.endpoints.reduce(function (sum, endpoint) {
          return sum + endpoint.cases.length;
        }, 0),
      },
      null,
      2
    ),
    "utf8"
  );

  const indexPath = path.join(outDir, "index.md");
  fs.writeFileSync(indexPath, renderPlanMarkdown(plan), "utf8");
  return { files: index, indexPath: indexPath };
}

function printList(): void {
  const endpoints = listEndpoints();
  const counts: Record<string, number> = {};
  endpoints.forEach(function (endpoint) {
    counts[endpoint.module] = (counts[endpoint.module] || 0) + 1;
  });
  console.log("Modules (" + endpoints.length + " endpoints):");
  Object.keys(counts).sort().forEach(function (name) {
    console.log("  " + name + ": " + counts[name]);
  });
  console.log("");
  console.log("Payload families:");
  Object.keys(PAYLOAD_FAMILIES).sort().forEach(function (family) {
    console.log("  " + family + ": " + PAYLOAD_FAMILIES[family as PayloadFamily].title);
  });
}

function main(): void {
  if (require.main !== module) {
    return;
  }
  let options: CliOptions;
  try {
    options = parseArgs(process.argv.slice(2));
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 2;
    return;
  }

  if (options.list) {
    printList();
    return;
  }

  let endpointIds: string[];
  try {
    endpointIds = selectedEndpointIds(options);
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 2;
    return;
  }

  const config = loadTestConfig();
  const planOptions: CasePlanOptions = {
    families: options.families,
    includeBaseline: !options.noBaseline,
    includeGenericPayloads: !options.noPayloads,
    includeEndpointCases: !options.noEndpointCases,
    includeVariations: options.variations,
    includeDestructive: options.includeDestructive,
    maxPerEndpoint: options.perEndpoint,
    maxTotal: options.max,
    path: options.path,
    query: options.query,
    siteUrl: options.site,
  };

  const plan = buildCasePlan(endpointIds, config, planOptions);
  const result = writePlan(plan, options.out);

  if (options.json) {
    console.log(
      JSON.stringify(
        {
          out: options.out,
          indexPath: result.indexPath,
          totals: plan.totals,
          skipped: plan.endpoints
            .filter(function (endpoint) { return endpoint.skipped.length > 0; })
            .map(function (endpoint) { return { endpointId: endpoint.endpointId, skipped: endpoint.skipped }; }),
        },
        null,
        2
      )
    );
    return;
  }

  console.log("Wrote " + result.files + " raw request file(s) to " + path.join(options.out, "requests"));
  console.log("Index: " + result.indexPath);
  console.log("Plan:  " + path.join(options.out, "plan.json"));
  console.log("");
  console.log(
    "Endpoints: " + plan.totals.endpoints + " | cases: " + plan.totals.cases +
      " | destructive included: " + plan.totals.destructive + " | skip notes: " + plan.totals.skipped
  );
  console.log("No request was sent. Digest values remain {{REQUEST_DIGEST}}; add them in your authenticated session.");
  if (!options.includeDestructive) {
    console.log("Destructive cases were excluded. Use --include-destructive only on a disposable site collection.");
  }
}

main();
