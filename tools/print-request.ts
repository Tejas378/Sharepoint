/**
 * Print a single generated request as raw HTTP for Burp Suite.
 *
 * This tool never sends anything. It builds the request from the registry,
 * using the documented examples as defaults, and prints it to stdout.
 *
 * Usage:
 *   npm run research:request -- lists.collection.get
 *   npm run research:request -- items.getById --path itemId=2 --method GET
 *   npm run research:request -- items.update --body "{\"Title\":\"probe\"}" --method PATCH
 *   npm run research:request -- lists.collection.get --query "\$top=5" --raw
 *
 * Flags:
 *   --site <url>        Override the configured site URL.
 *   --method <verb>     GET|POST|PATCH|PUT|DELETE|HEAD|OPTIONS|MERGE.
 *   --path name=value   Path placeholder (repeatable).
 *   --query name=value  Query parameter (repeatable).
 *   --header name=value Extra request header (repeatable).
 *   --body <json>       Request body (JSON string).
 *   --body-file <path>  Read the request body from a JSON file (recommended on Windows).
 *   --no-digest         Omit the X-RequestDigest header.
 *   --raw               Print only the raw HTTP request (no comments).
 *   --json              Print the structured HttpRequestSpec as JSON.
 */
import { getEndpoint } from "../src/webparts/hacktheworld/research/registry";
import { buildRequest } from "../src/webparts/hacktheworld/research/request";
import { loadTestConfig } from "../src/webparts/hacktheworld/config/loader";
import {
  BuildRequestInput,
  HttpMethod,
  HttpRequestSpec,
  QueryValue,
} from "../src/webparts/hacktheworld/research/types";
import { getOrigin, getPathAndQuery } from "../src/webparts/hacktheworld/research/url";
import * as fs from "fs";

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

function configuredPathValue(name: string, config: ReturnType<typeof loadTestConfig>): string | number | undefined {
  if (name === "listId" && config.resources.lists.primary.id) {
    return config.resources.lists.primary.id;
  }
  if (name === "itemId") {
    return config.resources.items.primaryItemId;
  }
  if (name === "folderId") {
    return config.resources.folders.primaryFolderId;
  }
  if (name === "viewId" && config.resources.views.viewId) {
    return config.resources.views.viewId;
  }
  if (name === "contentTypeId") {
    return config.resources.contentTypes.customContentTypeId;
  }
  return undefined;
}

function parseArgs(argv: string[]): { endpointId: string; input: BuildRequestInput; raw: boolean; json: boolean } {
  if (argv.length === 0) {
    throw new Error("Missing endpoint id. Example: lists.collection.get");
  }
  const endpointId = argv[0];
  const endpoint = getEndpoint(endpointId);
  const config = loadTestConfig();
  const input: BuildRequestInput = {
    siteUrl: config.environment.siteUrl,
    path: {},
    query: {},
    headers: {},
  };
  let raw = false;
  let json = false;

  // Default path values: central config first, then documented examples.
  for (let i = 0; i < endpoint.parameters.length; i++) {
    const parameter = endpoint.parameters[i];
    if (parameter.location !== "path") {
      continue;
    }
    const configured = configuredPathValue(parameter.name, config);
    if (configured !== undefined && configured !== "") {
      (input.path as Record<string, string | number>)[parameter.name] = configured;
      continue;
    }
    if (parameter.example === undefined) {
      continue;
    }
    (input.path as Record<string, string | number>)[parameter.name] =
      typeof parameter.example === "boolean" ? String(parameter.example) : parameter.example;
  }

  for (let i = 1; i < argv.length; i++) {
    const flag = argv[i];
    if (flag === "--raw") {
      raw = true;
      continue;
    }
    if (flag === "--json") {
      json = true;
      continue;
    }
    if (flag === "--no-digest") {
      input.digestPlaceholder = false;
      continue;
    }
    const value = argv[++i];
    if (value === undefined) {
      throw new Error("Missing value for " + flag);
    }
    if (flag === "--site") {
      input.siteUrl = value;
    } else if (flag === "--method") {
      input.method = value.toUpperCase() as HttpMethod;
    } else if (flag === "--path") {
      const pair = splitAssignment(value);
      (input.path as Record<string, string | number>)[pair.name] = pair.value;
    } else if (flag === "--query") {
      const pair = splitAssignment(value);
      (input.query as Record<string, QueryValue>)[pair.name] = coerce(pair.value);
    } else if (flag === "--header") {
      const pair = splitAssignment(value);
      (input.headers as Record<string, string>)[pair.name] = pair.value;
    } else if (flag === "--body") {
      try {
        input.body = JSON.parse(value);
      } catch (error) {
        throw new Error("--body must be valid JSON: " + value);
      }
    } else if (flag === "--body-file") {
      try {
        input.body = JSON.parse(fs.readFileSync(value, "utf8"));
      } catch (error) {
        throw new Error("--body-file could not be read as JSON: " + value);
      }
    } else {
      throw new Error("Unknown flag: " + flag);
    }
  }
  return { endpointId: endpointId, input: input, raw: raw, json: json };
}

export function toRawHttp(spec: HttpRequestSpec): string {
  const origin = getOrigin(spec.url);
  const lines: string[] = [];
  lines.push(spec.method + " " + getPathAndQuery(spec.url) + " HTTP/1.1");
  lines.push("Host: " + origin.replace(/^https?:\/\//, ""));
  const names = Object.keys(spec.headers);
  for (let i = 0; i < names.length; i++) {
    lines.push(names[i] + ": " + spec.headers[names[i]]);
  }
  lines.push("");
  if (spec.body !== undefined) {
    lines.push(spec.body);
  }
  return lines.join("\r\n");
}

if (require.main === module) {
  try {
    const parsed = parseArgs(process.argv.slice(2));
    const spec = buildRequest(parsed.endpointId, parsed.input);
    if (parsed.json) {
      console.log(JSON.stringify(spec, null, 2));
    } else {
      if (!parsed.raw) {
        console.log("# " + spec.endpointId + " (" + spec.family + ") intended=" + spec.intendedMethod);
        for (let i = 0; i < spec.buildNotes.length; i++) {
          console.log("# note: " + spec.buildNotes[i]);
        }
        console.log("# This request was NOT sent. Paste it into Burp against an authorized tenant.");
        console.log("");
      }
      console.log(toRawHttp(spec));
    }
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 2;
  }
}
