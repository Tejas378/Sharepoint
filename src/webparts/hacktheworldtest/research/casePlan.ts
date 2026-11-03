/**
 * Case planning: turn an endpoint definition plus a payload family into a set
 * of concrete, inspectable HTTP requests.
 *
 * This module is the bridge between the generators (payloads.ts,
 * endpointCases.ts, variations.ts) and the two consumers that need real
 * requests: the in-browser runner and the `research:plan` CLI exporter.
 *
 * It never sends anything. Every request keeps the digest as
 * {{REQUEST_DIGEST}} and strips credential headers so an exported plan can be
 * reviewed, diffed or opened in Burp without leaking a live session.
 */
import { TestConfig } from "../config/loader";
import { DIGEST_PLACEHOLDER, buildRequest } from "./request";
import { generatePayloads } from "./payloads";
import { generateEndpointCases } from "./endpointCases";
import { generateVariations } from "./variations";
import { getEndpoint } from "./registry";
import { redactHeaders, redactUrl } from "./redact";
import {
  BuildRequestInput,
  EndpointDefinition,
  HttpMethod,
  HttpRequestSpec,
  Payload,
  PayloadFamily,
  PayloadGenerationContext,
  QueryValue,
  TestClassification,
} from "./types";

export type CaseSource = "baseline" | "payload" | "endpoint-case" | "variation";

/** Mutable request inputs used while applying a payload / mutation. */
export interface CaseRequestInput {
  path: Record<string, string | number>;
  query: Record<string, QueryValue>;
  headers: Record<string, string>;
  body: unknown;
  /** Set by a method-override payload (POST + X-HTTP-Method). */
  method?: HttpMethod;
}

export interface CasePlanOptions {
  families?: PayloadFamily[];
  /** Include the benign control request for the endpoint. Default true. */
  includeBaseline?: boolean;
  /** Include the generic per-family payload library. Default true. */
  includeGenericPayloads?: boolean;
  /** Include endpoint-derived parameter/body/enum cases. Default true. */
  includeEndpointCases?: boolean;
  /** Include method/query/body/header/identifier variations. Default false. */
  includeVariations?: boolean;
  /** Keep write/method-override cases. Default false: plans are read-first. */
  includeDestructive?: boolean;
  /** Upper bound per endpoint after ordering. Default 100. */
  maxPerEndpoint?: number;
  /** Upper bound for the whole plan. Default 1000. */
  maxTotal?: number;
  /** Explicit path overrides, e.g. { listId: "...", itemId: 2 }. */
  path?: Record<string, string | number>;
  /** Explicit query overrides. */
  query?: Record<string, QueryValue>;
  /** Override the baseline request body (the browser passes the edited body). */
  body?: unknown;
  /** Override the site URL (defaults to config.environment.siteUrl). */
  siteUrl?: string;
}

export interface PlannedCase {
  caseId: string;
  endpointId: string;
  module: string;
  operation: string;
  source: CaseSource;
  family: PayloadFamily | "baseline" | "variation";
  classification: TestClassification;
  name: string;
  rationale: string;
  expectedObservations: string[];
  destructive: boolean;
  control: boolean;
  request: HttpRequestSpec;
  payloadId?: string;
  variationId?: string;
  /** Reason the case is safety-gated or otherwise noteworthy. */
  notes: string[];
}

export interface PlannedEndpoint {
  endpointId: string;
  module: string;
  operation: string;
  pathValues: Record<string, string | number>;
  cases: PlannedCase[];
  skipped: string[];
}

export interface CasePlan {
  generatedAt: string;
  siteUrl: string;
  includeDestructive: boolean;
  endpoints: PlannedEndpoint[];
  totals: {
    endpoints: number;
    cases: number;
    destructive: number;
    skipped: number;
    bySource: Record<string, number>;
    byFamily: Record<string, number>;
  };
  notes: string[];
}

const DESTRUCTIVE_METHODS: HttpMethod[] = ["PUT", "PATCH", "DELETE", "MERGE"];

function cloneInput(input: CaseRequestInput): CaseRequestInput {
  const path: Record<string, string | number> = {};
  const query: Record<string, QueryValue> = {};
  const headers: Record<string, string> = {};
  Object.keys(input.path).forEach(function (key) {
    path[key] = input.path[key];
  });
  Object.keys(input.query).forEach(function (key) {
    query[key] = input.query[key];
  });
  Object.keys(input.headers).forEach(function (key) {
    headers[key] = input.headers[key];
  });
  return {
    path: path,
    query: query,
    headers: headers,
    body: cloneValue(input.body),
    method: input.method,
  };
}

function cloneValue(value: unknown): unknown {
  if (value === undefined || value === null || typeof value !== "object") {
    return value;
  }
  try {
    return JSON.parse(JSON.stringify(value));
  } catch (_error) {
    return value;
  }
}

function readTestCaseValue(config: TestConfig, path: string[]): unknown {
  let current: unknown = config.testData;
  for (let i = 0; i < path.length; i++) {
    if (!current || typeof current !== "object") {
      return undefined;
    }
    current = (current as Record<string, unknown>)[path[i]];
  }
  return current;
}

function asString(value: unknown): string {
  return value === undefined || value === null ? "" : String(value);
}

/**
 * Resolves a path placeholder from the central config. The registry and the UI
 * both use the same names, so one mapping keeps the CLI and the browser in
 * sync instead of each inventing its own defaults.
 */
export function configuredValueForParameter(
  name: string,
  config: TestConfig,
  endpoint?: EndpointDefinition
): string | number | undefined {
  const key = name.toLowerCase();
  const env = config.environment;
  const resources = config.resources;
  const principals = config.principals;

  if (key === "listid") {
    return resources.lists.primary.id || undefined;
  }
  if (key === "itemid" || key === "id") {
    return resources.items.primaryItemId || undefined;
  }
  if (key === "folderid") {
    return resources.folders.primaryFolderId || undefined;
  }
  if (key === "filepath") {
    return resources.files.serverRelativePath || undefined;
  }
  if (key === "folderpath") {
    return asString(resources.folders.serverRelativePath) || asString(resources.lists.primary.serverRelativeUrl) || undefined;
  }
  if (key === "filename") {
    return resources.files.testFileName || undefined;
  }
  if (key === "viewid" || key === "view" || key === "formid") {
    return resources.views.viewId || undefined;
  }
  if (key === "contenttypeid") {
    return resources.contentTypes.customContentTypeId || resources.contentTypes.itemContentTypeId || undefined;
  }
  if (key === "fieldid") {
    return resources.fields.titleFieldId || undefined;
  }
  if (key === "fieldname" || key === "fieldinternalname") {
    return resources.fields.titleFieldName || undefined;
  }
  if (key === "groupid") {
    if (endpoint && endpoint.module === "taxonomy") {
      return config.taxonomy.termGroupId || undefined;
    }
    if (endpoint && endpoint.id === "lists.relatedFields") {
      return resources.fields.relatedFieldGroupId || undefined;
    }
    return principals.groups.primaryGroupId || undefined;
  }
  if (key === "userid") {
    return principals.users.primary.id || undefined;
  }
  if (key === "principalid") {
    return principals.principalIds.length > 0 ? principals.principalIds[0] : principals.users.primary.id || undefined;
  }
  if (key === "roledefid" || key === "roledefinitionid") {
    return principals.roleDefinitions.read || undefined;
  }
  if (key === "siteid") {
    if (!env.siteCollectionId || !env.webId) {
      return undefined;
    }
    try {
      return [new URL(env.siteUrl).hostname, env.siteCollectionId, env.webId].join(",");
    } catch (_error) {
      return undefined;
    }
  }
  if (key === "webid") {
    return env.webId || undefined;
  }
  if (key === "hostname") {
    try {
      return new URL(env.siteUrl).hostname;
    } catch (_error) {
      return undefined;
    }
  }
  if (key === "setid") {
    return config.taxonomy.termSetId || undefined;
  }
  if (key === "termstoreid") {
    return config.taxonomy.termStoreId || undefined;
  }
  if (key === "termgroupid") {
    return config.taxonomy.termGroupId || undefined;
  }
  if (key === "termsetid") {
    return config.taxonomy.termSetId || undefined;
  }
  if (key === "termid") {
    return config.taxonomy.termId || undefined;
  }
  if (key === "commentid") {
    return asString(readTestCaseValue(config, ["constants", "entityIds", "commentId"])) || undefined;
  }
  if (key === "recyclebinitemid") {
    return asString(readTestCaseValue(config, ["constants", "entityIds", "recycleBinItemId"])) || undefined;
  }
  if (key === "attachmentid" || key === "attachmentname") {
    return resources.attachments.fileName || undefined;
  }
  if (key === "uploadid") {
    return resources.files.uploadId || undefined;
  }
  return undefined;
}

/** Builds the baseline input (path + documented query + example body). */
export function createBaseInput(
  endpoint: EndpointDefinition,
  config: TestConfig,
  options?: CasePlanOptions
): CaseRequestInput {
  const path: Record<string, string | number> = {};
  const query: Record<string, QueryValue> = {};
  const overrides = (options && options.path) || {};
  const queryOverrides = (options && options.query) || {};

  for (let i = 0; i < endpoint.parameters.length; i++) {
    const parameter = endpoint.parameters[i];
    if (parameter.location === "path") {
      const explicit = overrides[parameter.name];
      const configured = explicit !== undefined ? explicit : configuredValueForParameter(parameter.name, config, endpoint);
      const value = configured !== undefined ? configured : parameter.example;
      if (value === undefined) {
        throw new Error(
          "No value for path parameter {" + parameter.name + "}; set it in test-config.json or pass --path " +
            parameter.name + "=..."
        );
      }
      path[parameter.name] = typeof value === "boolean" ? String(value) : value;
    } else if (parameter.location === "query") {
      const explicit = queryOverrides[parameter.name];
      if (explicit !== undefined) {
        query[parameter.name] = explicit;
      } else if (parameter.example !== undefined) {
        query[parameter.name] = parameter.example;
      }
    }
  }

  return {
    path: path,
    query: query,
    headers: {},
    body: options && options.body !== undefined
      ? cloneValue(options.body)
      : endpoint.requestBody && endpoint.requestBody.example !== undefined
        ? cloneValue(endpoint.requestBody.example)
        : undefined,
    method: endpoint.methods.length > 0 ? endpoint.methods[0] : "GET",
  };
}

/**
 * Applies one payload / mutation to a mutable request input.
 *
 * Exported so the browser runner and the CLI use exactly the same semantics:
 * path/query/body/header replacement and omission, duplicate query keys,
 * hidden-property merges, and POST + X-HTTP-Method override probes.
 */
export function applyPayloadToRequestInput(
  input: CaseRequestInput,
  payload: Payload,
  endpoint: EndpointDefinition
): void {
  if (payload.mutations && payload.mutations.length > 0) {
    payload.mutations.forEach(function (mutation) {
      const single = Object.assign({}, payload, mutation, { mutations: undefined }) as Payload;
      applyPayloadToRequestInput(input, single, endpoint);
    });
    return;
  }

  if (payload.location === "path" || payload.location === "query") {
    const wireValue = function (value: unknown): string | number | boolean {
      if (value === null) {
        return "null";
      }
      if (value !== null && typeof value === "object") {
        return JSON.stringify(value) || "";
      }
      return value as string | number | boolean;
    };
    let name = payload.parameter;
    const current: Record<string, unknown> = payload.location === "path" ? input.path : input.query;
    const isDeclared = endpoint.parameters.some(function (parameter) {
      return parameter.location === payload.location && parameter.name === name;
    });
    if (payload.location === "path" && !(name in current) && !isDeclared && payload.action !== "omit") {
      const candidate = endpoint.parameters.find(function (parameter) {
        return parameter.location === payload.location;
      });
      if (candidate) {
        name = candidate.name;
      }
    }
    if (payload.location === "path") {
      if (name && payload.action === "omit") {
        input.path[name] = "";
      } else if (name) {
        const value = wireValue(payload.value);
        input.path[name] = typeof value === "boolean" ? String(value) : (value as string | number);
      }
    } else {
      if (name && payload.action === "omit") {
        delete input.query[name];
      } else if (name && payload.action === "duplicate" && Array.isArray(payload.value)) {
        input.query[name] = payload.value.map(wireValue);
      } else if (name) {
        input.query[name] = wireValue(payload.value);
      }
    }
    return;
  }

  if (payload.location === "header") {
    if (payload.action === "omit") {
      deleteHeader(input.headers, payload.parameter);
    } else {
      input.headers[payload.parameter] = String(payload.value);
      // A method-override probe is only meaningful as POST + X-HTTP-Method.
      if (payload.parameter.toLowerCase() === "x-http-method") {
        input.method = "POST";
      }
    }
    return;
  }

  if (payload.action === "omit") {
    if (input.body && typeof input.body === "object") {
      const body = Object.assign({}, input.body) as Record<string, unknown>;
      delete body[payload.parameter];
      input.body = body;
    }
    return;
  }

  if (payload.parameter === "body") {
    if (payload.value && typeof payload.value === "object") {
      const body = Object.assign({}, input.body && typeof input.body === "object" ? input.body : {}) as Record<string, unknown>;
      const values = payload.value as Record<string, unknown>;
      if (Array.isArray(body.formValues)) {
        const formValues = (body.formValues as Array<Record<string, unknown>>).slice();
        Object.keys(values).forEach(function (fieldName) {
          formValues.push({ FieldName: fieldName, FieldValue: values[fieldName] });
        });
        body.formValues = formValues;
      } else if (body.request && typeof body.request === "object" && values.request && typeof values.request === "object") {
        body.request = Object.assign({}, body.request, values.request);
      } else if (body.creationInfo && typeof body.creationInfo === "object") {
        body.creationInfo = Object.assign({}, body.creationInfo, values);
      } else {
        Object.assign(body, values);
      }
      input.body = body;
    } else {
      input.body = payload.value;
    }
    return;
  }

  const body = input.body && typeof input.body === "object"
    ? (Object.assign({}, input.body) as Record<string, unknown>)
    : {};
  const isDeclaredBodyParameter = endpoint.parameters.some(function (parameter) {
    return parameter.location === "body" && parameter.name === payload.parameter;
  });
  if (isDeclaredBodyParameter && payload.parameter !== "body") {
    body[payload.parameter] = payload.value;
  } else if (payload.value && typeof payload.value === "object" && !Array.isArray(payload.value)) {
    const values = payload.value as Record<string, unknown>;
    if (Array.isArray(body.formValues)) {
      const formValues = (body.formValues as Array<Record<string, unknown>>).slice();
      Object.keys(values).forEach(function (fieldName) {
        formValues.push({ FieldName: fieldName, FieldValue: values[fieldName] });
      });
      body.formValues = formValues;
    } else if (body.request && typeof body.request === "object" && values.request && typeof values.request === "object") {
      body.request = Object.assign({}, body.request, values.request);
    } else {
      Object.assign(body, values);
    }
  } else {
    let property = payload.parameter;
    if (!property || property === "Title") {
      property = endpoint.module === "comments" ? "text" : endpoint.module === "search" ? "Querytext" : "Title";
    }
    if (endpoint.module === "search" && body.request && typeof body.request === "object") {
      body.request = Object.assign({}, body.request, { [property]: payload.value });
    } else if (Array.isArray(body.formValues)) {
      const formValues = (body.formValues as Array<Record<string, unknown>>).slice();
      formValues.push({ FieldName: property, FieldValue: payload.value });
      body.formValues = formValues;
    } else {
      body[property] = payload.value;
    }
  }
  input.body = body;
}

function deleteHeader(headers: Record<string, string>, name: string): void {
  const names = Object.keys(headers);
  for (let i = 0; i < names.length; i++) {
    if (names[i].toLowerCase() === name.toLowerCase()) {
      delete headers[names[i]];
    }
  }
}

/** Picks the endpoint-specific mutation target used by the generic payloads. */
export function payloadContextFor(
  endpoint: EndpointDefinition,
  family: PayloadFamily,
  bodyText: string
): PayloadGenerationContext {
  const params = endpoint.parameters;
  const has = function (name: string, location?: string): boolean {
    return params.some(function (parameter) {
      return parameter.name === name && (!location || parameter.location === location);
    });
  };
  const find = function (names: string[], location: string): string {
    for (let i = 0; i < names.length; i++) {
      if (has(names[i], location)) {
        return names[i];
      }
    }
    return "";
  };

  let parameter = "";
  let location: PayloadGenerationContext["location"] = "body";
  if (family === "idor") {
    location = "path";
    parameter = find(["itemId", "filePath", "folderPath", "listId", "groupId", "userId"], "path") || "itemId";
  } else if (family === "path-traversal") {
    location = "path";
    parameter = find(["filePath", "folderPath", "listPath", "serverRelativePath"], "path") || "filePath";
  } else if (family === "odata-injection") {
    location = "query";
    parameter = find(["$filter", "queryText", "$select", "$orderby", "$top"], "query") || "$filter";
  } else if (family === "authz") {
    location = "path";
    parameter = find(["principalId", "groupId", "userId", "itemId", "listId"], "path") || "principalId";
  } else if (family === "method-confusion") {
    location = "header";
    parameter = "X-HTTP-Method";
  } else if (family === "ssrf") {
    location = endpoint.requestBody ? "body" : "query";
    parameter = endpoint.requestBody ? "url" : (find(["url", "queryText"], "query") || "url");
  } else if (family === "business-logic" || family === "mass-assignment") {
    location = "body";
    parameter = "body";
  } else if (endpoint.requestBody) {
    location = "body";
    parameter = endpoint.module === "comments" ? "text" : endpoint.module === "search" ? "Querytext" : "Title";
  } else if (has("queryText", "path")) {
    location = "path";
    parameter = "queryText";
  } else if (has("queryText", "query")) {
    location = "query";
    parameter = "queryText";
  } else if (params.some(function (p) { return p.location === "path"; })) {
    location = "path";
    parameter = params.filter(function (p) { return p.location === "path"; })[0].name;
  } else {
    location = "query";
    parameter = find(["$filter", "FieldInternalName", "queryText"], "query") || "$filter";
  }

  let baselineValue: unknown;
  if (location === "body") {
    try {
      baselineValue = bodyText.trim() ? JSON.parse(bodyText) : endpoint.requestBody && endpoint.requestBody.example;
    } catch (_error) {
      baselineValue = bodyText;
    }
  } else {
    const definition = params.filter(function (p) { return p.name === parameter && p.location === location; })[0];
    baselineValue = definition && definition.example;
  }

  return {
    parameter: parameter,
    location: location,
    resourceType:
      endpoint.module === "files" ? "file" :
      endpoint.module === "folders" ? "folder" :
      endpoint.module === "attachments" ? "attachment" :
      endpoint.module === "users" ? "user" :
      endpoint.module === "groups" ? "group" :
      endpoint.module === "search" ? "search" :
      endpoint.module === "items" ? "item" :
      endpoint.module === "lists" || endpoint.module === "libraries" ? "list" : "unknown",
    baselineValue: baselineValue,
  };
}

function isGenericPayloadRelevant(payload: Payload, endpoint: EndpointDefinition, family: PayloadFamily): boolean {
  if (family === "method-confusion") {
    return true;
  }
  if (payload.mutations && payload.mutations.length > 0) {
    return true;
  }
  if (payload.location === "body") {
    return !!endpoint.requestBody;
  }
  return endpoint.parameters.some(function (parameter) {
    return parameter.location === payload.location &&
      parameter.name.toLowerCase() === payload.parameter.toLowerCase();
  });
}

function isDestructiveRequest(spec: HttpRequestSpec, payload?: Payload): boolean {
  if (spec.destructive || (payload && payload.destructive)) {
    return true;
  }
  if (DESTRUCTIVE_METHODS.indexOf(spec.intendedMethod) >= 0) {
    return true;
  }
  return spec.intendedMethod === "POST" &&
    /x-http-method/i.test(Object.keys(spec.headers).join(","));
}

function exportSpec(spec: HttpRequestSpec): HttpRequestSpec {
  const headers = redactHeaders(spec.headers);
  const names = Object.keys(spec.headers);
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    if (name.toLowerCase() === "x-requestdigest" && spec.headers[name] === DIGEST_PLACEHOLDER) {
      headers[name] = DIGEST_PLACEHOLDER;
    }
  }
  return {
    method: spec.method,
    intendedMethod: spec.intendedMethod,
    url: redactUrl(spec.url),
    headers: headers,
    body: spec.body,
    endpointId: spec.endpointId,
    family: spec.family,
    requiresDigest: spec.requiresDigest,
    destructive: spec.destructive,
    buildNotes: spec.buildNotes.slice(),
  };
}

function buildSpec(
  endpoint: EndpointDefinition,
  input: CaseRequestInput,
  options: CasePlanOptions,
  config: TestConfig
): HttpRequestSpec {
  const buildInput: BuildRequestInput = {
    siteUrl: options.siteUrl || config.environment.siteUrl,
    path: input.path,
    query: input.query,
    method: input.method,
    body: input.body,
    headers: input.headers,
    digestPlaceholder: true,
  };
  return exportSpec(buildRequest(endpoint.id, buildInput));
}

function casePriority(source: CaseSource): number {
  if (source === "baseline") {
    return 0;
  }
  if (source === "endpoint-case") {
    return 1;
  }
  if (source === "payload") {
    return 2;
  }
  return 3;
}

/**
 * Builds all planned cases for one endpoint.
 *
 * A missing path placeholder is reported as a skip instead of inventing a
 * value. Any single generator failure is isolated to that source so one bad
 * endpoint does not abort the whole plan.
 */
export function planEndpoint(
  endpoint: EndpointDefinition,
  config: TestConfig,
  options?: CasePlanOptions
): PlannedEndpoint {
  const opts: CasePlanOptions = options || {};
  const skipped: string[] = [];
  const cases: PlannedCase[] = [];
  const seen: Record<string, boolean> = {};

  let base: CaseRequestInput;
  try {
    base = createBaseInput(endpoint, config, opts);
  } catch (error) {
    return {
      endpointId: endpoint.id,
      module: endpoint.module,
      operation: endpoint.operation,
      pathValues: {},
      cases: [],
      skipped: [error instanceof Error ? error.message : String(error)],
    };
  }

  const add = function (
    caseId: string,
    source: CaseSource,
    family: PlannedCase["family"],
    classification: TestClassification,
    name: string,
    rationale: string,
    expectedObservations: string[],
    destructive: boolean,
    control: boolean,
    request: HttpRequestSpec,
    extras?: { payloadId?: string; variationId?: string }
  ): void {
    if (seen[caseId]) {
      return;
    }
    seen[caseId] = true;
    cases.push({
      caseId: caseId,
      endpointId: endpoint.id,
      module: endpoint.module,
      operation: endpoint.operation,
      source: source,
      family: family,
      classification: classification,
      name: name,
      rationale: rationale,
      expectedObservations: expectedObservations,
      destructive: destructive,
      control: control,
      request: request,
      payloadId: extras && extras.payloadId,
      variationId: extras && extras.variationId,
      notes: destructive && !opts.includeDestructive
        ? ["Safety-gated: destructive case. Re-run with --include-destructive on a disposable fixture."]
        : [],
    });
  };

  if (opts.includeBaseline !== false) {
    try {
      const spec = buildSpec(endpoint, base, opts, config);
      add(
        endpoint.id + ".baseline",
        "baseline",
        "baseline",
        "documented",
        "Baseline request as declared for the endpoint",
        "Control request. Compare every mutation against this response before treating a difference as a finding.",
        ["A normal response for the current user. Record status, body shape, ETag and duration."],
        spec.destructive,
        true,
        spec
      );
    } catch (error) {
      skipped.push("baseline: " + (error instanceof Error ? error.message : String(error)));
    }
  }

  const families = opts.families && opts.families.length > 0 ? opts.families : endpoint.payloadFamilies;

  if (opts.includeGenericPayloads !== false || opts.includeEndpointCases !== false) {
    families.forEach(function (family) {
      const context = payloadContextFor(endpoint, family, base.body === undefined ? "" : JSON.stringify(base.body));
      const payloads: Payload[] = [];
      if (opts.includeEndpointCases !== false) {
        try {
          payloads.push.apply(payloads, generateEndpointCases(endpoint, family, config));
        } catch (error) {
          skipped.push(family + " endpoint cases: " + (error instanceof Error ? error.message : String(error)));
        }
      }
      if (opts.includeGenericPayloads !== false) {
        try {
          generatePayloads(family, context).forEach(function (payload) {
            if (payload.control || isGenericPayloadRelevant(payload, endpoint, family)) {
              payloads.push(payload);
            }
          });
        } catch (error) {
          skipped.push(family + " payloads: " + (error instanceof Error ? error.message : String(error)));
        }
      }

      payloads.forEach(function (payload) {
        const input = cloneInput(base);
        input.method = base.method;
        try {
          applyPayloadToRequestInput(input, payload, endpoint);
          const spec = buildSpec(endpoint, input, opts, config);
          const destructive = isDestructiveRequest(spec, payload);
          add(
            endpoint.id + "." + payload.id,
            payload.id.indexOf("endpoint.") === 0 ? "endpoint-case" : "payload",
            family,
            payload.classification,
            payload.name,
            payload.rationale,
            payload.expectedObservations,
            destructive,
            payload.control,
            spec,
            { payloadId: payload.id }
          );
        } catch (error) {
          skipped.push(family + "/" + payload.id + ": " + (error instanceof Error ? error.message : String(error)));
        }
      });
    });
  }

  if (opts.includeVariations) {
    try {
      const variations = generateVariations(endpoint.id, {
        siteUrl: opts.siteUrl || config.environment.siteUrl,
        path: base.path,
        query: base.query,
        method: base.method,
        body: base.body,
        digestPlaceholder: true,
      });
      variations.forEach(function (variation) {
        if (variation.id === endpoint.id + ".baseline") {
          return;
        }
        const destructive = isDestructiveRequest(variation.request);
        add(
          variation.id,
          "variation",
          "variation",
          variation.classification,
          variation.description,
          variation.rationale,
          variation.expectedObservations || [],
          destructive,
          false,
          exportSpec(variation.request),
          { variationId: variation.id }
        );
      });
    } catch (error) {
      skipped.push("variations: " + (error instanceof Error ? error.message : String(error)));
    }
  }

  const ordered = cases.slice().sort(function (left, right) {
    const delta = casePriority(left.source) - casePriority(right.source);
    return delta !== 0 ? delta : left.caseId.localeCompare(right.caseId);
  });

  const maxPerEndpoint = opts.maxPerEndpoint !== undefined ? opts.maxPerEndpoint : 100;
  const kept: PlannedCase[] = [];
  let destructiveKept = 0;
  for (let i = 0; i < ordered.length && kept.length < maxPerEndpoint; i++) {
    const item = ordered[i];
    if (item.destructive && !opts.includeDestructive) {
      destructiveKept++;
      continue;
    }
    kept.push(item);
  }
  if (destructiveKept > 0) {
    skipped.push(destructiveKept + " destructive case(s) omitted; use --include-destructive on a disposable fixture.");
  }
  if (ordered.length > kept.length + destructiveKept) {
    skipped.push((ordered.length - kept.length - destructiveKept) + " case(s) exceeded maxPerEndpoint=" + maxPerEndpoint + ".");
  }

  return {
    endpointId: endpoint.id,
    module: endpoint.module,
    operation: endpoint.operation,
    pathValues: base.path,
    cases: kept,
    skipped: skipped,
  };
}

/** Plans one or more endpoints, applying the global cap across all of them. */
export function buildCasePlan(
  endpointIds: string[],
  config: TestConfig,
  options?: CasePlanOptions
): CasePlan {
  const opts: CasePlanOptions = options || {};
  const maxTotal = opts.maxTotal !== undefined ? opts.maxTotal : 1000;
  const endpoints: PlannedEndpoint[] = [];
  const notes: string[] = [];
  const unique: Record<string, boolean> = {};
  const bySource: Record<string, number> = {};
  const byFamily: Record<string, number> = {};
  let caseCount = 0;
  let destructiveCount = 0;
  let skippedCount = 0;

  for (let i = 0; i < endpointIds.length; i++) {
    const id = endpointIds[i];
    if (unique[id]) {
      continue;
    }
    unique[id] = true;

    let endpoint: EndpointDefinition;
    try {
      endpoint = getEndpoint(id);
    } catch (error) {
      endpoints.push({
        endpointId: id,
        module: "",
        operation: "",
        pathValues: {},
        cases: [],
        skipped: [error instanceof Error ? error.message : String(error)],
      });
      skippedCount++;
      continue;
    }

    const planned = planEndpoint(endpoint, config, opts);
    skippedCount += planned.skipped.length;

    const remaining = maxTotal - caseCount;
    if (remaining <= 0) {
      planned.cases = [];
      planned.skipped.push("Global maxTotal=" + maxTotal + " reached before this endpoint.");
      endpoints.push(planned);
      continue;
    }
    if (planned.cases.length > remaining) {
      planned.skipped.push((planned.cases.length - remaining) + " case(s) exceeded maxTotal=" + maxTotal + ".");
      planned.cases = planned.cases.slice(0, remaining);
    }

    planned.cases.forEach(function (item) {
      bySource[item.source] = (bySource[item.source] || 0) + 1;
      byFamily[String(item.family)] = (byFamily[String(item.family)] || 0) + 1;
      if (item.destructive) {
        destructiveCount++;
      }
    });
    caseCount += planned.cases.length;
    endpoints.push(planned);
  }

  if (!opts.includeDestructive) {
    notes.push("Destructive cases are excluded by default. Use --include-destructive only on a disposable site collection.");
  }
  notes.push("Requests are generated, not sent. Digest values stay as " + DIGEST_PLACEHOLDER + " and credential headers are redacted.");
  notes.push("A successful response is not a finding. Verify persisted state and effective access with a second account.");

  return {
    generatedAt: new Date().toISOString(),
    siteUrl: opts.siteUrl || config.environment.siteUrl,
    includeDestructive: !!opts.includeDestructive,
    endpoints: endpoints,
    totals: {
      endpoints: endpoints.length,
      cases: caseCount,
      destructive: destructiveCount,
      skipped: skippedCount,
      bySource: bySource,
      byFamily: byFamily,
    },
    notes: notes,
  };
}

/** Markdown index for a plan. Kept here so the CLI only handles file I/O. */
export function renderPlanMarkdown(plan: CasePlan): string {
  const lines: string[] = [];
  lines.push("# Generated security case plan");
  lines.push("");
  lines.push("Generated: " + plan.generatedAt);
  lines.push("");
  lines.push("Target site: `" + plan.siteUrl + "`");
  lines.push("");
  lines.push("| Metric | Value |");
  lines.push("| --- | --- |");
  lines.push("| Endpoints | " + plan.totals.endpoints + " |");
  lines.push("| Cases | " + plan.totals.cases + " |");
  lines.push("| Destructive (included) | " + plan.totals.destructive + " |");
  lines.push("| Skipped notes | " + plan.totals.skipped + " |");
  lines.push("");
  lines.push("## Safety");
  lines.push("");
  for (let i = 0; i < plan.notes.length; i++) {
    lines.push("- " + plan.notes[i]);
  }
  lines.push("");
  lines.push("## Cases");
  lines.push("");
  lines.push("| # | Endpoint | Family | Source | Classification | Method | Case |");
  lines.push("| ---: | --- | --- | --- | --- | --- | --- |");
  let index = 0;
  for (let e = 0; e < plan.endpoints.length; e++) {
    const endpoint = plan.endpoints[e];
    for (let c = 0; c < endpoint.cases.length; c++) {
      const item = endpoint.cases[c];
      index++;
      lines.push(
        "| " + index + " | `" + item.endpointId + "` | " + item.family + " | " + item.source + " | " +
          item.classification + " | " + item.request.method + " | " +
          escapePipes(item.name) + (item.destructive ? " (destructive)" : "") + " |"
      );
    }
    for (let s = 0; s < endpoint.skipped.length; s++) {
      lines.push("| - | `" + endpoint.endpointId + "` | - | skipped | - | - | " + escapePipes(endpoint.skipped[s]) + " |");
    }
  }
  if (index === 0) {
    lines.push("| - | (none) | - | - | - | - | No cases were generated. |");
  }
  lines.push("");
  lines.push("Each request also exists as a raw `.http` file next to this index.");
  lines.push("");
  return lines.join("\n");
}

function escapePipes(value: string): string {
  return value.replace(/\|/g, "\\|").replace(/\r?\n/g, " ");
}
