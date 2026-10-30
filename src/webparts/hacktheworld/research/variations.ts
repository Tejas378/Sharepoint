import { buildRequest } from "./request";
import { configuredTimestamp } from "../config/runtimeValues";
import { getEndpoint, METHODS_OF_INTEREST } from "./registry";
import {
  BuildRequestInput,
  EndpointDefinition,
  HttpMethod,
  QueryValue,
  RequestChange,
  RequestVariation,
  TestClassification,
} from "./types";

export interface VariationOptions {
  includeMethodProbes?: boolean;
  includeQueryProbes?: boolean;
  includeBodyProbes?: boolean;
  includeHeaderProbes?: boolean;
  includeIdentifierProbes?: boolean;
  /** Upper bound per category; keeps the generated set reviewable. Default 8. */
  maxPerCategory?: number;
}

interface VariationDraft {
  id: string;
  description: string;
  classification: TestClassification;
  rationale: string;
  changes: RequestChange[];
  expectedObservations: string[];
}

/**
 * Builds baseline and mutated request variations for an endpoint.
 *
 * The baseline is always first. Documented variants use documented inputs;
 * probes and hypotheses are labelled so they are never mistaken for supported
 * behaviour.
 */
export function generateVariations(
  endpointId: string,
  base: BuildRequestInput,
  options?: VariationOptions
): RequestVariation[] {
  const endpoint = getEndpoint(endpointId);
  const opts = options || {};
  const max = opts.maxPerCategory !== undefined ? opts.maxPerCategory : 8;
  const out: RequestVariation[] = [];

  out.push({
    id: endpoint.id + ".baseline",
    description: "Baseline request as declared for the endpoint",
    classification: "documented",
    rationale: "Control request. Compare every mutation against this response.",
    changes: [],
    request: buildRequest(endpointId, base),
    expectedObservations: ["A normal response for the current user. Note status, body shape and ETag."],
  });

  if (opts.includeMethodProbes !== false) {
    pushMethods(out, endpoint, base, max);
  }
  if (opts.includeQueryProbes !== false) {
    pushQueryVariants(out, endpoint, base, max);
  }
  if (opts.includeBodyProbes !== false) {
    pushBodyVariants(out, endpoint, base, max);
  }
  if (opts.includeHeaderProbes !== false) {
    pushHeaderVariants(out, endpoint, base);
  }
  if (opts.includeIdentifierProbes !== false) {
    pushIdentifierVariants(out, endpoint, base);
  }
  return out;
}

function pushMethods(
  out: RequestVariation[],
  endpoint: EndpointDefinition,
  base: BuildRequestInput,
  max: number
): void {
  const selected = base.method || endpoint.methods[0];
  let added = 0;
  for (let i = 0; i < endpoint.methods.length; i++) {
    const method = endpoint.methods[i];
    if (method === selected || added >= max) {
      continue;
    }
    added++;
    out.push({
      id: endpoint.id + ".method-declared-" + method.toLowerCase(),
      description: "Declared method: " + method,
      classification: "documented",
      rationale: "The registry declares this method for the endpoint. Same URL, different verb semantics.",
      changes: [{ location: "method", name: "method", before: selected, after: method }],
      request: buildRequest(endpoint.id, withMethod(base, method)),
      expectedObservations: [
        "Expected successful status for the declared method (assuming permissions).",
        "Different error codes between verbs can reveal inconsistent authorization.",
      ],
    });
  }
  if (selected === "MERGE" && base.useMethodOverride !== false) {
    out.push({
      id: endpoint.id + ".method-merge-literal",
      description: "Send MERGE as a literal request method",
      classification: "controlled-probe",
      rationale:
        "SharePoint normally requires POST + X-HTTP-Method: MERGE. Some proxies/services reject or mishandle the literal verb.",
      changes: [{ location: "method", name: "method", before: "POST (override)", after: "MERGE" }],
      request: buildRequest(endpoint.id, withMethod(withOverride(base, false), "MERGE")),
      expectedObservations: ["405/501 instead of a successful update", "Unusual proxy error pages"],
    });
  }
  if (endpoint.methods.indexOf(selected) < 0) {
    return;
  }
  for (let i = 0; i < METHODS_OF_INTEREST.length && added < max; i++) {
    const method = METHODS_OF_INTEREST[i];
    if (endpoint.methods.indexOf(method) >= 0 && method !== selected) {
      continue;
    }
    if (method === selected) {
      continue;
    }
    added++;
    out.push({
      id: endpoint.id + ".method-probe-" + method.toLowerCase(),
      description: "Undeclared method probe: " + method,
      classification: "controlled-probe",
      rationale:
        "The registry does not declare this verb. Testing it checks whether routing and authorization disagree about what the endpoint allows.",
      changes: [{ location: "method", name: "method", before: selected, after: method }],
      request: buildRequest(endpoint.id, withMethod(base, method)),
      expectedObservations: [
        "405 is the expected response. 200/204 for an undeclared destructive verb is interesting.",
        "OPTIONS may advertise supported verbs; compare with behaviour.",
      ],
    });
  }
  const overrideTargets: HttpMethod[] = ["DELETE", "PATCH", "MERGE", "PUT"];
  for (let i = 0; i < overrideTargets.length; i++) {
    const target = overrideTargets[i];
    out.push({
      id: endpoint.id + ".method-override-" + target.toLowerCase(),
      description: "POST with X-HTTP-Method: " + target,
      classification: "hypothesis",
      rationale:
        "Checks whether authorization is evaluated against the effective verb rather than the outer POST.",
      changes: [
        { location: "method", name: "method", before: selected, after: "POST" },
        { location: "header", name: "X-HTTP-Method", after: target },
      ],
      request: buildRequest(endpoint.id, withHeaders(withMethod(base, "POST"), { "X-HTTP-Method": target })),
      expectedObservations: [
        "The effective verb is executed where the direct verb is denied.",
        "Inconsistent status between POST+override and direct request.",
      ],
    });
  }
}

function pushQueryVariants(
  out: RequestVariation[],
  endpoint: EndpointDefinition,
  base: BuildRequestInput,
  max: number
): void {
  const hasQueryOptions =
    endpoint.parameters.some(function (p) {
      return p.location === "query" && p.name.charAt(0) === "$";
    }) || endpoint.module === "items" || endpoint.module === "lists";
  if (!hasQueryOptions) {
    return;
  }

  const documented = endpoint.parameters.filter(function (p) {
    return p.location === "query" && p.example !== undefined;
  });
  for (let i = 0; i < documented.length && i < max; i++) {
    const param = documented[i];
    out.push({
      id: endpoint.id + ".query-documented-" + sanitize(param.name),
      description: "Documented query option: " + param.name,
      classification: "documented",
      rationale: param.description,
      changes: [{ location: "query", name: param.name, after: String(param.example) }],
      request: buildRequest(endpoint.id, withQuery(base, param.name, param.example as QueryValue)),
      expectedObservations: ["Normal result shape for the documented option."],
    });
  }

  const probes: Array<{ name: string; value: QueryValue; note: string }> = [
    { name: "$top", value: 5000, note: "Above the default page size; checks paging/throttling rather than a vulnerability." },
    { name: "$top", value: -1, note: "Negative page size." },
    { name: "$top", value: 0, note: "Zero page size." },
    { name: "$filter", value: "1 eq 1", note: "Tautology filter; results must still be security trimmed." },
    { name: "$filter", value: "ID gt 0 or 1 eq 1", note: "Boolean precedence in the OData parser." },
    { name: "$select", value: "ID,Title,*", note: "Wildcard projection." },
    { name: "$select", value: "ID,Title,Author/Title", note: "Navigation property without $expand." },
    { name: "$expand", value: "Author,Editor", note: "Expand people fields; check for unexpected principal data." },
    { name: "$orderby", value: "ID desc", note: "Ordering by a non-indexed property can reveal query costs." },
    { name: "$skiptoken", value: "Paged=TRUE&p_ID=1", note: "Valid paging token." },
    { name: "$skiptoken", value: "'", note: "Malformed token; look for parser errors." },
  ];
  for (let i = 0; i < probes.length && i < max; i++) {
    const probe = probes[i];
    out.push({
      id: endpoint.id + ".query-probe-" + sanitize(probe.name) + "-" + i,
      description: "Query probe: " + probe.name + "=" + String(probe.value),
      classification: "controlled-probe",
      rationale: probe.note,
      changes: [{ location: "query", name: probe.name, after: String(probe.value) }],
      request: buildRequest(endpoint.id, withQuery(base, probe.name, probe.value)),
      expectedObservations: [
        "Results outside the expected view or permission scope.",
        "500 with a parser/CAML error, or a dramatically slower response.",
      ],
    });
  }

  // Parameter ordering and duplication.
  const firstDocumented = documented.length > 0 ? documented[0] : null;
  if (firstDocumented && firstDocumented.example !== undefined) {
    const name = firstDocumented.name;
    out.push({
      id: endpoint.id + ".query-duplicate",
      description: "Duplicate query parameter: " + name,
      classification: "hypothesis",
      rationale:
        "Different parsers pick the first or last duplicate value. Can sometimes bypass validation applied to one copy.",
      changes: [
        { location: "query", name: name, after: String(firstDocumented.example) + " (twice)" },
      ],
      request: buildRequest(
        endpoint.id,
        withQueryPairs(base, [
          { name: name, value: firstDocumented.example as QueryValue },
          { name: name, value: name === "$top" ? 1 : "probe" },
        ])
      ),
      expectedObservations: ["No error for a duplicate option", "Result count reflects the unvalidated copy"],
    });
    out.push({
      id: endpoint.id + ".query-raw-encoding",
      description: "Unencoded query value (encoding test)",
      classification: "hypothesis",
      rationale:
        "Sending characters unencoded checks whether the service normalises differently from the browser/Burp encoding path.",
      changes: [{ location: "query", name: name, after: String(firstDocumented.example) + " (raw)" }],
      request: buildRequest(
        endpoint.id,
        withRawQuery(base, name + "=" + String(firstDocumented.example) + "&x='\"<>")
      ),
      expectedObservations: ["400 with an encoding error", "Behaviour that differs from the encoded baseline"],
    });
  }
}

function pushBodyVariants(
  out: RequestVariation[],
  endpoint: EndpointDefinition,
  base: BuildRequestInput,
  max: number
): void {
  const definition = endpoint.requestBody;
  if (!definition) {
    return;
  }
  const example = definition.example;
  let added = 0;
  const add = function (draft: VariationDraft, input: BuildRequestInput): void {
    if (added >= max) {
      return;
    }
    added++;
    out.push({
      id: endpoint.id + "." + draft.id,
      description: draft.description,
      classification: draft.classification,
      rationale: draft.rationale,
      changes: draft.changes,
      request: buildRequest(endpoint.id, input),
      expectedObservations: draft.expectedObservations,
    });
  };

  add(
    {
      id: "body-omitted",
      description: "Omit the request body",
      classification: "controlled-probe",
      rationale: "Required-field enforcement and error verbosity when the body is missing.",
      changes: [{ location: "body", name: "body", before: "present", after: "absent" }],
      expectedObservations: ["400 naming required properties", "500 on missing-body handling"],
    },
    withBody(base, undefined)
  );
  add(
    {
      id: "body-empty-object",
      description: "Empty JSON object body",
      classification: "controlled-probe",
      rationale: "Distinguishes 'missing body' from 'body present but empty' validation paths.",
      changes: [{ location: "body", name: "body", after: "{}" }],
      expectedObservations: ["Validation errors listing required fields"],
    },
    withBody(base, {})
  );

  const required = definition.requiredProperties || [];
  const firstRequired = required.length > 0 ? required[0] : "Title";
  const wrongType: Record<string, unknown> = {};
  wrongType[firstRequired] = 12345;
  add(
    {
      id: "body-wrong-type",
      description: "Wrong type for required property '" + firstRequired + "'",
      classification: "controlled-probe",
      rationale: "Type coercion can turn a validation failure into a stored value.",
      changes: [{ location: "body", name: firstRequired, after: "12345 (number)" }],
      expectedObservations: ["Coerced value persisted as '12345'", "500 instead of a validation error"],
    },
    withBody(base, mergeExample(example, wrongType))
  );

  const extraProperty = "ResearchUnexpectedProperty";
  const withExtra: Record<string, unknown> = {};
  withExtra[extraProperty] = "probe";
  add(
    {
      id: "body-extra-property",
      description: "Unexpected extra property in the body",
      classification: "controlled-probe",
      rationale: "Unknown-property handling; some serializers silently bind extra members.",
      changes: [{ location: "body", name: extraProperty, after: "probe" }],
      expectedObservations: ["The extra property is echoed or persisted", "Silently ignored (expected)"],
    },
    withBody(base, mergeExample(example, withExtra))
  );

  const protectedProps = definition.protectedProperties || [];
  for (let i = 0; i < protectedProps.length && added < max; i++) {
    const property = protectedProps[i];
    const body: Record<string, unknown> = {};
    body[property] = property === "Created"
      ? configuredTimestamp("createdDate")
      : property === "Modified"
      ? configuredTimestamp("futureExpiration")
      : 1;
    add(
      {
        id: "body-protected-" + sanitize(property),
        description: "Protected property supplied: " + property,
        classification: "controlled-probe",
        rationale: "Mass-assignment check for a property the service documents as server-controlled.",
        changes: [{ location: "body", name: property, after: String(body[property]) }],
        expectedObservations: ["The value is persisted or echoed", "A clean 400 is the expected safe behaviour"],
      },
      withBody(base, mergeExample(example, body))
    );
  }

  add(
    {
      id: "body-boundary-string",
      description: "32KB string in a text field",
      classification: "controlled-probe",
      rationale: "Length validation and storage limits.",
      changes: [{ location: "body", name: firstRequired, before: "(normal)", after: "32KB string" }],
      expectedObservations: ["Truncation", "500 on length handling", "Silent acceptance above the documented limit"],
    },
    withBody(base, mergeExample(example, { Title: repeatString("A", 32768) }))
  );

  add(
    {
      id: "body-content-type-text",
      description: "JSON body sent as text/plain",
      classification: "hypothesis",
      rationale: "Content-type confusion can bypass body validation or reach a different parser.",
      changes: [{ location: "header", name: "Content-Type", after: "text/plain" }],
      expectedObservations: ["415 Unsupported Media Type (expected)", "The JSON is still parsed and applied"],
    },
    withHeaders(withBody(base, example), { "Content-Type": "text/plain" })
  );
}

function pushHeaderVariants(
  out: RequestVariation[],
  endpoint: EndpointDefinition,
  base: BuildRequestInput
): void {
  const selected = base.method || endpoint.methods[0];
  if (endpoint.requiresDigest) {
    const built = buildRequest(
      endpoint.id,
      withHeaders(base, { "X-RequestDigest": "0000000000000000000000000000000000000000000" })
    );
    out.push({
      id: endpoint.id + ".header-invalid-digest",
      description: "Invalid X-RequestDigest value",
      classification: "controlled-probe",
      rationale:
        "Verifies digest enforcement. A write that succeeds with a bogus digest is interesting; the expected result is 403 security validation.",
      changes: [{ location: "header", name: "X-RequestDigest", after: "invalid value" }],
      request: built,
      expectedObservations: ["403 with 'security validation' text", "A 2xx response means digest validation is not enforced here"],
    });

    const noDigest = buildRequest(endpoint.id, base);
    delete noDigest.headers["X-RequestDigest"];
    noDigest.requiresDigest = false;
    out.push({
      id: endpoint.id + ".header-missing-digest",
      description: "Omit X-RequestDigest entirely",
      classification: "hypothesis",
      rationale: "Checks whether the endpoint becomes reachable without the anti-CSRF token.",
      changes: [{ location: "header", name: "X-RequestDigest", before: "present", after: "absent" }],
      request: noDigest,
      expectedObservations: ["403 security validation (expected)", "2xx without a digest is a CSRF-relevant finding"],
    });
  }

  if (endpoint.conditional) {
    out.push({
      id: endpoint.id + ".header-if-match-mismatch",
      description: "If-Match with a mismatched ETag",
      classification: "controlled-probe",
      rationale: "Optimistic concurrency: a mismatched ETag should yield 412 Precondition Failed.",
      changes: [{ location: "header", name: "If-Match", after: '"1"' }],
      request: buildRequest(endpoint.id, withHeaders(base, { "If-Match": '"1"' })),
      expectedObservations: ["412 (expected)", "A 2xx update means ETag preconditions are ignored"],
    });
    out.push({
      id: endpoint.id + ".header-if-none-match",
      description: "If-None-Match: *",
      classification: "hypothesis",
      rationale: "Conditional creation/update semantics.",
      changes: [{ location: "header", name: "If-None-Match", after: "*" }],
      request: buildRequest(endpoint.id, withHeaders(base, { "If-None-Match": "*" })),
      expectedObservations: ["412 when the resource exists", "A 2xx response indicates the precondition is ignored"],
    });
  }

  if (endpoint.family === "sharepoint-rest" && selected !== "GET") {
    out.push({
      id: endpoint.id + ".header-odata-nometadata",
      description: "Accept/Content-Type application/json;odata=nometadata",
      classification: "controlled-probe",
      rationale: "Response shape and validation can differ between OData modes.",
      changes: [{ location: "header", name: "Accept", after: "application/json;odata=nometadata" }],
      request: buildRequest(endpoint.id, withODataMode(base, "nometadata")),
      expectedObservations: ["Different envelope (no d/results)", "Different validation messages"],
    });
    out.push({
      id: endpoint.id + ".header-odata-minimal",
      description: "OData minimal metadata mode",
      classification: "controlled-probe",
      rationale: "Cross-check serialisation differences.",
      changes: [{ location: "header", name: "odataMode", after: "minimal" }],
      request: buildRequest(endpoint.id, withODataMode(base, "minimal")),
      expectedObservations: ["Metadata-only envelope changes"],
    });
  }
}

function pushIdentifierVariants(
  out: RequestVariation[],
  endpoint: EndpointDefinition,
  base: BuildRequestInput
): void {
  const pathParams = endpoint.parameters.filter(function (parameter) {
    return parameter.location === "path";
  });
  for (let i = 0; i < pathParams.length; i++) {
    const parameter = pathParams[i];
    if (parameter.type === "int") {
      const values: Array<number | string> = [0, -1, 2147483647, -2147483648];
      for (let j = 0; j < values.length; j++) {
        out.push({
          id: endpoint.id + ".path-" + sanitize(parameter.name) + "-" + j,
          description: "Identifier boundary for " + parameter.name + ": " + String(values[j]),
          classification: "controlled-probe",
          rationale: "Identifier parsing, integer overflow and negative-id handling.",
          changes: [{ location: "path", name: parameter.name, after: String(values[j]) }],
          request: buildRequest(endpoint.id, withPath(base, parameter.name, values[j])),
          expectedObservations: [
            "A different object than requested",
            "500 instead of 400/404",
            "Negative ids resolving to another item",
          ],
        });
      }
    } else if (parameter.type === "guid") {
      // The :guid modifier validates the value, so a malformed string would make
      // variation generation throw. Use well-formed GUIDs that the caller is
      // unlikely to own; that is what the IDOR check actually needs.
      const guids = [
        "00000000-0000-0000-0000-000000000000",
        "ffffffff-ffff-ffff-ffff-ffffffffffff",
        "11111111-1111-1111-1111-111111111111",
      ];
      for (let j = 0; j < guids.length; j++) {
        out.push({
          id: endpoint.id + ".path-" + sanitize(parameter.name) + "-" + j,
          description: "Foreign GUID for " + parameter.name + ": " + guids[j],
          classification: "controlled-probe",
          rationale: "Checks whether a well-formed GUID the caller should not own resolves to a resource (IDOR).",
          changes: [{ location: "path", name: parameter.name, after: guids[j] }],
          request: buildRequest(endpoint.id, withPath(base, parameter.name, guids[j])),
          expectedObservations: [
            "404 for every GUID the caller does not own",
            "A resource returned for a synthetic GUID is an IDOR finding",
            "Different error text between valid and invalid GUIDs can leak existence",
          ],
        });
      }
    }
  }
}

function mergeExample(example: unknown, extra: Record<string, unknown>): Record<string, unknown> {
  const merged: Record<string, unknown> = {};
  if (example && typeof example === "object" && !(example instanceof Array)) {
    const keys = Object.keys(example as Record<string, unknown>);
    for (let i = 0; i < keys.length; i++) {
      merged[keys[i]] = (example as Record<string, unknown>)[keys[i]];
    }
  }
  const extraKeys = Object.keys(extra);
  for (let i = 0; i < extraKeys.length; i++) {
    merged[extraKeys[i]] = extra[extraKeys[i]];
  }
  return merged;
}

function withMethod(base: BuildRequestInput, method: HttpMethod): BuildRequestInput {
  return clone(base, { method: method });
}

function withOverride(base: BuildRequestInput, use: boolean): BuildRequestInput {
  return clone(base, { useMethodOverride: use });
}

function withBody(base: BuildRequestInput, body: unknown): BuildRequestInput {
  // Assigned after cloning so an explicit `undefined` removes the body instead
  // of falling back to the baseline body (clone treats undefined as "no change").
  const result = clone(base, {});
  result.body = body;
  return result;
}

function withHeaders(base: BuildRequestInput, headers: Record<string, string>): BuildRequestInput {
  const merged: Record<string, string> = {};
  const baseHeaders = base.headers || {};
  const keys = Object.keys(baseHeaders);
  for (let i = 0; i < keys.length; i++) {
    merged[keys[i]] = baseHeaders[keys[i]];
  }
  const newKeys = Object.keys(headers);
  for (let i = 0; i < newKeys.length; i++) {
    merged[newKeys[i]] = headers[newKeys[i]];
  }
  return clone(base, { headers: merged });
}

function withODataMode(base: BuildRequestInput, mode: "verbose" | "minimal" | "nometadata"): BuildRequestInput {
  return clone(base, { odataMode: mode });
}

function withQuery(base: BuildRequestInput, name: string, value: QueryValue): BuildRequestInput {
  const query: Record<string, QueryValue> = {};
  const existing = base.query || {};
  const keys = Object.keys(existing);
  for (let i = 0; i < keys.length; i++) {
    query[keys[i]] = existing[keys[i]];
  }
  query[name] = value;
  return clone(base, { query: query, queryPairs: [] });
}

function withQueryPairs(base: BuildRequestInput, pairs: Array<{ name: string; value: QueryValue }>): BuildRequestInput {
  return clone(base, { queryPairs: pairs });
}

function withRawQuery(base: BuildRequestInput, raw: string): BuildRequestInput {
  return clone(base, { rawQuery: raw });
}

function withPath(base: BuildRequestInput, name: string, value: string | number): BuildRequestInput {
  const path: Record<string, string | number> = {};
  const existing = base.path || {};
  const keys = Object.keys(existing);
  for (let i = 0; i < keys.length; i++) {
    path[keys[i]] = existing[keys[i]];
  }
  path[name] = value;
  return clone(base, { path: path });
}

function clone(base: BuildRequestInput, changes: BuildRequestInput): BuildRequestInput {
  return {
    siteUrl: changes.siteUrl !== undefined ? changes.siteUrl : base.siteUrl,
    path: changes.path !== undefined ? changes.path : base.path,
    query: changes.query !== undefined ? changes.query : base.query,
    queryPairs: changes.queryPairs !== undefined ? changes.queryPairs : base.queryPairs,
    rawQuery: changes.rawQuery !== undefined ? changes.rawQuery : base.rawQuery,
    body: changes.body !== undefined ? changes.body : base.body,
    method: changes.method !== undefined ? changes.method : base.method,
    headers: changes.headers !== undefined ? changes.headers : base.headers,
    digest: changes.digest !== undefined ? changes.digest : base.digest,
    odataMode: changes.odataMode !== undefined ? changes.odataMode : base.odataMode,
    ifMatch: changes.ifMatch !== undefined ? changes.ifMatch : base.ifMatch,
    ifNoneMatch: changes.ifNoneMatch !== undefined ? changes.ifNoneMatch : base.ifNoneMatch,
    encodeQueryValues: changes.encodeQueryValues !== undefined ? changes.encodeQueryValues : base.encodeQueryValues,
    digestPlaceholder:
      changes.digestPlaceholder !== undefined ? changes.digestPlaceholder : base.digestPlaceholder,
    useMethodOverride:
      changes.useMethodOverride !== undefined ? changes.useMethodOverride : base.useMethodOverride,
    defaultIfMatch: changes.defaultIfMatch !== undefined ? changes.defaultIfMatch : base.defaultIfMatch,
  };
}

function sanitize(value: string): string {
  return value.replace(/[^A-Za-z0-9_-]/g, "_").toLowerCase();
}

function repeatString(value: string, count: number): string {
  let out = "";
  for (let i = 0; i < count; i++) {
    out += value;
  }
  return out;
}
