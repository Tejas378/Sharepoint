import { EndpointDefinition, Payload, PayloadFamily } from "./types";
import { TestConfig } from "../config/loader";

const MAX_CASES_PER_ENDPOINT = 300;

interface ParameterVariant {
  label: string;
  value: unknown;
}

function isMutationEndpoint(endpoint: EndpointDefinition): boolean {
  return endpoint.destructive || endpoint.methods.some((method) => ["PUT", "PATCH", "DELETE", "MERGE"].indexOf(method) >= 0);
}

function createCase(
  endpoint: EndpointDefinition,
  family: PayloadFamily,
  parameter: string,
  location: Payload["location"],
  label: string,
  value: unknown,
  rationale: string,
  action: Payload["action"] = "replace",
  forceDestructive: boolean = false
): Payload {
  return {
    id: "endpoint." + endpoint.id + "." + family + "." + location + "." + parameter + "." + label
      .toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, ""),
    family: family,
    name: location + " · " + parameter + " · " + label,
    classification: "controlled-probe",
    location: location,
    parameter: parameter,
    value: value,
    action: action,
    rationale: rationale,
    expectedObservations: [
      "Compare the result with the documented behavior for this endpoint and this test identity.",
      "For reads, verify returned objects and fields are within the signed-in user's access.",
      "For writes, verify the resulting state and permissions from the second test account.",
    ],
    control: false,
    destructive: forceDestructive || isMutationEndpoint(endpoint),
  };
}

function idVariants(name: string, config: TestConfig): ParameterVariant[] {
  const key = name.toLowerCase();
  const out: ParameterVariant[] = [];
  const add = (label: string, value: unknown): void => {
    if (value !== undefined && value !== null && value !== "" && value !== 0) {
      if (!out.some((candidate) => String(candidate.value) === String(value))) out.push({ label: label, value: value });
    }
  };
  if (key.indexOf("list") >= 0) {
    Object.keys(config.resources.lists).forEach((listKey) => add("configured list · " + listKey, config.resources.lists[listKey].id));
  }
  if (key.indexOf("item") >= 0 || key === "id") {
    add("primary item", config.resources.items.primaryItemId);
    add("secondary item", config.resources.items.secondaryItemId);
    add("nonexistent item", config.resources.items.nonExistentItemId);
  }
  if (key.indexOf("user") >= 0 || key.indexOf("principal") >= 0) {
    add("primary user", config.principals.users.primary.id);
    add("secondary user", config.principals.users.secondary.id);
    config.principals.principalIds.forEach((id) => add("configured principal", id));
  }
  if (key.indexOf("group") >= 0 || key.indexOf("principal") >= 0) {
    add("primary group", config.principals.groups.primaryGroupId);
    add("owner group", config.principals.groups.ownerGroupId);
    add("member group", config.principals.groups.memberGroupId);
    add("visitor group", config.principals.groups.visitorGroupId);
    config.principals.groups.groupsToEnumerate.forEach((id) => add("configured group", id));
  }
  if (key.indexOf("site") >= 0) {
    add("site collection", config.environment.siteCollectionId);
    add("site", config.environment.siteId);
    if (key === "siteid" && config.environment.siteCollectionId && config.environment.webId) {
      try {
        add("configured composite site id", new URL(config.environment.siteUrl).hostname + "," + config.environment.siteCollectionId + "," + config.environment.webId);
      } catch (_error) { /* The config validator reports an invalid site URL. */ }
    }
  }
  if (key.indexOf("web") >= 0) add("web", config.environment.webId);
  if (key.indexOf("folder") >= 0) add("primary folder", config.resources.folders.primaryFolderId);
  return out;
}

function variantsForType(type: EndpointDefinition["parameters"][number]["type"], example?: unknown, allowedValues?: Array<string | number | boolean>): ParameterVariant[] {
  const values: ParameterVariant[] = [{ label: "null", value: null }, { label: "empty", value: "" }];
  if (type === "int") {
    values.push(
      { label: "zero", value: 0 }, { label: "negative one", value: -1 },
      { label: "maximum signed 32-bit", value: 2147483647 },
      { label: "minimum signed 32-bit", value: -2147483648 },
      { label: "overflow integer", value: "2147483648" },
      { label: "underflow integer", value: "-2147483649" },
      { label: "fractional number", value: "1.5" }, { label: "numeric exponent", value: "1e309" },
      { label: "hex-like integer", value: "0x10" }, { label: "leading zero", value: "0001" },
      { label: "plus sign integer", value: "+1" }, { label: "integer with whitespace", value: " 1 " }
    );
  } else if (type === "guid") {
    values.push(
      { label: "malformed identifier", value: "not-a-guid" },
      { label: "zero GUID", value: "00000000-0000-0000-0000-000000000000" },
      { label: "braced GUID", value: "{00000000-0000-0000-0000-000000000000}" },
      { label: "uppercase GUID", value: "AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE" },
      { label: "GUID without separators", value: "aaaaaaaabbbbccccddddeeeeeeeeeeee" },
      { label: "GUID with trailing data", value: "00000000-0000-0000-0000-000000000000-probe" }
    );
  } else if (type === "bool") {
    values.push({ label: "false", value: false }, { label: "true", value: true }, { label: "zero boolean", value: 0 }, { label: "one boolean", value: 1 }, { label: "unexpected boolean text", value: "yes" }, { label: "on boolean", value: "on" }, { label: "off boolean", value: "off" });
  } else if (type === "date") {
    values.push({ label: "invalid date", value: "not-a-date" }, { label: "epoch date", value: "1970-01-01T00:00:00Z" }, { label: "far future date", value: "9999-12-31T23:59:59Z" }, { label: "leap day", value: "2024-02-29T23:59:59Z" }, { label: "invalid leap day", value: "2023-02-29T00:00:00Z" }, { label: "timezone offset", value: "2024-01-01T00:30:00+14:00" }, { label: "date without timezone", value: "2024-01-01T00:00:00" });
  } else if (type === "enum") {
    (allowedValues || []).forEach((value, index) => values.push({ label: "declared value " + index + " · " + String(value), value: value }));
    values.push({ label: "undeclared enum value", value: "__unexpected_enum_value__" });
  } else if (type === "json") {
    values.push(
      { label: "empty array", value: [] }, { label: "empty object", value: {} },
      { label: "wrong JSON scalar", value: "wrong-type" }, { label: "malformed JSON text", value: "{" },
      { label: "nested array", value: [["probe"]] }, { label: "deep object", value: { a: { b: { c: "probe" } } } },
      { label: "unexpected null property", value: { probe: null } }, { label: "unexpected boolean property", value: { probe: true } }
    );
  } else if (type === "path") {
    values.push(
      { label: "encoded slash", value: "%2f" }, { label: "double encoded slash", value: "%252f" },
      { label: "encoded dot segments", value: "%2e%2e%2f" }, { label: "double encoded dot segments", value: "%252e%252e%252f" },
      { label: "backslash", value: "..\\security-probe" }, { label: "mixed separators", value: "..%2f..\\security-probe" },
      { label: "unicode path", value: "security-probe-你好" }, { label: "absolute URL path", value: "https://example.invalid/security-probe" }
    );
  } else if (type === "binary") {
    values.push({ label: "empty body", value: "" }, { label: "text in binary field", value: "controlled-probe" });
  } else {
    values.push(
      { label: "whitespace", value: " " }, { label: "unicode", value: "probe-üñîçødé-你好-😀" },
      { label: "encoded delimiter", value: "%2f%3f%23" }, { label: "double encoded delimiter", value: "%252f%253f%2523" },
      { label: "CRLF characters", value: "probe\\r\\nHeader: value" }, { label: "null character", value: "probe\\u0000tail" },
      { label: "right-to-left override", value: "probe\\u202e.txt" }, { label: "lone surrogate", value: "\\uD800" },
      { label: "256 character boundary", value: new Array(257).join("A") }, { label: "4KB boundary", value: new Array(4097).join("A") }
    );
  }
  if (example !== undefined && !values.some((variant) => JSON.stringify(variant.value) === JSON.stringify(example))) {
    values.unshift({ label: "declared example", value: example });
  }
  return values;
}

/**
 * Adds endpoint-specific parameter cases to the shared payload families.
 * Cases are generated from that endpoint's own parameter/body metadata; each
 * remains a single operator-selected request in the UI.
 */
export function generateEndpointCases(endpoint: EndpointDefinition, family: PayloadFamily, config: TestConfig): Payload[] {
  const cases: Payload[] = [];
  const destructive = isMutationEndpoint(endpoint);

  if (family === "input-validation") {
    const dimensions: Array<{ parameter: EndpointDefinition["parameters"][number]; variants: ParameterVariant[] }> = [];
    endpoint.parameters.forEach((parameter) => {
      if (cases.length < MAX_CASES_PER_ENDPOINT) {
        const missingLabel = parameter.location === "path" ? "empty path segment" : "omit parameter";
        cases.push(createCase(endpoint, family, parameter.name, parameter.location, missingLabel, undefined, "Checks required/optional parameter handling.", "omit", destructive));
      }
      const variants = variantsForType(parameter.type, parameter.example, parameter.allowedValues);
      dimensions.push({ parameter: parameter, variants: variants });
      variants.forEach((variant) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) {
          cases.push(createCase(endpoint, family, parameter.name, parameter.location, variant.label, variant.value, "Type-aware boundary probe derived from the endpoint parameter definition.", "replace", destructive));
        }
      });
      if (parameter.location === "query" && parameter.example !== undefined && cases.length < MAX_CASES_PER_ENDPOINT) {
        cases.push(createCase(endpoint, family, parameter.name, "query", "duplicate query parameter", [String(parameter.example), String(parameter.example)], "Checks how repeated query keys are resolved.", "duplicate", destructive));
      }
      if (parameter.location === "query" && parameter.example !== undefined && cases.length < MAX_CASES_PER_ENDPOINT) {
        cases.push(createCase(endpoint, family, parameter.name, "query", "conflicting duplicate query parameter", [String(parameter.example), "__conflicting_probe__"], "Checks whether first-value or last-value parsing changes validation or authorization.", "duplicate", destructive));
      }
    });
    if (endpoint.requestBody) {
      endpoint.requestBody.requiredProperties?.forEach((name) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, name, "body", "omit required body property", undefined, "Checks server-side validation of required request properties.", "omit", destructive));
      });
      if (endpoint.requestBody.example !== undefined && cases.length < MAX_CASES_PER_ENDPOINT) {
        cases.push(createCase(endpoint, family, "body", "body", "null request body", null, "Checks null handling against the documented request body shape.", "replace", destructive));
      }
      if (cases.length < MAX_CASES_PER_ENDPOINT) {
        cases.push(createCase(endpoint, family, "body", "body", "malformed JSON body", "{", "Checks malformed JSON handling against the endpoint's declared JSON content type.", "replace", destructive));
      }
      ["[]", "null", "true", "{\"Title\":\"first\",\"Title\":\"second\"}", "{\"Title\":\"probe\"} trailing", "{\"__proto__\":{\"probe\":true}}"].forEach((rawBody) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, "body", "body", "raw JSON parser edge · " + rawBody.slice(0, 24), rawBody, "Checks parser behavior for primitive, duplicate-key, trailing-data, and prototype-like JSON keys. Use only a disposable fixture for writes.", "replace", destructive));
      });
      ["application/json;odata=verbose", "application/json;odata=minimalmetadata", "application/json;odata=nometadata", "text/plain"].forEach((contentType) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, "Content-Type", "header", "content type · " + contentType, contentType, "Checks parser and validation behavior across content types.", "replace", destructive));
      });
      ["application/xml", "application/octet-stream", "application/x-www-form-urlencoded", "application/json-patch+json"].forEach((contentType) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, "Content-Type", "header", "content type mismatch · " + contentType, contentType, "Checks whether body parsing, validation, and authorization remain consistent when media types do not match the body.", "replace", destructive));
      });
      ["application/json", "application/json;odata=verbose", "application/xml", "text/html", "*/*"].forEach((accept) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, "Accept", "header", "accept · " + accept, accept, "Checks representation negotiation and response-path consistency.", "replace", destructive));
      });
      ["identity", "gzip", "br", "deflate"].forEach((encoding) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, "Content-Encoding", "header", "content encoding · " + encoding, encoding, "Checks rejection or safe handling of a body that is not encoded as declared.", "replace", destructive));
      });
    }
    ["If-Match", "If-None-Match"].forEach((headerName) => {
      ["*", "W/\"stale-probe\"", "\"stale-probe\"", "invalid-etag"].forEach((etag) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, headerName, "header", headerName + " · " + etag, etag, "Checks stale-write protection and conditional-read behavior; observe whether unauthorized or stale writes are rejected.", "replace", destructive));
      });
    });
    for (let left = 0; left < dimensions.length && cases.length < MAX_CASES_PER_ENDPOINT; left++) {
      for (let right = left + 1; right < dimensions.length && cases.length < MAX_CASES_PER_ENDPOINT; right++) {
        const first = dimensions[left];
        const second = dimensions[right];
        const firstNull = first.variants.find((variant) => variant.label === "null");
        const secondNull = second.variants.find((variant) => variant.label === "null");
        const firstEdge = first.variants.find((variant) => variant.label !== "null" && variant.label !== "empty" && variant.label.indexOf("declared example") !== 0);
        const secondEdge = second.variants.find((variant) => variant.label !== "null" && variant.label !== "empty" && variant.label.indexOf("declared example") !== 0);
        if (firstNull && secondEdge && cases.length < MAX_CASES_PER_ENDPOINT) {
          const combined = createCase(endpoint, family, "multiple", "body", "pairwise · " + first.parameter.name + " null × " + second.parameter.name + " boundary", undefined, "Pairwise interaction between two endpoint parameters.", "replace", destructive);
          combined.mutations = [
            { location: first.parameter.location, parameter: first.parameter.name, value: firstNull.value },
            { location: second.parameter.location, parameter: second.parameter.name, value: secondEdge.value },
          ];
          cases.push(combined);
        }
        if (firstEdge && secondNull && cases.length < MAX_CASES_PER_ENDPOINT) {
          const combined = createCase(endpoint, family, "multiple", "body", "pairwise · " + first.parameter.name + " boundary × " + second.parameter.name + " null", undefined, "Pairwise interaction between two endpoint parameters.", "replace", destructive);
          combined.mutations = [
            { location: first.parameter.location, parameter: first.parameter.name, value: firstEdge.value },
            { location: second.parameter.location, parameter: second.parameter.name, value: secondNull.value },
          ];
          cases.push(combined);
        }
      }
    }
  }

  if (family === "authz" || family === "idor") {
    const dimensions: Array<{ parameter: EndpointDefinition["parameters"][number]; variants: ParameterVariant[] }> = [];
    endpoint.parameters.forEach((parameter) => {
      if (cases.length >= MAX_CASES_PER_ENDPOINT) return;
      if (parameter.location !== "path" && parameter.location !== "query") return;
      if (["guid", "int"].indexOf(parameter.type) < 0 && !/id$/i.test(parameter.name)) return;
      const variants = idVariants(parameter.name, config);
      if (variants.length > 0) dimensions.push({ parameter: parameter, variants: variants });
      variants.forEach((variant) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, parameter.name, parameter.location, variant.label, variant.value, "Uses only identifiers explicitly present in the selected test config. Verify ownership and effective access with the second account.", "replace", destructive));
      });
    });
    for (let left = 0; left < dimensions.length && cases.length < MAX_CASES_PER_ENDPOINT; left++) {
      for (let right = left + 1; right < dimensions.length && cases.length < MAX_CASES_PER_ENDPOINT; right++) {
        const first = dimensions[left];
        const second = dimensions[right];
        first.variants.slice(0, 6).forEach((firstValue) => second.variants.slice(0, 6).forEach((secondValue) => {
          if (cases.length >= MAX_CASES_PER_ENDPOINT) return;
          const combined = createCase(endpoint, family, "multiple", "body", "pairwise · " + first.parameter.name + " × " + second.parameter.name, undefined, "Combines two configured resource identifiers to check object-level authorization. Verify the target ownership before execution.", "replace", destructive);
          combined.mutations = [
            { location: first.parameter.location, parameter: first.parameter.name, value: firstValue.value },
            { location: second.parameter.location, parameter: second.parameter.name, value: secondValue.value },
          ];
          cases.push(combined);
        }));
      }
    }
  }

  if (family === "business-logic") {
    const dimensions = endpoint.parameters.filter((parameter) => parameter.allowedValues && parameter.allowedValues.length > 0);
    dimensions.forEach((parameter) => {
      (parameter.allowedValues || []).forEach((value) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, parameter.name, parameter.location, "declared enum · " + String(value), value, "Exercises a documented endpoint parameter value."));
      });
    });
    for (let left = 0; left < dimensions.length && cases.length < MAX_CASES_PER_ENDPOINT; left++) {
      for (let right = left + 1; right < dimensions.length && cases.length < MAX_CASES_PER_ENDPOINT; right++) {
        dimensions[left].allowedValues?.slice(0, 6).forEach((leftValue) => dimensions[right].allowedValues?.slice(0, 6).forEach((rightValue) => {
          if (cases.length >= MAX_CASES_PER_ENDPOINT) return;
          const combined = createCase(endpoint, family, "multiple", "body", "declared enum pair · " + dimensions[left].name + " × " + dimensions[right].name, undefined, "Combines declared enum values to test parameter interaction.");
          combined.mutations = [
            { location: dimensions[left].location, parameter: dimensions[left].name, value: leftValue },
            { location: dimensions[right].location, parameter: dimensions[right].name, value: rightValue },
          ];
          cases.push(combined);
        }));
      }
    }
    // Probe server-side allow-listing for undocumented properties. These are
    // hypotheses only; verify persisted state and the second account's access.
    if (endpoint.requestBody && endpoint.requestBody.example && typeof endpoint.requestBody.example === "object" && !Array.isArray(endpoint.requestBody.example)) {
      const candidates: Array<{ name: string; value: unknown }> = [
        { name: "isAdmin", value: true }, { name: "role", value: "Owner" }, { name: "permissions", value: "FullControl" },
        { name: "tenantId", value: config.environment.tenantId }, { name: "siteId", value: config.environment.siteId },
        { name: "userId", value: config.principals.users.secondary.id }, { name: "principalId", value: config.principals.users.secondary.id },
        { name: "AuthorId", value: config.principals.users.primary.id }, { name: "EditorId", value: config.principals.users.secondary.id },
        { name: "Created", value: "2000-01-01T00:00:00Z" }, { name: "Modified", value: "2099-12-31T23:59:59Z" },
        { name: "_ModerationStatus", value: 0 }, { name: "OData__UIVersionString", value: "999.0" },
        { name: "HasUniqueRoleAssignments", value: true }, { name: "FileSystemObjectType", value: 1 }, { name: "isSiteAdmin", value: true },
        { name: "allowAnonymousAccess", value: true }, { name: "inheritPermissions", value: false }, { name: "scope", value: "anonymous" },
        { name: "linkKind", value: "anonymous" }, { name: "ownerId", value: config.principals.users.secondary.id },
        { name: "__metadata", value: "SP.Data.SecurityProbeListItem" },
      ];
      candidates.forEach((candidate) => {
        const name = candidate.name;
        if (cases.length >= MAX_CASES_PER_ENDPOINT || endpoint.requestBody?.protectedProperties?.indexOf(name) !== -1) return;
        if (candidate.value !== undefined && candidate.value !== "") cases.push(createCase(endpoint, family, name, "body", "unexpected property · " + name, candidate.value, "Hypothesis test for hidden-property mass assignment or business-logic changes. Acceptance alone is not a finding; verify persisted state and authorization.", "replace", true));
      });
    }
    ["debug", "isAdmin", "role", "tenantId", "siteId", "userId", "includeDeleted", "bypassCache", "bypassAuthorization", "skipPermissionCheck", "includeHidden", "includeDrafts", "includeVersions", "scope", "allowAnonymousAccess"].forEach((name) => {
      if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, name, "query", "unexpected query parameter · " + name, "controlled-security-probe", "Checks whether an undocumented query flag changes filtering, identity scope, sharing behavior, or authorization. Compare actual objects and both accounts' access.", "replace"));
    });
  }

  if ((family === "authz" || family === "idor") && config.safety.crossTenantTargets) {
    config.safety.crossTenantTargets.forEach((target) => {
      if (cases.length >= MAX_CASES_PER_ENDPOINT) return;
      const idParams = endpoint.parameters.filter((parameter) => parameter.location === "path" || parameter.location === "query")
        .filter((parameter) => Object.prototype.hasOwnProperty.call(target.identifiers, parameter.name));
      // Do not invent a foreign tenant identifier: an explicit matching fixture ID is required.
      if (idParams.length === 0 || endpoint.family === "microsoft-graph") return;
      const first = idParams[0];
      const generated = createCase(endpoint, family, first.name, first.location, "cross-tenant boundary · " + target.name, target.identifiers[first.name], "Uses only an explicitly configured test-tenant target and identifier. Confirm both tenants are authorized and compare the same operation using each browser account.", "replace", isMutationEndpoint(endpoint));
      generated.siteUrlOverride = target.siteUrl;
      generated.mutations = idParams.slice(1).map((parameter) => ({ location: parameter.location, parameter: parameter.name, value: target.identifiers[parameter.name] }));
      cases.push(generated);
    });
  }

  if (family === "mass-assignment" && endpoint.requestBody) {
    (endpoint.requestBody.protectedProperties || []).forEach((name) => {
      if (cases.length < MAX_CASES_PER_ENDPOINT) {
        cases.push(createCase(endpoint, family, "body", "body", "attempt protected property · " + name, { [name]: "controlled-security-probe" }, "Tests a property explicitly marked protected by endpoint metadata.", "replace", true));
      }
    });
  }

  if (family === "path-traversal") {
    endpoint.parameters.filter((parameter) => parameter.location === "path" && (parameter.type === "path" || /path|file|folder|url/i.test(parameter.name))).forEach((parameter) => {
      ["../security-probe", "%2e%2e%2fsecurity-probe", "%252e%252e%252fsecurity-probe", "..\\security-probe", "/security-probe"].forEach((value) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, parameter.name, "path", "path variant · " + value, value, "Checks path normalization at this endpoint without targeting files outside the disposable fixture.", "replace", destructive));
      });
    });
  }

  if (family === "ssrf") {
    const urlParameters = endpoint.parameters.filter((parameter) => /url|uri|source|destination|target|image|thumbnail|redirect/i.test(parameter.name));
    urlParameters.forEach((parameter) => {
      ["https://127.0.0.1/", "https://localhost/", "https://example.invalid/"].forEach((value) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, parameter.name, parameter.location, "controlled URL · " + value, value, "Tests server-side URL handling with loopback or reserved example targets; no external callback is used.", "replace", destructive));
      });
    });
  }

  if (family === "xss") {
    endpoint.parameters.filter((parameter) => parameter.location === "body" && ["string", "json"].indexOf(parameter.type) >= 0).forEach((parameter) => {
      ["<b>security-probe</b>", "&lt;b&gt;security-probe&lt;/b&gt;", "\"'><security-probe>", "Prueba-üñîçødé-你好-😀"].forEach((value) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, parameter.name, "body", "rendering probe · " + value, value, "Uses inert markup and Unicode to inspect output encoding; it contains no script or event handler.", "replace", destructive));
      });
    });
  }

  if (family === "sqli") {
    endpoint.parameters.filter((parameter) => parameter.location === "query" || parameter.location === "body").forEach((parameter) => {
      ["' OR '1'='1", "' AND '1'='2", "\" OR \"1\"=\"1", "'--", "'/*probe*/OR/*probe*/'1'='1", "%27%20OR%20%271%27%3D%271", "'' OR 1=1", "1) OR (1=1"].forEach((value) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, parameter.name, parameter.location, "inert syntax probe · " + value, value, "Checks whether this parameter reaches a custom SQL sink. SharePoint's own data store is not assumed to be a SQL sink; no time-delay or stacked destructive query is used.", "replace", destructive));
      });
    });
  }

  if (family === "odata-injection") {
    endpoint.parameters.filter((parameter) => parameter.location === "query" && /filter|select|expand|orderby|top|skip|query/i.test(parameter.name)).forEach((parameter) => {
      ["1 eq 1", "'", "(Id eq 1)", "__unknown_field__ eq null", "0"].forEach((value) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, parameter.name, "query", "OData variant · " + value, value, "Checks OData parser and returned-row boundaries; compare actual records, not only the response status.", "replace", destructive));
      });
    });
  }

  if (family === "method-confusion") {
    ["GET", "POST", "PUT", "PATCH", "DELETE", "MERGE"].forEach((method) => {
      if (cases.length < MAX_CASES_PER_ENDPOINT) {
        cases.push(createCase(endpoint, family, "X-HTTP-Method", "header", "method override · " + method, method, "Probe method override handling; verify resulting resource state and authorization.", "replace", true));
      }
    });
    if (endpoint.requestBody) {
      ["application/json;odata=verbose", "application/json;odata=minimalmetadata", "application/json;odata=nometadata", "text/plain"].forEach((contentType) => {
        if (cases.length < MAX_CASES_PER_ENDPOINT) cases.push(createCase(endpoint, family, "Content-Type", "header", "content type · " + contentType, contentType, "Checks parser and authorization consistency across content types.", "replace", true));
      });
    }
  }

  return cases;
}
