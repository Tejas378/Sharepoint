import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Label,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
  TextField,
} from "@fluentui/react";
import { loadTestConfig, assertAllowedUrl } from "../../config/loader";
import { buildRequest } from "../../research/request";
import { listEndpoints, getEndpoint } from "../../research/registry";
import { RequestExecutor } from "../../research/executor";
import { DigestProvider, parseContextInfoResponse } from "../../research/digest";
import { createGraphTransport } from "../../research/graph-transport";
import { EndpointDefinition, ExecutionRecord, HttpMethod, HttpRequestSpec, Payload, PayloadFamily, PayloadGenerationContext, Transport } from "../../research/types";
import { generatePayloads, PAYLOAD_FAMILIES } from "../../research/payloads";
import ConfigSetupAssistant from "./ConfigSetupAssistant";
import "./ResearchApiPanel.scss";

interface ResearchApiPanelProps {
  context: WebPartContext;
}

const endpointOptions: IDropdownOption[] = listEndpoints().map((endpoint) => ({
  key: endpoint.id,
  text: endpoint.id + "  ·  " + endpoint.module + "  ·  " + endpoint.operation,
}));
const moduleNames: string[] = Array.from(new Set(listEndpoints().map((endpoint) => endpoint.module))).sort();

function configuredPath(name: string, config: ReturnType<typeof loadTestConfig>, endpoint?: EndpointDefinition): string {
  if (name === "listId") return config.resources.lists.primary.id;
  if (name === "itemId") return String(config.resources.items.primaryItemId || "");
  if (name === "folderId") return String(config.resources.folders.primaryFolderId || "");
  if (name === "viewId" || name === "View" || name === "formId") return config.resources.views.viewId;
  if (name === "contentTypeId") return config.resources.contentTypes.customContentTypeId;
  if (name === "fieldId") return config.resources.fields.titleFieldId;
  if (name === "groupId") {
    if (endpoint && endpoint.module === "taxonomy") return config.taxonomy.termGroupId;
    if (endpoint && endpoint.id === "lists.relatedFields") return config.resources.fields.relatedFieldGroupId || "";
    return String(config.principals.groups.primaryGroupId || "");
  }
  if (name === "setId") return config.taxonomy.termSetId;
  if (name === "termStoreId") return config.taxonomy.termStoreId;
  if (name === "termGroupId") return config.taxonomy.termGroupId;
  if (name === "termSetId") return config.taxonomy.termSetId;
  if (name === "termId") return config.taxonomy.termId;
  if (name === "commentId") return String((config.testData.constants as any).entityIds.commentId || "");
  if (name === "recycleBinItemId") return String((config.testData.constants as any).entityIds.recycleBinItemId || "");
  if (name === "siteId") {
    const env = config.environment;
    if (!env.siteCollectionId || !env.webId) return "";
    const hostname = safeHostname(env.siteUrl);
    return hostname ? [hostname, env.siteCollectionId, env.webId].join(",") : "";
  }
  if (name === "hostname") return safeHostname(config.environment.siteUrl);
  return "";
}

function safeHostname(url: string): string {
  try { return new URL(url).hostname; }
  catch (_error) { return ""; }
}

interface ConfiguredIdVariant { label: string; value: string; }

function configuredIdVariants(name: string, config: ReturnType<typeof loadTestConfig>): ConfiguredIdVariant[] {
  const candidates: ConfiguredIdVariant[] = [];
  const add = (label: string, value: unknown): void => {
    const text = value === undefined || value === null ? "" : String(value).trim();
    if (text && text !== "0" && text !== "-1") candidates.push({ label, value: text });
  };
  const resources = config.resources as any;
  const principals = config.principals as any;
  if (name === "listId" || name === "listID") {
    Object.keys(resources.lists || {}).forEach((key) => add("resources.lists." + key + ".id", resources.lists[key].id));
  } else if (name === "itemId" || name === "ID") {
    const items = resources.items || {};
    add("resources.items.primaryItemId", items.primaryItemId);
    add("resources.items.secondaryItemId", items.secondaryItemId);
    add("resources.items.nonExistentItemId", items.nonExistentItemId);
    (items.boundaryItemIds || []).forEach((value: unknown, index: number) => add("resources.items.boundaryItemIds[" + index + "]", value));
  } else if (name === "principalId" || name === "principalid") {
    (principals.principalIds || []).forEach((value: unknown, index: number) => add("principals.principalIds[" + index + "]", value));
    Object.keys(principals.users || {}).forEach((key) => add("principals.users." + key + ".id", principals.users[key].id));
    Object.keys(principals.groups || {}).forEach((key) => {
      if (/id$/i.test(key)) add("principals.groups." + key, principals.groups[key]);
      (principals.groups[key] || []).forEach && (principals.groups[key] || []).forEach((value: unknown, index: number) => add("principals.groups." + key + "[" + index + "]", value));
    });
  } else if (name === "userId" || name === "userID") {
    Object.keys(principals.users || {}).forEach((key) => add("principals.users." + key + ".id", principals.users[key].id));
  } else if (name === "groupId" || name === "groupID") {
    Object.keys(principals.groups || {}).forEach((key) => {
      if (/id$/i.test(key)) add("principals.groups." + key, principals.groups[key]);
      const list = principals.groups[key];
      if (Array.isArray(list)) list.forEach((value: unknown, index: number) => add("principals.groups." + key + "[" + index + "]", value));
    });
  } else if (name === "roleDefId" || name === "roleDefinitionId" || name === "roleDefID") {
    Object.keys(principals.roleDefinitions || {}).forEach((key) => {
      const value = principals.roleDefinitions[key];
      if (Array.isArray(value)) value.forEach((id: unknown, index: number) => add("principals.roleDefinitions." + key + "[" + index + "]", id));
      else add("principals.roleDefinitions." + key, value);
    });
  }
  const unique: ConfiguredIdVariant[] = [];
  const seen: { [value: string]: boolean } = {};
  candidates.forEach((candidate) => {
    if (!seen[candidate.value]) { seen[candidate.value] = true; unique.push(candidate); }
  });
  return unique;
}

function initialValues(endpoint: EndpointDefinition, config: ReturnType<typeof loadTestConfig>): Record<string, string> {
  const values: Record<string, string> = {};
  endpoint.parameters.forEach((parameter) => {
    if (parameter.location !== "path" && parameter.location !== "query") return;
    const fromConfig = configuredPath(parameter.name, config, endpoint);
  const requiresTenantConfig = ["listId", "siteId", "viewId", "View", "formId", "contentTypeId", "fieldId", "groupId", "setId", "termStoreId", "termGroupId", "termSetId", "termId", "commentId", "recycleBinItemId"].indexOf(parameter.name) >= 0;
    values[parameter.location + ":" + parameter.name] = fromConfig || (requiresTenantConfig ? "" :
      (parameter.example === undefined ? "" : String(parameter.example)));
  });
  return values;
}

function initialBody(endpoint: EndpointDefinition): string {
  if (!endpoint.requestBody || endpoint.requestBody.example === undefined) return "";
  return typeof endpoint.requestBody.example === "string"
    ? endpoint.requestBody.example
    : JSON.stringify(endpoint.requestBody.example, null, 2);
}

function payloadContext(endpoint: EndpointDefinition, family: PayloadFamily, bodyText: string): PayloadGenerationContext {
  const params = endpoint.parameters;
  const has = (name: string, location?: string): boolean => params.some((parameter) => parameter.name === name && (!location || parameter.location === location));
  const find = (names: string[], location: string): string => {
    for (let i = 0; i < names.length; i++) if (has(names[i], location)) return names[i];
    return "";
  };
  let parameter = "";
  let location: PayloadGenerationContext["location"] = "body";
  if (family === "idor") { location = "path"; parameter = find(["itemId", "filePath", "folderPath", "listId", "groupId", "userId"], "path") || "itemId"; }
  else if (family === "path-traversal") { location = "path"; parameter = find(["filePath", "folderPath", "listPath", "serverRelativePath"], "path") || "filePath"; }
  else if (family === "odata-injection") { location = "query"; parameter = find(["$filter", "queryText", "$select", "$orderby", "$top"], "query") || "$filter"; }
  else if (family === "authz") { location = "path"; parameter = find(["principalId", "groupId", "userId", "itemId", "listId"], "path") || "principalId"; }
  else if (family === "method-confusion") { location = "header"; parameter = "X-HTTP-Method"; }
  else if (family === "ssrf") { location = endpoint.requestBody ? "body" : "query"; parameter = endpoint.requestBody ? "url" : (find(["url", "queryText"], "query") || "url"); }
  else if (family === "business-logic" || family === "mass-assignment") { location = "body"; parameter = "body"; }
  else if (endpoint.requestBody) {
    location = "body";
    parameter = endpoint.module === "comments" ? "text" : endpoint.module === "search" ? "Querytext" : "Title";
  } else if (has("queryText", "path")) { location = "path"; parameter = "queryText"; }
  else if (has("queryText", "query")) { location = "query"; parameter = "queryText"; }
  else if (params.some((p) => p.location === "path")) { location = "path"; parameter = params.filter((p) => p.location === "path")[0].name; }
  else { location = "query"; parameter = find(["$filter", "FieldInternalName", "queryText"], "query") || "$filter"; }

  let baselineValue: unknown;
  if (location === "body") {
    try { baselineValue = bodyText.trim() ? JSON.parse(bodyText) : endpoint.requestBody && endpoint.requestBody.example; }
    catch (_error) { baselineValue = bodyText; }
  } else {
    const definition = params.find((p) => p.name === parameter && p.location === location);
    baselineValue = definition && definition.example;
  }
  return { parameter, location, resourceType: endpoint.module === "files" ? "file" : endpoint.module === "folders" ? "folder" : endpoint.module === "attachments" ? "attachment" : endpoint.module === "users" ? "user" : endpoint.module === "groups" ? "group" : endpoint.module === "search" ? "search" : endpoint.module === "items" ? "item" : endpoint.module === "lists" || endpoint.module === "libraries" ? "list" : "unknown", baselineValue };
}

function applyPayloadToInput(input: { path: Record<string, string | number>; query: Record<string, string | number | boolean>; headers: Record<string, string>; body: unknown }, payload: Payload, endpoint: EndpointDefinition): void {
  if (payload.location === "path" || payload.location === "query") {
    const target = payload.location === "path" ? input.path : input.query;
    let name = payload.parameter;
    if (!(name in target)) {
      const candidate = endpoint.parameters.find((parameter) => parameter.location === payload.location);
      if (candidate) name = candidate.name;
    }
    if (name) target[name] = payload.value as string | number | boolean;
    return;
  }
  if (payload.location === "header") {
    input.headers[payload.parameter] = String(payload.value);
    return;
  }
  if (payload.parameter === "body") {
    if (payload.value && typeof payload.value === "object") {
      const body = Object.assign({}, input.body && typeof input.body === "object" ? input.body : {}) as Record<string, unknown>;
      const values = payload.value as Record<string, unknown>;
      if (Array.isArray(body.formValues)) {
        const formValues = (body.formValues as Array<Record<string, unknown>>).slice();
        Object.keys(values).forEach((fieldName) => formValues.push({ FieldName: fieldName, FieldValue: values[fieldName] }));
        body.formValues = formValues;
      } else if (body.request && typeof body.request === "object" && values.request && typeof values.request === "object") {
        body.request = Object.assign({}, body.request, values.request);
      } else if (body.creationInfo && typeof body.creationInfo === "object") {
        body.creationInfo = Object.assign({}, body.creationInfo, values);
      } else {
        Object.assign(body, values);
      }
      input.body = body;
    } else input.body = payload.value;
    return;
  }
  const body = input.body && typeof input.body === "object" ? Object.assign({}, input.body) as Record<string, unknown> : {};
  if (payload.value && typeof payload.value === "object" && !Array.isArray(payload.value)) {
    const values = payload.value as Record<string, unknown>;
    if (Array.isArray(body.formValues)) {
      const formValues = (body.formValues as Array<Record<string, unknown>>).slice();
      Object.keys(values).forEach((fieldName) => formValues.push({ FieldName: fieldName, FieldValue: values[fieldName] }));
      body.formValues = formValues;
    } else if (body.request && typeof body.request === "object" && values.request && typeof values.request === "object") {
      body.request = Object.assign({}, body.request, values.request);
    } else Object.assign(body, values);
  } else {
    let property = payload.parameter;
    if (!property || property === "Title") property = endpoint.module === "comments" ? "text" : endpoint.module === "search" ? "Querytext" : "Title";
    if (endpoint.module === "search" && body.request && typeof body.request === "object") {
      body.request = Object.assign({}, body.request, { [property]: payload.value });
    } else if (Array.isArray(body.formValues)) {
      const formValues = (body.formValues as Array<Record<string, unknown>>).slice();
      formValues.push({ FieldName: property, FieldValue: payload.value });
      body.formValues = formValues;
    } else body[property] = payload.value;
  }
  input.body = body;
}

function makeDigestProvider(siteUrl: string, timeoutMs: number): DigestProvider {
  return new DigestProvider({
    fetchDigest: async () => {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), timeoutMs);
      try {
        const response = await fetch(siteUrl.replace(/\/$/, "") + "/_api/contextinfo", {
          method: "POST",
          credentials: "same-origin",
          headers: { Accept: "application/json;odata=nometadata" },
          signal: controller.signal,
        });
        const body = await response.json();
        if (!response.ok) throw new Error("Could not get SharePoint request digest (HTTP " + response.status + ").");
        return parseContextInfoResponse(body);
      } finally {
        clearTimeout(timeout);
      }
    },
  });
}

function makeSharePointTransport(config: ReturnType<typeof loadTestConfig>): Transport {
  return async (spec: HttpRequestSpec) => {
    assertAllowedUrl(spec.url, config);
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), config.safety.requestTimeoutMs);
    try {
      const response = await fetch(spec.url, {
        method: spec.method,
        credentials: "same-origin",
        headers: spec.headers,
        body: spec.body,
        signal: controller.signal,
      });
      const headers: Record<string, string> = {};
      response.headers.forEach((value, name) => { headers[name.toLowerCase()] = value; });
      return { status: response.status, statusText: response.statusText, headers, bodyText: await response.text() };
    } finally {
      clearTimeout(timeout);
    }
  };
}

export default function ResearchApiPanel(props: ResearchApiPanelProps): React.ReactElement {
  const config = React.useMemo(() => loadTestConfig(), []);
  const [endpointId, setEndpointId] = React.useState<string>(endpointOptions[0].key as string);
  const [selectedModule, setSelectedModule] = React.useState<string>(getEndpoint(endpointOptions[0].key as string).module);
  const [payloadFamily, setPayloadFamily] = React.useState<PayloadFamily | "">("");
  const [payloadId, setPayloadId] = React.useState<string>("");
  const endpoint = React.useMemo(() => getEndpoint(endpointId), [endpointId]);
  const [values, setValues] = React.useState<Record<string, string>>(() => initialValues(getEndpoint(endpointOptions[0].key as string), config));
  const [method, setMethod] = React.useState<HttpMethod>(endpoint.methods[0] || "GET");
  const [bodyText, setBodyText] = React.useState<string>(initialBody(endpoint));
  const [headers, setHeaders] = React.useState<Record<string, string>>({});
  const [busy, setBusy] = React.useState<boolean>(false);
  const [record, setRecord] = React.useState<ExecutionRecord | null>(null);
  const [notice, setNotice] = React.useState<string>("");
  const [builtRequest, setBuiltRequest] = React.useState<string>("");
  const [sweepResults, setSweepResults] = React.useState<string[]>([]);
  const executionCount = React.useRef<number>(0);
  const lastExecutionAt = React.useRef<number>(0);
  const payloadFamilyOptions: IDropdownOption[] = endpoint.payloadFamilies.map((family) => ({ key: family, text: PAYLOAD_FAMILIES[family].title }));
  const payloadCases: Payload[] = React.useMemo(
    () => payloadFamily ? generatePayloads(payloadFamily, payloadContext(endpoint, payloadFamily, bodyText)) : [],
    [endpointId, payloadFamily, bodyText]
  );
  const selectedPayload = payloadCases.find((candidate) => candidate.id === payloadId);
  const payloadCaseOptions: IDropdownOption[] = payloadCases.map((candidate) => ({ key: candidate.id, text: (candidate.control ? "Control · " : candidate.classification + " · ") + candidate.name }));

  const changeEndpoint = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;
    const selected = getEndpoint(String(option.key));
    setEndpointId(selected.id);
    setSelectedModule(selected.module);
    setPayloadFamily("");
    setPayloadId("");
    setValues(initialValues(selected, config));
    setMethod(selected.methods[0] || "GET");
    setBodyText(initialBody(selected));
    setHeaders({});
    setRecord(null);
    setSweepResults([]);
    setBuiltRequest("");
    setNotice("");
  };

  const selectModule = (moduleName: string): void => {
    const firstEndpoint = listEndpoints().find((candidate) => candidate.module === moduleName);
    if (!firstEndpoint) return;
    setSelectedModule(moduleName);
    setEndpointId(firstEndpoint.id);
    setPayloadFamily("");
    setPayloadId("");
    setValues(initialValues(firstEndpoint, config));
    setMethod(firstEndpoint.methods[0] || "GET");
    setBodyText(initialBody(firstEndpoint));
    setHeaders({});
    setRecord(null);
    setSweepResults([]);
    setBuiltRequest("");
    setNotice("");
  };

  const updateValue = (key: string, value?: string): void => setValues((current) => ({ ...current, [key]: value || "" }));

  const build = (overrides: Record<string, string> = {}): HttpRequestSpec => {
    const path: Record<string, string | number> = {};
    const query: Record<string, string | number | boolean> = {};
    const requestHeaders: Record<string, string> = {};
    endpoint.parameters.forEach((parameter) => {
      if (parameter.location !== "path" && parameter.location !== "query") return;
      const key = parameter.location + ":" + parameter.name;
      const raw = overrides[key] !== undefined ? overrides[key] : (values[key] || "");
      if (parameter.required && !raw.trim()) throw new Error("Enter the required " + parameter.location + " value: " + parameter.name + ". Use the Test config setup check above to discover tenant values.");
      if (!raw) return;
      if (parameter.location === "path") path[parameter.name] = raw;
      else query[parameter.name] = raw;
    });
    const parsedHeaders: Record<string, string> = {};
    endpoint.requiredHeaders?.forEach((name) => {
      const value = headers[name];
      if (value && !/^(authorization|cookie|x-requestdigest)$/i.test(name)) parsedHeaders[name] = value;
    });
    let body: unknown;
    if (bodyText.trim()) {
      const contentType = endpoint.requestBody?.contentType || "application/json";
      if (contentType.indexOf("application/json") === 0) {
        try { body = JSON.parse(bodyText); }
        catch { throw new Error("Request body must be valid JSON."); }
      } else {
        body = bodyText;
      }
    }
    const requestInput = { path, query, headers: requestHeaders, body };
    if (selectedPayload) applyPayloadToInput(requestInput, selectedPayload, endpoint);
    return buildRequest(endpoint.id, {
      siteUrl: config.environment.siteUrl,
      path: requestInput.path,
      query: requestInput.query,
      method,
      body: requestInput.body,
      headers: Object.assign(parsedHeaders, requestInput.headers),
      digestPlaceholder: true,
    });
  };

  const previewRequest = (): void => {
    try {
      const spec = build();
      setBuiltRequest(spec.method + " " + spec.url + "\n" + JSON.stringify(spec.headers, null, 2) + (spec.body ? "\n\n" + spec.body : ""));
      setNotice("");
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error));
    }
  };

  const execute = async (): Promise<void> => {
    setNotice("");
    setRecord(null);
    setBusy(true);
    try {
      if (endpoint.family !== "microsoft-graph" && new URL(config.environment.siteUrl).hostname.endsWith(".invalid")) {
        throw new Error("Set environment.siteUrl and safety.allowedHosts in src/config/test-config.json before executing requests.");
      }
      if (endpoint.destructive && !config.safety.allowDestructive) {
        throw new Error("This endpoint is marked destructive. Enable safety.allowDestructive in test-config.json before running it.");
      }
      if (endpoint.destructive && endpoint.payloadFamilies.indexOf("authz") >= 0 && !config.safety.allowPermissionChanges) {
        throw new Error("This endpoint may change permissions. Enable safety.allowPermissionChanges in test-config.json before running it.");
      }
      if (executionCount.current >= config.safety.maxRequestsPerRun) {
        throw new Error("The configured maxRequestsPerRun limit has been reached for this page session.");
      }
      const spec = build();
      if (spec.body && new Blob([spec.body]).size > config.safety.maxPayloadBytes) {
        throw new Error("Request body exceeds safety.maxPayloadBytes.");
      }
      if (spec.destructive && !window.confirm("This endpoint is marked destructive. Send " + spec.intendedMethod + " to " + spec.url + "?")) return;
      const waitMs = config.safety.minimumDelayBetweenRequestsMs - (Date.now() - lastExecutionAt.current);
      if (lastExecutionAt.current > 0 && waitMs > 0) {
        setNotice("Waiting " + Math.ceil(waitMs / 1000) + " seconds for the configured request interval…");
        await new Promise<void>((resolve) => setTimeout(resolve, waitMs));
      }
      setBuiltRequest(spec.method + " " + spec.url);

      let transport: Transport;
      let digest: DigestProvider | null = null;
      if (endpoint.family === "microsoft-graph") {
        const providerFactory = (props.context as any).aadTokenProviderFactory;
        if (!providerFactory || typeof providerFactory.getTokenProvider !== "function") {
          throw new Error("Microsoft Graph token provider is unavailable in this SharePoint context.");
        }
        transport = createGraphTransport(async () => {
          const provider = await providerFactory.getTokenProvider();
          return provider.getToken("https://graph.microsoft.com");
        });
      } else {
        transport = makeSharePointTransport(config);
        if (spec.requiresDigest) digest = makeDigestProvider(config.environment.siteUrl, config.safety.requestTimeoutMs);
      }
      const executor = new RequestExecutor(transport, digest, {
        maxTransientRetries: 0,
        maxDigestRetries: 1,
        previewLength: config.safety.maxPayloadBytes > 0 ? Math.min(config.safety.maxPayloadBytes, 12000) : 12000,
      });
      const result = await executor.execute(spec);
      executionCount.current += 1;
      lastExecutionAt.current = Date.now();
      setRecord(result);
      setNotice(result.error || (result.status >= 200 && result.status < 300 ? "Request completed." : "Request completed with HTTP " + result.status + "."));
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  };

  const executeConfiguredIdSweep = async (): Promise<void> => {
    setNotice("");
    setRecord(null);
    setSweepResults([]);
    setBusy(true);
    try {
      if (method !== "GET" || endpoint.destructive) throw new Error("Configured ID sweeps are limited to non-destructive GET requests.");
      if (selectedPayload) throw new Error("Clear the payload selection before running a configured ID sweep.");
      if (endpoint.family !== "microsoft-graph" && new URL(config.environment.siteUrl).hostname.endsWith(".invalid")) throw new Error("Set environment.siteUrl and safety.allowedHosts in src/config/test-config.json before executing requests.");
      const parameters = endpoint.parameters.filter((p) => p.location === "path" || p.location === "query");
      const dimensions = parameters.map((p) => ({ key: p.location + ":" + p.name, variants: configuredIdVariants(p.name, config) })).filter((d) => d.variants.length > 0);
      if (!dimensions.length) throw new Error("No multiple configured IDs found for this endpoint. Add authorized list, item, user, group, principal, or role IDs to test-config.json.");
      const plans: Array<{ label: string; overrides: Record<string, string> }> = [];
      const expand = (index: number, overrides: Record<string, string>, labels: string[]): void => {
        if (index >= dimensions.length) { plans.push({ label: labels.join(" · "), overrides }); return; }
        const dimension = dimensions[index];
        dimension.variants.forEach((variant) => expand(index + 1, Object.assign({}, overrides, { [dimension.key]: variant.value }), labels.concat([variant.label + "=" + variant.value])));
      };
      // Combine identifiers present on the same resource route (for example list × item × principal).
      // The configured request cap rejects oversized products before any request is sent.
      expand(0, {}, []);
      const seenPlans: { [key: string]: boolean } = {};
      const uniquePlans = plans.filter((plan) => {
        const signature = JSON.stringify(plan.overrides);
        if (seenPlans[signature]) return false;
        seenPlans[signature] = true;
        return true;
      });
      if (executionCount.current + uniquePlans.length > config.safety.maxRequestsPerRun) throw new Error("This sweep needs " + uniquePlans.length + " requests, exceeding the remaining maxRequestsPerRun allowance (" + (config.safety.maxRequestsPerRun - executionCount.current) + "). Reduce configured ID sets or increase the safety cap deliberately.");
      const summaries: string[] = [];
      let transport: Transport;
      if (endpoint.family === "microsoft-graph") {
        const providerFactory = (props.context as any).aadTokenProviderFactory;
        if (!providerFactory || typeof providerFactory.getTokenProvider !== "function") throw new Error("Microsoft Graph token provider is unavailable in this SharePoint context.");
        transport = createGraphTransport(async () => { const provider = await providerFactory.getTokenProvider(); return provider.getToken("https://graph.microsoft.com"); });
      } else transport = makeSharePointTransport(config);
      const executor = new RequestExecutor(transport, null, { maxTransientRetries: 0, maxDigestRetries: 0, previewLength: Math.min(config.safety.maxPayloadBytes, 12000) });
      setNotice("Running " + uniquePlans.length + " configured ID variants sequentially…");
      for (const plan of uniquePlans) {
        const spec = build(plan.overrides);
        if (spec.destructive || spec.method !== "GET") throw new Error("Safety check stopped a non-read-only sweep request.");
        if (endpoint.family !== "microsoft-graph") assertAllowedUrl(spec.url, config);
        const waitMs = config.safety.minimumDelayBetweenRequestsMs - (Date.now() - lastExecutionAt.current);
        if (lastExecutionAt.current > 0 && waitMs > 0) await new Promise<void>((resolve) => setTimeout(resolve, waitMs));
        const result = await executor.execute(spec);
        executionCount.current += 1;
        lastExecutionAt.current = Date.now();
        summaries.push(plan.label + " → HTTP " + result.status + " · " + result.responseClass + (result.error ? " · " + result.error : ""));
        setSweepResults(summaries.slice());
        setRecord(result);
        setBuiltRequest(spec.method + " " + spec.url);
        setNotice("Sweep " + summaries.length + "/" + uniquePlans.length + " complete.");
      }
      setNotice("Configured ID sweep complete: " + uniquePlans.length + " requests. Review each status below.");
      setSweepResults(summaries);
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error));
    } finally { setBusy(false); }
  };

  const pathAndQuery = endpoint.parameters.filter((parameter) => parameter.location === "path" || parameter.location === "query");
  const bodyRequired = endpoint.requestBody !== undefined;
  const methodOptions: IDropdownOption[] = endpoint.methods.map((value) => ({ key: value, text: value }));

  return (
    <>
    <ConfigSetupAssistant context={props.context} />
    <section className="researchApiPanel">
      <div className="researchApiPanel__heading">
        <h2>Research API runner</h2>
        <p>Run registry endpoints in the current SharePoint session. Requests use the configured tenant and the signed-in user’s permissions.</p>
      </div>
      {notice && <MessageBar messageBarType={record && record.status >= 200 && record.status < 300 ? MessageBarType.success : MessageBarType.warning}>{notice}</MessageBar>}
      <Stack tokens={{ childrenGap: 12 }}>
        <div className="researchApiPanel__modules" role="group" aria-label="SharePoint API modules">
          <Label>API modules · select a module to test</Label>
          <div className="researchApiPanel__moduleButtons">
            {moduleNames.map((moduleName) => {
              const count = endpointOptions.filter((option) => getEndpoint(String(option.key)).module === moduleName).length;
              const label = moduleName.replace(/[-_]/g, " ").replace(/\b\w/g, (letter) => letter.toUpperCase());
              return moduleName === selectedModule
                ? <PrimaryButton key={moduleName} text={label + " · " + count} aria-pressed={true} onClick={() => selectModule(moduleName)} disabled={busy} />
                : <DefaultButton key={moduleName} text={label + " · " + count} aria-pressed={false} onClick={() => selectModule(moduleName)} disabled={busy} />;
            })}
          </div>
        </div>
        <Dropdown label="Endpoint in selected module" options={endpointOptions.filter((option) => getEndpoint(String(option.key)).module === selectedModule)} selectedKey={endpointId} onChange={changeEndpoint} />
        <div className="researchApiPanel__meta"><b>{endpoint.family}</b><span>{endpoint.description}</span>{endpoint.destructive && <strong>Destructive operation</strong>}</div>
        <div className="researchApiPanel__payloads">
          <h3>Bug-family payload sets</h3>
          <p>Choose a family and one case. The selected case is applied to the request preview and only that single request is sent when you execute.</p>
          <div className="researchApiPanel__payloadSelectors">
            <Dropdown label="Bug type" placeholder="Select a payload family" options={payloadFamilyOptions} selectedKey={payloadFamily || undefined} onChange={(_e, option) => { setPayloadFamily(option ? String(option.key) as PayloadFamily : ""); setPayloadId(""); }} />
            <Dropdown label="Payload case" placeholder={payloadFamily ? "Select a case to apply" : "Select a bug type first"} options={payloadCaseOptions} selectedKey={payloadId || undefined} disabled={!payloadFamily} onChange={(_e, option) => setPayloadId(option ? String(option.key) : "")} />
          </div>
          {selectedPayload && <div className="researchApiPanel__payloadDetails"><b>{selectedPayload.classification}{endpoint.destructive ? " · write request" : " · request mutation"}</b><span>{selectedPayload.rationale}</span><span>Review for: {(selectedPayload.expectedObservations || []).join(" · ")}</span></div>}
        </div>
        <Dropdown label="HTTP method" options={methodOptions} selectedKey={method} onChange={(_e, option) => option && setMethod(String(option.key) as HttpMethod)} />
        <div className="researchApiPanel__fields">
          {pathAndQuery.map((parameter) => {
            const key = parameter.location + ":" + parameter.name;
            const variants = configuredIdVariants(parameter.name, config);
            const options: IDropdownOption[] = variants.map((variant) => ({ key: variant.value, text: variant.label + " · " + variant.value }));
            return <React.Fragment key={key}><TextField label={parameter.location + " · " + parameter.name} description={parameter.description} value={values[key] || ""} onChange={(_e, value) => updateValue(key, value)} required={parameter.required} />{variants.length > 0 && <Dropdown label={"Configured values · " + parameter.name} placeholder="Choose a configured ID" options={options} selectedKey={variants.some((variant) => variant.value === values[key]) ? values[key] : undefined} onChange={(_e, option) => option && updateValue(key, String(option.key))} />}</React.Fragment>;
          })}
        </div>
        {bodyRequired && <TextField label={"Request body (" + (endpoint.requestBody?.contentType || "application/json") + ")"} description={endpoint.requestBody?.description} multiline rows={8} value={bodyText} onChange={(_e, value) => setBodyText(value || "")} />}
        {(endpoint.requiredHeaders || []).filter((name) => !/^(authorization|cookie|x-requestdigest)$/i.test(name)).map((name) => <TextField key={name} label={"Header · " + name} value={headers[name] || ""} onChange={(_e, value) => setHeaders((current) => ({ ...current, [name]: value || "" }))} />)}
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Preview request" onClick={previewRequest} disabled={busy} />
          <PrimaryButton text="Execute API call" onClick={() => { void execute(); }} disabled={busy} />
          <DefaultButton text="Sweep configured IDs (read only)" onClick={() => { void executeConfiguredIdSweep(); }} disabled={busy || method !== "GET" || endpoint.destructive || !pathAndQuery.some((p) => configuredIdVariants(p.name, config).length > 0)} />
          {busy && <Spinner size={SpinnerSize.small} label="Executing" />}
        </Stack>
        {sweepResults.length > 0 && <div><Label>Configured ID sweep results · {sweepResults.length} requests</Label><pre className="researchApiPanel__output">{sweepResults.join("\n")}</pre></div>}
        {builtRequest && <div><Label>Request</Label><pre className="researchApiPanel__output">{builtRequest}</pre></div>}
        {record && <div><Label>Response · HTTP {record.status} · {record.responseClass} · {record.durationMs} ms</Label><pre className="researchApiPanel__output">{record.responsePreview || "(empty response body)"}</pre></div>}
      </Stack>
    </section>
    </>
  );
}
