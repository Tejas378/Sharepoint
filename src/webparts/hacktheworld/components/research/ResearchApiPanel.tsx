import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import {
  DefaultButton,
  IconButton,
  Checkbox,
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
import { loadTestConfig, assertAllowedUrl, saveTestConfigOverride, clearTestConfigOverride } from "../../config/loader";
import { buildRequest } from "../../research/request";
import { listEndpoints, getEndpoint } from "../../research/registry";
import { RequestExecutor } from "../../research/executor";
import { DigestProvider, parseContextInfoResponse } from "../../research/digest";
import { createGraphTransport } from "../../research/graph-transport";
import { EndpointDefinition, ExecutionRecord, HttpMethod, HttpRequestSpec, Payload, PayloadFamily, QueryValue, Transport } from "../../research/types";
import { generatePayloads, generateShareLinkMatrix, PAYLOAD_FAMILIES } from "../../research/payloads";
import { generateEndpointCases } from "../../research/endpointCases";
import { applyPayloadToRequestInput, buildCasePlan, CaseRequestInput, payloadContextFor } from "../../research/casePlan";
import { redactUrl } from "../../research/redact";
import { toRawHttp } from "../../research/http";
import ConfigSetupAssistant from "./ConfigSetupAssistant";
import "./ResearchApiPanel.scss";

interface ResearchApiPanelProps {
  context: WebPartContext;
}

const CONFIG_FILE_LABEL = "src/config/test-config.json";

const endpointOptions: IDropdownOption[] = listEndpoints().map((endpoint) => ({
  key: endpoint.id,
  text: endpoint.id + "  ·  " + endpoint.module + "  ·  " + endpoint.operation,
}));
const moduleNames: string[] = Array.from(new Set(listEndpoints().map((endpoint) => endpoint.module))).sort();
const moduleGroups: Array<{ label: string; modules: string[] }> = [
  { label: "Content and lists", modules: ["attachments", "content-types", "fields", "files", "folders", "items", "libraries", "lists", "recycle-bin", "related-items", "views"] },
  { label: "People and permissions", modules: ["approvals", "group-site-manager", "groups", "permissions", "profiles", "sharing", "users"] },
  { label: "Sites and pages", modules: ["app-catalog", "client-pages", "features", "hubsites", "navigation", "publishing-site-page-service", "site-designs", "site-scripts", "sites", "webparts", "webs"] },
  { label: "Search and integrations", modules: ["search", "webhooks"] },
  { label: "Administration and other", modules: ["batch", "changes", "comments", "configuration", "context", "regional-settings", "social", "taxonomy", "utilities"] },
];
const firstModuleName: string = moduleNames[0] || getEndpoint(String(endpointOptions[0].key)).module;
const firstModuleEndpoint: EndpointDefinition = listEndpoints().find((candidate) => candidate.module === firstModuleName) || getEndpoint(String(endpointOptions[0].key));

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

function toggleDropdownKey(current: string[], option?: IDropdownOption): string[] {
  if (!option) return current;
  const key = String(option.key);
  return option.selected ? Array.from(new Set(current.concat([key]))) : current.filter((value) => value !== key);
}

interface ConfiguredIdVariant { label: string; value: string; }

interface SharePointChoice { key: string; text: string; }

function resultArray<T>(json: any): T[] {
  if (json && json.d && Array.isArray(json.d.results)) return json.d.results as T[];
  if (json && Array.isArray(json.value)) return json.value as T[];
  return [];
}

async function getAllSharePointRows<T>(context: WebPartContext, url: string): Promise<T[]> {
  const rows: T[] = [];
  let nextUrl: string | undefined = url;
  while (nextUrl) {
    const json = await getSharePointJson(context, nextUrl);
    rows.push(...resultArray<T>(json));
    const next = json && json.d && json.d.__next
      ? json.d.__next
      : json && (json["@odata.nextLink"] || json["odata.nextLink"]);
    nextUrl = typeof next === "string" && next.length > 0 ? next : undefined;
  }
  return rows;
}

async function getSharePointJson(context: WebPartContext, url: string): Promise<any> {
  const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
    headers: { Accept: "application/json;odata=nometadata" },
  });
  if (!response.ok) throw new Error("HTTP " + response.status + " loading test choices from " + new URL(url).pathname + ".");
  return response.json();
}

function choice(key: unknown, text: unknown): SharePointChoice {
  return { key: String(key), text: String(text || key) + " · " + String(key) };
}

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
  const [endpointId, setEndpointId] = React.useState<string>(firstModuleEndpoint.id);
  const [endpointKeys, setEndpointKeys] = React.useState<string[]>([firstModuleEndpoint.id]);
  const [selectedModule, setSelectedModule] = React.useState<string>(firstModuleName);
  const [payloadFamily, setPayloadFamily] = React.useState<PayloadFamily | "">("");
  const [payloadFamilyKeys, setPayloadFamilyKeys] = React.useState<string[]>([]);
  const [payloadId, setPayloadId] = React.useState<string>("");
  const [payloadKeys, setPayloadKeys] = React.useState<string[]>([]);
  const endpoint = React.useMemo(() => getEndpoint(endpointId), [endpointId]);
  const [values, setValues] = React.useState<Record<string, string>>(() => initialValues(firstModuleEndpoint, config));
  const [method, setMethod] = React.useState<HttpMethod>(endpoint.methods[0] || "GET");
  const [methodKeys, setMethodKeys] = React.useState<string[]>([endpoint.methods[0] || "GET"]);
  const [discoveredChoices, setDiscoveredChoices] = React.useState<Record<string, SharePointChoice[]>>({});
  const [choiceNotice, setChoiceNotice] = React.useState<string>("");
  const [enabledQueryFilters, setEnabledQueryFilters] = React.useState<Record<string, boolean>>({});
  const [customItemIds, setCustomItemIds] = React.useState<string[]>([]);
  const [customItemIdsInput, setCustomItemIdsInput] = React.useState<string>("");
  const [bodyText, setBodyText] = React.useState<string>(initialBody(endpoint));
  const [headers, setHeaders] = React.useState<Record<string, string>>({});
  const [busy, setBusy] = React.useState<boolean>(false);
  const [record, setRecord] = React.useState<ExecutionRecord | null>(null);
  const [notice, setNotice] = React.useState<string>("");
  const [builtRequest, setBuiltRequest] = React.useState<string>("");
  const [sweepResults, setSweepResults] = React.useState<string[]>([]);
  const [configEditorOpen, setConfigEditorOpen] = React.useState<boolean>(false);
  const [configJson, setConfigJson] = React.useState<string>(() => JSON.stringify(config, null, 2));
  const [configNotice, setConfigNotice] = React.useState<string>("");
  const idVariantsFor = (name: string): ConfiguredIdVariant[] => {
    const configured = configuredIdVariants(name, config);
    if (name !== "itemId" && name !== "ID") return configured;
    const seen = new Set(configured.map((variant) => variant.value));
    return configured.concat(customItemIds.filter((id) => !seen.has(id)).map((id) => ({ label: "runtime item ID", value: id })));
  };
  const addCustomItemIds = (): void => {
    const incoming = customItemIdsInput.split(/[\s,;]+/).map((id) => id.trim()).filter(Boolean);
    setCustomItemIds((current) => Array.from(new Set(current.concat(incoming))));
    setCustomItemIdsInput("");
  };
  const payloadFamilyOptions: IDropdownOption[] = endpoint.payloadFamilies.map((family) => ({ key: family, text: PAYLOAD_FAMILIES[family].title }));
  const payloadCases: Payload[] = React.useMemo(() => {
    if (!payloadFamily) return [];
    const cases = generatePayloads(payloadFamily, payloadContextFor(endpoint, payloadFamily, bodyText))
      .concat(generateEndpointCases(endpoint, payloadFamily, config));
    if (payloadFamily === "business-logic" && /\.shareLink$/.test(endpoint.id)) {
      return cases.concat(generateShareLinkMatrix(
        config.principals.shareLinkKinds,
        config.principals.shareLinkRoleKinds
      ));
    }
    return cases;
  }, [endpointId, payloadFamily, bodyText, config]);
  const selectedPayload = payloadCases.find((candidate) => candidate.id === payloadId);
  const payloadCaseOptions: IDropdownOption[] = payloadCases.map((candidate) => ({ key: candidate.id, text: (candidate.control ? "Control · " : candidate.classification + " · ") + candidate.name }));

  React.useEffect(() => {
    let cancelled = false;
    const baseUrl = config.environment.tenantUrl.replace(/\/+$/, "");
    if (!baseUrl) return () => { cancelled = true; };
    const api = baseUrl + "/_api/web";
    const load = async (): Promise<void> => {
      try {
        const [lists, users, groups] = await Promise.all([
          getAllSharePointRows<any>(props.context, api + "/lists"),
          getAllSharePointRows<any>(props.context, api + "/siteusers"),
          getAllSharePointRows<any>(props.context, api + "/sitegroups"),
        ]);
        if (cancelled) return;
        setDiscoveredChoices((current) => Object.assign({}, current, {
          listId: lists.map((row) => choice(row.Id, row.Title)),
          listID: lists.map((row) => choice(row.Id, row.Title)),
          userId: users.map((row) => choice(row.Id, row.Title || row.LoginName)),
          userID: users.map((row) => choice(row.Id, row.Title || row.LoginName)),
          principalId: users.concat(groups).map((row) => choice(row.Id, row.Title || row.LoginName)),
          principalid: users.concat(groups).map((row) => choice(row.Id, row.Title || row.LoginName)),
          groupId: groups.map((row) => choice(row.Id, row.Title)),
          groupID: groups.map((row) => choice(row.Id, row.Title)),
        }));
        setChoiceNotice("");
      } catch (error) {
        if (!cancelled) setChoiceNotice("Could not load live list, user and group choices: " + (error instanceof Error ? error.message : String(error)));
      }
    };
    void load();
    return () => { cancelled = true; };
  }, [props.context, config.environment.tenantUrl]);

  React.useEffect(() => {
    let cancelled = false;
    const listId = values["path:listId"] || values["path:listID"];
    const baseUrl = config.environment.tenantUrl.replace(/\/+$/, "");
    if (!listId || !baseUrl) return () => { cancelled = true; };
    const escapedId = listId.replace(/[{}]/g, "");
    const listUrl = baseUrl + "/_api/web/lists(guid'" + encodeURIComponent(escapedId) + "')";
    setDiscoveredChoices((current) => Object.assign({}, current, { itemId: [], ID: [], viewId: [], View: [], formId: [], fieldId: [], contentTypeId: [] }));
    const load = async (): Promise<void> => {
      try {
        const [items, views, fields, contentTypes] = await Promise.all([
          getAllSharePointRows<any>(props.context, listUrl + "/items"),
          getAllSharePointRows<any>(props.context, listUrl + "/views"),
          getAllSharePointRows<any>(props.context, listUrl + "/fields"),
          getAllSharePointRows<any>(props.context, listUrl + "/contenttypes"),
        ]);
        if (cancelled) return;
        const itemChoices = items.map((row) => choice(row.Id, row.Title || "Item #" + row.Id));
        const fieldChoices = fields.map((row) => choice(row.Id, row.Title || row.InternalName));
        const viewChoices = views.map((row) => choice(row.Id, row.Title));
        const typeChoices = contentTypes.map((row) => choice(row.StringId, row.Name));
        setDiscoveredChoices((current) => Object.assign({}, current, {
          itemId: itemChoices, ID: itemChoices,
          viewId: viewChoices, View: viewChoices, formId: viewChoices,
          fieldId: fieldChoices,
          contentTypeId: typeChoices,
        }));
        setChoiceNotice("");
      } catch (error) {
        if (!cancelled) setChoiceNotice("Could not load choices for the selected list: " + (error instanceof Error ? error.message : String(error)));
      }
    };
    void load();
    return () => { cancelled = true; };
  }, [props.context, config.environment.tenantUrl, values["path:listId"], values["path:listID"]]);

  const changeEndpoint = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;
    const nextKeys = toggleDropdownKey(endpointKeys, option);
    setEndpointKeys(nextKeys);
    const selected = getEndpoint(String(option.selected ? option.key : nextKeys[nextKeys.length - 1] || endpointId));
    setEndpointId(selected.id);
    setSelectedModule(selected.module);
    setPayloadFamily("");
    setPayloadFamilyKeys([]);
    setPayloadId("");
    setPayloadKeys([]);
    setValues(initialValues(selected, config));
    setEnabledQueryFilters({});
    setMethod(selected.methods[0] || "GET");
    setMethodKeys([selected.methods[0] || "GET"]);
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
    setEndpointKeys([firstEndpoint.id]);
    setPayloadFamily("");
    setPayloadFamilyKeys([]);
    setPayloadId("");
    setPayloadKeys([]);
    setValues(initialValues(firstEndpoint, config));
    setEnabledQueryFilters({});
    setMethod(firstEndpoint.methods[0] || "GET");
    setMethodKeys([firstEndpoint.methods[0] || "GET"]);
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
    const query: Record<string, QueryValue> = {};
    const requestHeaders: Record<string, string> = {};
    endpoint.parameters.forEach((parameter) => {
      if (parameter.location !== "path" && parameter.location !== "query") return;
      const key = parameter.location + ":" + parameter.name;
      if (parameter.location === "query" && !parameter.required && !enabledQueryFilters[key]) return;
      const raw = overrides[key] !== undefined ? overrides[key] : (values[key] || "");
      const selectedOmission = selectedPayload && selectedPayload.action === "omit" && selectedPayload.location === parameter.location && selectedPayload.parameter === parameter.name;
      if (parameter.required && !raw.trim() && !selectedOmission) throw new Error("Enter the required " + parameter.location + " value: " + parameter.name + ". Use the Test config setup check above to discover tenant values.");
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
    const requestInput: CaseRequestInput = { path, query, headers: requestHeaders, body, method };
    if (selectedPayload) applyPayloadToRequestInput(requestInput, selectedPayload, endpoint);
    // A method-override payload rewrites the effective verb to POST and carries
    // the intended verb in X-HTTP-Method, exactly as the exported plan does.
    const effectiveMethod = requestInput.method || method;
    const spec = buildRequest(endpoint.id, {
      siteUrl: selectedPayload?.siteUrlOverride || config.environment.siteUrl,
      path: requestInput.path,
      query: requestInput.query,
      method: effectiveMethod,
      body: requestInput.body,
      headers: Object.assign(parsedHeaders, requestInput.headers),
      digestPlaceholder: true,
    });
    if (selectedPayload && selectedPayload.destructive) spec.destructive = true;
    if (["PUT", "PATCH", "DELETE", "MERGE"].indexOf(effectiveMethod) >= 0 ||
        (endpoint.methods.indexOf(effectiveMethod) < 0 && effectiveMethod === "POST")) spec.destructive = true;
    return spec;
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
      const spec = build();
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
      const dimensions = parameters.map((p) => ({ key: p.location + ":" + p.name, variants: idVariantsFor(p.name) })).filter((d) => d.variants.length > 0);
      if (!dimensions.length) throw new Error("No IDs are available for this endpoint. Add authorized IDs in the UI or test-config.json.");
      const plans: Array<{ label: string; overrides: Record<string, string> }> = [];
      const expand = (index: number, overrides: Record<string, string>, labels: string[]): void => {
        if (index >= dimensions.length) { plans.push({ label: labels.join(" · "), overrides }); return; }
        const dimension = dimensions[index];
        dimension.variants.forEach((variant) => expand(index + 1, Object.assign({}, overrides, { [dimension.key]: variant.value }), labels.concat([variant.label + "=" + variant.value])));
      };
      // Combine identifiers present on the same resource route (for example list x item x principal).
      // The configured request cap rejects oversized products before any request is sent.
      expand(0, {}, []);
      const seenPlans: { [key: string]: boolean } = {};
      const uniquePlans = plans.filter((plan) => {
        const signature = JSON.stringify(plan.overrides);
        if (seenPlans[signature]) return false;
        seenPlans[signature] = true;
        return true;
      });
      const summaries: string[] = [];
      let transport: Transport;
      if (endpoint.family === "microsoft-graph") {
        const providerFactory = (props.context as any).aadTokenProviderFactory;
        if (!providerFactory || typeof providerFactory.getTokenProvider !== "function") throw new Error("Microsoft Graph token provider is unavailable in this SharePoint context.");
        transport = createGraphTransport(async () => { const provider = await providerFactory.getTokenProvider(); return provider.getToken("https://graph.microsoft.com"); });
      } else transport = makeSharePointTransport(config);
      const executor = new RequestExecutor(transport, null, { maxTransientRetries: 0, maxDigestRetries: 0, previewLength: Math.min(config.safety.maxPayloadBytes, 12000) });
      setNotice("Running " + uniquePlans.length + " configured ID variants sequentially...");
      for (const plan of uniquePlans) {
        const spec = build(plan.overrides);
        if (spec.destructive || spec.method !== "GET") throw new Error("Safety check stopped a non-read-only sweep request.");
        if (endpoint.family !== "microsoft-graph") assertAllowedUrl(spec.url, config);
        const result = await executor.execute(spec);
        summaries.push(plan.label + " -> HTTP " + result.status + " · " + result.responseClass + (result.error ? " · " + result.error : ""));
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
  const controlledProbeMethods: HttpMethod[] = ["GET", "POST", "PUT", "PATCH", "DELETE", "MERGE", "HEAD", "OPTIONS"];
  const visibleMethods = payloadFamily === "method-confusion"
    ? Array.from(new Set(endpoint.methods.concat(controlledProbeMethods)))
    : endpoint.methods;
  const methodOptions: IDropdownOption[] = visibleMethods.map((value) => ({
    key: value,
    text: endpoint.methods.indexOf(value) >= 0 ? value : value + " · undeclared probe",
  }));
  const signedInUser = props.context.pageContext.user;
  const signedInAccount = signedInUser.loginName || signedInUser.email || signedInUser.displayName || "Unknown signed-in user";
  const currentAccountKeys = [signedInUser.loginName, signedInUser.email, signedInUser.displayName]
    .filter((value): value is string => !!value)
    .map((value) => (value.split("|").pop() || value).toLowerCase());
  const accountMatch = (configured: string): boolean => !!configured && currentAccountKeys.indexOf((configured.split("|").pop() || configured).toLowerCase()) >= 0;
  const copyConfigJson = async (): Promise<void> => {
    setConfigEditorOpen(true);
    try {
      await navigator.clipboard.writeText(configJson);
      setConfigNotice("Current config JSON copied. Browser-only edits are saved separately from the source config file.");
    } catch (_error) {
      setConfigNotice("Clipboard access is unavailable. The config JSON is shown below so you can copy it manually.");
    }
  };
  const applyConfigJson = (): void => {
    try {
      saveTestConfigOverride(configJson);
      window.location.reload();
    } catch (error) {
      setConfigNotice(error instanceof Error ? error.message : String(error));
    }
  };
  const resetConfigJson = (): void => {
    try {
      clearTestConfigOverride();
      window.location.reload();
    } catch (error) {
      setConfigNotice(error instanceof Error ? error.message : String(error));
    }
  };
  const copyResponse = async (): Promise<void> => {
    if (!record) return;
    const responseHeaders = Object.keys(record.responseHeaders).map((name) => name + ": " + record.responseHeaders[name]).join("\n");
    const responseText = "HTTP " + record.status + " · " + record.responseClass + "\n" + responseHeaders + "\n\n" + (record.responseBody || "");
    try {
      await navigator.clipboard.writeText(responseText);
    } catch (_error) {
      const textarea = document.createElement("textarea");
      textarea.value = responseText;
      textarea.style.position = "fixed";
      textarea.style.opacity = "0";
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand("copy");
      document.body.removeChild(textarea);
    }
    setNotice("Complete response copied to clipboard.");
  };
  const downloadParameterInventory = (): void => {
    const inventory = listEndpoints().map((item) => ({ endpointId: item.id, module: item.module, operation: item.operation, family: item.family, methods: item.methods, route: item.urlTemplate, auth: item.auth, destructive: item.destructive, parameters: item.parameters, requestBody: item.requestBody ? { contentType: item.requestBody.contentType, requiredProperties: item.requestBody.requiredProperties || [], protectedProperties: item.requestBody.protectedProperties || [], exampleProperties: item.requestBody.example && typeof item.requestBody.example === "object" ? Object.keys(item.requestBody.example as object) : [] } : null, payloadFamilies: item.payloadFamilies }));
    const blob = new Blob([JSON.stringify({ generatedAt: new Date().toISOString(), note: "Registry-declared parameters only; undocumented hidden parameters are hypotheses, not a complete server contract.", endpoints: inventory }, null, 2)], { type: "application/json" });
    const href = URL.createObjectURL(blob); const anchor = document.createElement("a"); anchor.href = href; anchor.download = "sharepoint-api-parameter-inventory.json"; anchor.click(); window.setTimeout(() => URL.revokeObjectURL(href), 1000);
  };
  /**
   * Exports the currently selected endpoint(s) as a Burp-ready plan, using the
   * values the operator typed into this form rather than only config defaults.
   * The plan is generated, never sent, and includes write cases for testing.
   */
  const downloadCasePlan = (): void => {
    try {
      const allTargets = endpointKeys.length > 0 ? endpointKeys : [endpointId];
      // Keep the synchronous plan build bounded so a large multi-select does
      // not block the browser tab.
      const targets = allTargets.slice(0, 10);
      const pathOverrides: Record<string, string | number> = {};
      const queryOverrides: Record<string, QueryValue> = {};
      endpoint.parameters.forEach((parameter) => {
        if (parameter.location !== "path" && parameter.location !== "query") return;
        const key = parameter.location + ":" + parameter.name;
        // Match the preview/execute path: optional OData projections only
        // travel when the operator ticked them on.
        if (parameter.location === "query" && !parameter.required && !enabledQueryFilters[key]) return;
        const raw = values[key];
        if (!raw) return;
        if (parameter.location === "path") {
          pathOverrides[parameter.name] = /^-?\d+$/.test(raw) ? Number(raw) : raw;
        } else {
          queryOverrides[parameter.name] = raw;
        }
      });
      let body: unknown;
      if (bodyText.trim()) {
        const contentType = endpoint.requestBody?.contentType || "application/json";
        body = contentType.indexOf("application/json") === 0 ? JSON.parse(bodyText) : bodyText;
      }
      const plan = buildCasePlan(targets, config, {
        families: payloadFamilyKeys.length > 0 ? payloadFamilyKeys as PayloadFamily[] : undefined,
        path: pathOverrides,
        query: queryOverrides,
        // The form body belongs to the currently displayed endpoint. Applying
        // it to a multi-endpoint export would fabricate bodies for other routes.
        body: targets.length === 1 ? body : undefined,
        includeDestructive: true,
        maxPerEndpoint: 60,
        maxTotal: 300,
      });
      const blocks: string[] = [
        "# SharePoint security case plan",
        "# Generated: " + plan.generatedAt,
        "# Target: " + redactUrl(plan.siteUrl),
        "# Endpoints: " + plan.totals.endpoints + " · cases: " + plan.totals.cases,
        "# NOTHING WAS SENT. Replace {{REQUEST_DIGEST}} in your authenticated session.",
      ];
      plan.notes.forEach((note) => blocks.push("# " + note));
      plan.endpoints.forEach((planned) => {
        planned.cases.forEach((item) => {
          blocks.push("");
          blocks.push("### " + item.caseId);
          blocks.push("# endpoint: " + item.endpointId);
          blocks.push("# family: " + item.family + " · source: " + item.source + " · classification: " + item.classification);
          blocks.push("# case: " + item.name);
          blocks.push("# rationale: " + item.rationale);
          item.expectedObservations.forEach((observation) => blocks.push("# observe: " + observation));
          if (item.destructive) blocks.push("# DESTRUCTIVE: this request may modify tenant data.");
          blocks.push("");
          blocks.push(toRawHttp(item.request));
        });
        planned.skipped.forEach((reason) => {
          blocks.push("# skipped " + planned.endpointId + ": " + reason);
        });
      });
      const blob = new Blob([blocks.join("\r\n") + "\r\n"], { type: "text/plain;charset=utf-8" });
      const href = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = href;
      anchor.download = "sharepoint-case-plan.http";
      anchor.click();
      window.setTimeout(() => URL.revokeObjectURL(href), 1000);
      setNotice(
        "Case plan exported: " + plan.totals.cases + " request(s) across " + plan.totals.endpoints +
          " endpoint(s). Nothing was sent." +
          (allTargets.length > targets.length
            ? " Only the first " + targets.length + " selected endpoints were included."
            : "")
      );
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error));
    }
  };
  const copyRawRequest = async (): Promise<void> => {
    try {
      const spec = build();
      const raw = toRawHttp(spec);
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(raw);
      } else {
        // SharePoint may host the web part in an iframe without clipboard
        // permission. Fall back to a hidden textarea + execCommand.
        const textarea = document.createElement("textarea");
        textarea.value = raw;
        textarea.style.position = "fixed";
        textarea.style.opacity = "0";
        document.body.appendChild(textarea);
        textarea.select();
        const copied = document.execCommand("copy");
        document.body.removeChild(textarea);
        if (!copied) throw new Error("Clipboard API is unavailable in this browser context.");
      }
      setNotice("Copied raw HTTP request (" + spec.method + "). Nothing was sent.");
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error));
    }
  };

  return (
    <>
    <section className="researchApiPanel">
      <div className="researchApiPanel__heading">
        <h2>Research API runner</h2>
        <p>Run registry endpoints in the current SharePoint session. Requests use the configured tenant and the signed-in user's permissions.</p>
      </div>
      <div className="researchApiPanel__configEditor">
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
          <Label>Test config · {CONFIG_FILE_LABEL}</Label>
          <DefaultButton text="Copy config JSON" onClick={() => { void copyConfigJson(); }} />
          <DefaultButton text={configEditorOpen ? "Close editor" : "Edit config in UI"} onClick={() => setConfigEditorOpen((open) => !open)} />
        </Stack>
        {configEditorOpen && <>
          <TextField ariaLabel="Config JSON" placeholder="Paste or edit the full config JSON" multiline rows={16} value={configJson} onChange={(_event, value) => setConfigJson(value || "")} />
          <p>Edits are validated and saved in this browser for this webpart. Apply reloads the runner; Reset restores the checked-in config.</p>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Apply config and reload" onClick={applyConfigJson} />
            <DefaultButton text="Reset to checked-in config" onClick={resetConfigJson} />
          </Stack>
        </>}
        {configNotice && <MessageBar messageBarType={MessageBarType.info}>{configNotice}</MessageBar>}
      </div>
      <ConfigSetupAssistant context={props.context} />
      <div className="researchApiPanel__session" aria-label="Test session identity">
        <b>Current browser session</b>
        <span>Signed in as: {signedInAccount}</span>
        <span>Target site: {config.environment.siteUrl}</span>
        <span>Config file: {CONFIG_FILE_LABEL}</span>
        <span>Configured primary: {config.principals.users.primary.upn || config.principals.users.primary.loginName || "not set"}{(config.principals.users.primary.upn || config.principals.users.primary.loginName) ? (accountMatch(config.principals.users.primary.upn || config.principals.users.primary.loginName) ? " · matches this session" : " · different account") : ""}</span>
        <span>Configured secondary: {config.principals.users.secondary.upn || config.principals.users.secondary.loginName || "not set"}{(config.principals.users.secondary.upn || config.principals.users.secondary.loginName) ? (accountMatch(config.principals.users.secondary.upn || config.principals.users.secondary.loginName) ? " · matches this session" : " · different account") : ""}</span>
        <span>Requests use the signed-in account above; the configured identities are comparison references.</span>
      </div>
      {(notice || choiceNotice) && <MessageBar messageBarType={record && record.status >= 200 && record.status < 300 ? MessageBarType.success : MessageBarType.warning}>{[notice, choiceNotice].filter(Boolean).join(" ")}</MessageBar>}
      <Stack tokens={{ childrenGap: 12 }}>
        <div className="researchApiPanel__modules" role="group" aria-label="SharePoint API modules">
          <Label>API modules · select a module to test</Label>
          <div className="researchApiPanel__moduleGroups">
            {moduleGroups.map((group) => {
              const modules = group.modules.filter((moduleName) => moduleNames.indexOf(moduleName) >= 0);
              if (modules.length === 0) return null;
              return <div className="researchApiPanel__moduleGroup" key={group.label}>
                <span className="researchApiPanel__moduleGroupLabel">{group.label}</span>
                <div className="researchApiPanel__moduleButtons">
                  {modules.map((moduleName) => {
                    const count = endpointOptions.filter((option) => getEndpoint(String(option.key)).module === moduleName).length;
                    const label = moduleName.replace(/[-_]/g, " ").replace(/\b\w/g, (letter) => letter.toUpperCase());
                    return moduleName === selectedModule
                      ? <PrimaryButton key={moduleName} text={label + " · " + count} aria-pressed={true} onClick={() => selectModule(moduleName)} disabled={busy} />
                      : <DefaultButton key={moduleName} text={label + " · " + count} aria-pressed={false} onClick={() => selectModule(moduleName)} disabled={busy} />;
                  })}
                </div>
              </div>;
            })}
            {moduleNames.filter((moduleName) => !moduleGroups.some((group) => group.modules.indexOf(moduleName) >= 0)).length > 0 && <div className="researchApiPanel__moduleGroup">
              <span className="researchApiPanel__moduleGroupLabel">Other modules</span>
              <div className="researchApiPanel__moduleButtons">{moduleNames.filter((moduleName) => !moduleGroups.some((group) => group.modules.indexOf(moduleName) >= 0)).map((moduleName) => <DefaultButton key={moduleName} text={moduleName} onClick={() => selectModule(moduleName)} disabled={busy} />)}</div>
            </div>}
          </div>
        </div>
        <Dropdown multiSelect label="Endpoint in selected module" options={endpointOptions.filter((option) => getEndpoint(String(option.key)).module === selectedModule)} selectedKeys={endpointKeys} onChange={changeEndpoint} />
        <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
          <DefaultButton text="Download all declared API parameters (JSON)" onClick={downloadParameterInventory} />
          <DefaultButton text="Download this case plan for Burp (.http)" onClick={downloadCasePlan} disabled={busy} />
        </Stack>
        <div className="researchApiPanel__payloads">
          <h3>Bug-family payload sets</h3>
          <div className="researchApiPanel__payloadSelectors">
            <Dropdown multiSelect label="Bug type" placeholder="Select payload families" options={payloadFamilyOptions} selectedKeys={payloadFamilyKeys} onChange={(_e, option) => { const keys = toggleDropdownKey(payloadFamilyKeys, option); setPayloadFamilyKeys(keys); setPayloadFamily((option && option.selected ? String(option.key) : keys[keys.length - 1] || "") as PayloadFamily | ""); setPayloadId(""); setPayloadKeys([]); }} />
            <Dropdown multiSelect label={"Payload case · " + payloadCases.length + " generated"} placeholder={payloadFamily ? "Select cases to apply" : "Select a bug type first"} options={payloadCaseOptions} selectedKeys={payloadKeys} disabled={!payloadFamily} onChange={(_e, option) => { const keys = toggleDropdownKey(payloadKeys, option); setPayloadKeys(keys); setPayloadId(option && option.selected ? String(option.key) : keys[keys.length - 1] || ""); }} />
          </div>
          {selectedPayload && <div className="researchApiPanel__payloadDetails"><b>{selectedPayload.classification}{endpoint.destructive ? " · write request" : " · request mutation"}</b><span>{selectedPayload.rationale}</span><span>Review for: {(selectedPayload.expectedObservations || []).join(" · ")}</span></div>}
        </div>
        <Dropdown multiSelect label="HTTP method" options={methodOptions} selectedKeys={methodKeys} onChange={(_e, option) => { const keys = toggleDropdownKey(methodKeys, option); setMethodKeys(keys); const active = option && option.selected ? String(option.key) : keys[keys.length - 1]; if (active) setMethod(active as HttpMethod); }} />
        <div className="researchApiPanel__fields">
          {selectedModule === "items" && <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
            <TextField ariaLabel="Custom item IDs" placeholder="Add item IDs · comma or space separated" value={customItemIdsInput} onChange={(_e, value) => setCustomItemIdsInput(value || "")} onKeyDown={(event) => { if (event.key === "Enter") { event.preventDefault(); addCustomItemIds(); } }} />
            <DefaultButton text="Add IDs" onClick={addCustomItemIds} disabled={!customItemIdsInput.trim()} />
          </Stack>}
          {selectedModule === "items" && customItemIds.length > 0 && <div className="researchApiPanel__customIds">{customItemIds.map((id) => <span key={id}>{id}<IconButton iconProps={{ iconName: "Cancel" }} title={"Remove item ID " + id} ariaLabel={"Remove item ID " + id} onClick={() => setCustomItemIds((current) => current.filter((value) => value !== id))} /></span>)}</div>}
          {pathAndQuery.map((parameter) => {
            const key = parameter.location + ":" + parameter.name;
            const isOptionalQueryParameter = parameter.location === "query" && !parameter.required;
            const variants = idVariantsFor(parameter.name);
            const combined = variants.map((variant) => ({ key: variant.value, text: variant.label + " · " + variant.value })).concat(discoveredChoices[parameter.name] || []);
            const deduped = combined.filter((option, index) => combined.findIndex((candidate) => candidate.key === option.key) === index);
            const selectable = deduped.length > 0;
            const selected = values[key] || "";
            return <React.Fragment key={key}>
              {isOptionalQueryParameter && <Checkbox label={"Include " + parameter.name} checked={enabledQueryFilters[key] === true} onChange={(_event, checked) => setEnabledQueryFilters((current) => Object.assign({}, current, { [key]: checked === true }))} />}
              {(!isOptionalQueryParameter || enabledQueryFilters[key] === true) && (selectable && deduped.length > 0
                ? <Dropdown ariaLabel={parameter.location + " " + parameter.name} placeholder={parameter.name} options={deduped} selectedKey={deduped.some((option) => option.key === selected) ? selected : undefined} onChange={(_e, option) => { updateValue(key, option ? String(option.key) : ""); }} />
                : <TextField ariaLabel={parameter.location + " " + parameter.name} placeholder={parameter.name} value={selected} onChange={(_e, value) => updateValue(key, value)} required={parameter.required} />)}
            </React.Fragment>;
          })}
        </div>
        {bodyRequired && <TextField ariaLabel="Request body" placeholder={"Request body · " + (endpoint.requestBody?.contentType || "application/json")} multiline rows={8} value={bodyText} onChange={(_e, value) => setBodyText(value || "")} />}
        {(endpoint.requiredHeaders || []).filter((name) => !/^(authorization|cookie|x-requestdigest)$/i.test(name)).map((name) => <TextField key={name} ariaLabel={"Header " + name} placeholder={name} value={headers[name] || ""} onChange={(_e, value) => setHeaders((current) => ({ ...current, [name]: value || "" }))} />)}
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Preview request" onClick={previewRequest} disabled={busy} />
          <DefaultButton text="Copy raw HTTP" onClick={() => { void copyRawRequest(); }} disabled={busy} />
          <PrimaryButton text="Execute API call" onClick={() => { void execute(); }} disabled={busy} />
          <DefaultButton text="Sweep available IDs (read only)" onClick={() => { void executeConfiguredIdSweep(); }} disabled={busy || method !== "GET" || endpoint.destructive || !pathAndQuery.some((p) => idVariantsFor(p.name).length > 0)} />
          {busy && <Spinner size={SpinnerSize.small} label="Executing" />}
        </Stack>
        {sweepResults.length > 0 && <div><Label>ID sweep results · {sweepResults.length} requests</Label><pre className="researchApiPanel__output">{sweepResults.join("\n")}</pre></div>}
        {builtRequest && <div><Label>Request</Label><pre className="researchApiPanel__output">{builtRequest}</pre></div>}
        {record && <div>
          <Stack horizontal verticalAlign="center">
            <Label>Response · HTTP {record.status} · {record.responseClass} · {record.durationMs} ms</Label>
            <IconButton iconProps={{ iconName: "Copy" }} title="Copy complete response" ariaLabel="Copy complete response" onClick={() => { void copyResponse(); }} />
          </Stack>
          <pre className="researchApiPanel__output researchApiPanel__fullResponse">{record.responseBody || "(empty response body)"}</pre>
        </div>}
      </Stack>
    </section>
    </>
  );
}
