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
import { applyPayloadToRequestInput, CaseRequestInput, payloadContextFor } from "../../research/casePlan";
import {
  createSiteScriptActionDefinition,
  createSiteScriptContentTypeSubaction,
  createSiteScriptListSubaction,
  parseSiteScriptActionDefinition,
  validateSiteScriptActionDefinition,
  SITE_SCRIPT_LIST_SUBACTION_VERBS,
  SITE_SCRIPT_VERBS,
  SiteScriptActionDefinition,
  SiteScriptContentTypeSubactionVerb,
  SiteScriptListSubactionVerb,
  SiteScriptVerb,
} from "../../research/siteScriptActions";
import { toRawHttp } from "../../research/http";
import ConfigSetupAssistant from "./ConfigSetupAssistant";
import "./ResearchApiPanel.scss";

interface ResearchApiPanelProps {
  context: WebPartContext;
}

const CONFIG_FILE_LABEL = "src/config/test-config.second.json";

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
const firstModuleName: string = moduleNames.indexOf("items") >= 0 ? "items" : (moduleNames[0] || getEndpoint(String(endpointOptions[0].key)).module);
const firstModuleEndpoint: EndpointDefinition = listEndpoints().find((candidate) => candidate.module === firstModuleName) || getEndpoint(String(endpointOptions[0].key));

function configuredPath(name: string, config: ReturnType<typeof loadTestConfig>, endpoint?: EndpointDefinition): string {
  if (name === "listId") return config.resources.lists.primary.id;
  if (name === "itemId") return String(config.resources.items.primaryItemId || "");
  if (name === "pageId") return "";
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
  if (name === "commentId") return "";
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
interface ResponseHistoryEntry { id: number; label: string; request: string; requestPayload: string; record: ExecutionRecord; }
interface EditableRequest { method: string; url: string; headers: Record<string, string>; body: string; }
interface DynamicSiteScript { id: string; title: string; version?: number; content?: any; }
interface DynamicSiteDesign { id: string; title: string; siteScriptIds: string[]; }

function editableRequestJson(spec: HttpRequestSpec): string {
  return JSON.stringify({ method: spec.method, url: spec.url, headers: spec.headers, body: spec.body || "" }, null, 2);
}

function applyEditableRequest(spec: HttpRequestSpec, text: string, endpoint?: EndpointDefinition): HttpRequestSpec {
  let parsed: unknown;
  try { parsed = JSON.parse(text); }
  catch (_error) { throw new Error("Edited request must be valid JSON."); }
  if (!parsed || typeof parsed !== "object") throw new Error("Edited request must be a JSON object.");
  const value = parsed as Partial<EditableRequest>;
  const method = String(value.method || "").toUpperCase();
  if (["GET", "POST", "PUT", "PATCH", "DELETE", "MERGE", "HEAD", "OPTIONS"].indexOf(method) < 0) throw new Error("Edited request method is not supported.");
  if (typeof value.url !== "string" || !/^https:\/\//i.test(value.url)) throw new Error("Edited request URL must be an absolute HTTPS URL.");
  if (!value.headers || typeof value.headers !== "object" || Array.isArray(value.headers)) throw new Error("Edited request headers must be a JSON object.");
  const headers: Record<string, string> = {};
  Object.keys(value.headers).forEach((name) => {
    const headerValue = (value.headers as Record<string, unknown>)[name];
    if (!name.trim() || /[\r\n]/.test(name) || typeof headerValue !== "string" || /[\r\n]/.test(headerValue)) throw new Error("Each edited header must have a valid name and a single-line string value.");
    headers[name] = headerValue;
  });
  if (value.body !== undefined && typeof value.body !== "string") throw new Error("Edited request body must be a string.");
  const isWrite = ["GET", "HEAD", "OPTIONS"].indexOf(method) < 0;
  const body = value.body || "";
  if ((method === "GET" || method === "HEAD") && body.length > 0) throw new Error(method + " requests cannot include a body in the browser request transport.");
  const requiresDigest = !!endpoint && isWrite && (endpoint.requiresDigest || endpoint.auth === "cookie+digest");
  return Object.assign({}, spec, { method: method as HttpMethod, intendedMethod: method as HttpMethod, url: value.url, headers, body: method === "GET" || method === "HEAD" ? undefined : body, requiresDigest });
}

function requestSnapshot(spec: HttpRequestSpec): string {
  return JSON.stringify({ method: spec.method, intendedMethod: spec.intendedMethod, url: spec.url, headers: spec.headers, body: spec.body || "" }, null, 2);
}

function responseSizeBytes(record: ExecutionRecord): number {
  return new Blob([record.responseBody || ""]).size;
}

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

async function postSharePointJson(context: WebPartContext, url: string, body: unknown): Promise<any> {
  const response = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
    headers: {
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;charset=utf-8",
    },
    body: JSON.stringify(body),
  });
  const text = await response.text();
  let result: any;
  try { result = text ? JSON.parse(text) : null; } catch (_error) { result = text; }
  if (!response.ok) throw new Error("HTTP " + response.status + " loading SharePoint site script data: " + (typeof result === "string" ? result : "request failed") + ".");
  return result;
}

function choice(key: unknown, text: unknown): SharePointChoice {
  return { key: String(key), text: String(text || key) + " · " + String(key) };
}

function sharePointObject(value: any): any {
  let current = value;
  while (current && typeof current === "object" && !Array.isArray(current) && current.d && typeof current.d === "object") current = current.d;
  return current;
}

function dynamicRows(value: any): any[] {
  if (Array.isArray(value)) return value;
  if (!value || typeof value !== "object") return [];
  const preferred = ["value", "results", "SiteScripts", "siteScripts", "SiteDesigns", "siteDesigns"];
  for (let i = 0; i < preferred.length; i++) {
    const child = value[preferred[i]];
    if (Array.isArray(child)) return child;
    if (child && typeof child === "object") {
      const nested = dynamicRows(child);
      if (nested.length) return nested;
    }
  }
  const keys = Object.keys(value);
  for (let i = 0; i < keys.length; i++) {
    const child = value[keys[i]];
    if (Array.isArray(child)) {
      const matching = child.filter((row) => row && typeof row === "object" && (row.Id || row.ID || row.id));
      if (matching.length) return matching;
    }
  }
  return [];
}

function dynamicEntity(value: any): any {
  const root = sharePointObject(value);
  if (!root || typeof root !== "object") return {};
  if (root.Id || root.ID || root.id || root.Content || root.content || root.SiteScriptIds || root.SiteScripts) return root;
  const keys = Object.keys(root);
  for (let i = 0; i < keys.length; i++) {
    const child = root[keys[i]];
    if (child && typeof child === "object" && !Array.isArray(child) && (child.Id || child.ID || child.id || child.Content || child.content)) return child;
  }
  return root;
}

function dynamicContent(value: any): any {
  const entity = dynamicEntity(value);
  let content = entity.Content !== undefined ? entity.Content : entity.content;
  if (typeof content === "string") {
    try { content = JSON.parse(content); } catch (_error) { return undefined; }
  }
  if (!content || typeof content !== "object") content = entity;
  return content && typeof content === "object" ? content : undefined;
}

function dynamicId(value: any): string {
  if (typeof value === "string") return /^[{]?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}[}]?$/i.test(value) ? value.replace(/[{}]/g, "") : "";
  if (typeof value === "number") return String(value);
  const entity = dynamicEntity(value);
  const direct = entity.Id || entity.ID || entity.id || entity.TaskId || entity.taskId || entity.RunId || entity.runId || entity.SiteDesignRunId || entity.siteDesignRunId;
  if (direct !== undefined && direct !== null && String(direct)) return String(direct);
  const keys = Object.keys(entity || {});
  for (let i = 0; i < keys.length; i++) {
    const child = entity[keys[i]];
    if (child && typeof child === "object") {
      const nested = dynamicId(child);
      if (nested) return nested;
    }
  }
  return "";
}

function normalizeSiteScripts(value: any): DynamicSiteScript[] {
  return dynamicRows(sharePointObject(value)).map((row) => ({
    id: String(row.Id || row.ID || row.id || ""),
    title: String(row.Title || row.title || row.Name || row.name || row.Id || row.ID || row.id || "Site script"),
    version: typeof row.Version === "number" ? row.Version : undefined,
    content: dynamicContent(row),
  })).filter((row) => !!row.id);
}

function normalizeSiteDesigns(value: any): DynamicSiteDesign[] {
  return dynamicRows(sharePointObject(value)).map((row) => ({
    id: String(row.Id || row.ID || row.id || ""),
    title: String(row.Title || row.title || row.Name || row.name || row.Id || row.ID || row.id || "Site design"),
    siteScriptIds: Array.isArray(row.SiteScriptIds) ? row.SiteScriptIds.map(String) : (Array.isArray(row.siteScriptIds) ? row.siteScriptIds.map(String) : []),
  })).filter((row) => !!row.id);
}

function siteScriptActionsFrom(value: any): unknown[] | undefined {
  const content = dynamicContent(value);
  return content && Array.isArray(content.actions) ? content.actions : undefined;
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
    Object.keys(resources.lists || {}).forEach((key) => {
      const list = resources.lists[key];
      const label = list && list.title ? String(list.title) : key;
      add(label, list && list.id);
    });
  } else if (name === "itemId" || name === "ID") {
    const items = resources.items || {};
    add("Primary item", items.primaryItemId);
    add("Secondary item", items.secondaryItemId);
    add("Non-existent item", items.nonExistentItemId);
    (items.boundaryItemIds || []).forEach((value: unknown, index: number) => add("Boundary item " + (index + 1), value));
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
  const requiresTenantConfig = ["listId", "siteId", "viewId", "View", "formId", "contentTypeId", "fieldId", "groupId", "setId", "termStoreId", "termGroupId", "termSetId", "termId", "commentId", "pageId", "recycleBinItemId"].indexOf(parameter.name) >= 0;
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

function formatPreview(spec: HttpRequestSpec): string {
  return editableRequestJson(spec);
}

export default function ResearchApiPanel(props: ResearchApiPanelProps): React.ReactElement {
  const config = React.useMemo(() => loadTestConfig(), []);
  const [endpointId, setEndpointId] = React.useState<string>(firstModuleEndpoint.id);
  const [endpointKeys, setEndpointKeys] = React.useState<string[]>([firstModuleEndpoint.id]);
  const [selectedModule, setSelectedModule] = React.useState<string>(firstModuleName);
  const [payloadFamily, setPayloadFamily] = React.useState<PayloadFamily | "">("");
  const [payloadFamilyKeys, setPayloadFamilyKeys] = React.useState<PayloadFamily[]>([]);
  const [payloadId, setPayloadId] = React.useState<string>("");
  const [payloadKeys, setPayloadKeys] = React.useState<string[]>([]);
  const endpoint = React.useMemo(() => getEndpoint(endpointId), [endpointId]);
  const [values, setValues] = React.useState<Record<string, string>>(() => initialValues(firstModuleEndpoint, config));
  const [method, setMethod] = React.useState<HttpMethod>(endpoint.methods[0] || "GET");
  const [methodKeys, setMethodKeys] = React.useState<string[]>([endpoint.methods[0] || "GET"]);
  const [discoveredChoices, setDiscoveredChoices] = React.useState<Record<string, SharePointChoice[]>>({});
  const [choiceNotice, setChoiceNotice] = React.useState<string>("");
  const [siteScripts, setSiteScripts] = React.useState<DynamicSiteScript[]>([]);
  const [siteDesigns, setSiteDesigns] = React.useState<DynamicSiteDesign[]>([]);
  const [selectedSiteScriptId, setSelectedSiteScriptId] = React.useState<string>("");
  const [selectedSiteScriptIds, setSelectedSiteScriptIds] = React.useState<string[]>([]);
  const [selectedSiteDesignId, setSelectedSiteDesignId] = React.useState<string>("");
  const [selectedSiteDesignTaskId, setSelectedSiteDesignTaskId] = React.useState<string>("");
  const [siteScriptActionsText, setSiteScriptActionsText] = React.useState<string>("[]");
  const [selectedSiteScriptActionIndex, setSelectedSiteScriptActionIndex] = React.useState<number>(-1);
  const [siteScriptActionToAdd, setSiteScriptActionToAdd] = React.useState<SiteScriptVerb>("createSPList");
  const [siteScriptCatalogLoading, setSiteScriptCatalogLoading] = React.useState<boolean>(false);
  const [siteScriptCatalogError, setSiteScriptCatalogError] = React.useState<string>("");
  const [enabledQueryFilters, setEnabledQueryFilters] = React.useState<Record<string, boolean>>({});
  const [customItemIds, setCustomItemIds] = React.useState<string[]>([]);
  const [customItemIdsInput, setCustomItemIdsInput] = React.useState<string>("");
  const [bodyText, setBodyText] = React.useState<string>(initialBody(endpoint));
  const [siteScriptVerb, setSiteScriptVerb] = React.useState<SiteScriptVerb>("createSPList");
  const [siteScriptActionText, setSiteScriptActionText] = React.useState<string>(() => JSON.stringify(createSiteScriptActionDefinition("createSPList"), null, 2));
  const [siteScriptActionError, setSiteScriptActionError] = React.useState<string>("");
  const [siteScriptSubactionVerb, setSiteScriptSubactionVerb] = React.useState<string>("setTitle");
  const [headers, setHeaders] = React.useState<Record<string, string>>({});
  const [busy, setBusy] = React.useState<boolean>(false);
  const [record, setRecord] = React.useState<ExecutionRecord | null>(null);
  const [responseHistory, setResponseHistory] = React.useState<ResponseHistoryEntry[]>([]);
  const [responseHistoryOpen, setResponseHistoryOpen] = React.useState<boolean>(false);
  const [expandedRequestIds, setExpandedRequestIds] = React.useState<Record<number, boolean>>({});
  const [responseStatusFilter, setResponseStatusFilter] = React.useState<string>("");
  const [responseSizeFilter, setResponseSizeFilter] = React.useState<string>("");
  const [notice, setNotice] = React.useState<string>("");
  const [builtRequest, setBuiltRequest] = React.useState<string>("");
  const [sweepResults, setSweepResults] = React.useState<string[]>([]);
  const [batchSummary, setBatchSummary] = React.useState<{ attempted: number; succeeded: number; failed: number; notSent: number } | null>(null);
  const lastExecutionAt = React.useRef<number>(0);
  const [configEditorOpen, setConfigEditorOpen] = React.useState<boolean>(false);
  const [configJson, setConfigJson] = React.useState<string>(() => JSON.stringify(config, null, 2));
  const [configNotice, setConfigNotice] = React.useState<string>("");
  const idVariantsFor = (name: string): ConfiguredIdVariant[] => {
    const configured = configuredIdVariants(name, config);
    if (name === "pageId") {
      const discovered = (discoveredChoices.pageId || []).map((option) => ({ label: option.text.replace(/ · .*$/, ""), value: option.key }));
      return configured.concat(discovered);
    }
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
    const families = payloadFamilyKeys.length > 0 ? payloadFamilyKeys : (payloadFamily ? [payloadFamily] : []);
    return families.reduce((all, family) => {
      let cases = generatePayloads(family, payloadContextFor(endpoint, family, bodyText))
        .concat(generateEndpointCases(endpoint, family, config));
      if (family === "business-logic" && /\.shareLink$/.test(endpoint.id)) {
        cases = cases.concat(generateShareLinkMatrix(config.principals.shareLinkKinds, config.principals.shareLinkRoleKinds));
      }
      return all.concat(cases);
    }, [] as Payload[]);
  }, [endpointId, payloadFamily, payloadFamilyKeys, bodyText, config]);
  const selectedPayload = payloadCases.find((candidate) => candidate.id === payloadId);
  const selectedPayloads = React.useMemo(
    () => payloadCases.filter((candidate) => payloadKeys.indexOf(candidate.id) >= 0),
    [payloadCases, payloadKeys]
  );
  const payloadCaseOptions: IDropdownOption[] = payloadCases.map((candidate) => ({ key: candidate.id, text: (payloadFamilyKeys.length > 1 ? PAYLOAD_FAMILIES[candidate.family].title + " · " : "") + (candidate.control ? "Control · " : "") + candidate.name }));
  const siteScriptVerbOptions: IDropdownOption[] = SITE_SCRIPT_VERBS.map((verb) => ({ key: verb, text: verb }));
  const siteScriptActionVerbOptions: IDropdownOption[] = SITE_SCRIPT_VERBS.map((verb) => ({ key: verb, text: verb }));
  let loadedSiteScriptActions: unknown[] = [];
  try { const parsedActions: unknown = JSON.parse(siteScriptActionsText || "[]"); if (Array.isArray(parsedActions)) loadedSiteScriptActions = parsedActions; } catch (_error) { loadedSiteScriptActions = []; }
  const loadedSiteScriptActionOptions: IDropdownOption[] = loadedSiteScriptActions.map((action: any, index) => ({ key: index, text: String(action && action.verb || "Action") + " · " + (index + 1) }));
  const dynamicSiteScriptOptions: IDropdownOption[] = siteScripts.map((script) => ({ key: script.id, text: script.title + " · " + script.id }));
  const dynamicSiteDesignOptions: IDropdownOption[] = siteDesigns.map((design) => ({ key: design.id, text: design.title + " · " + design.id }));
  const isSiteScriptCollectionEditor = endpoint.id === "siteScripts.create" || endpoint.id === "siteScripts.update";
  const isDynamicSiteScriptsModule = selectedModule === "site-scripts" || selectedModule === "site-designs";
  const isSiteDesignCreateOrUpdate = endpoint.id === "siteDesigns.create" || endpoint.id === "siteDesigns.update";
  const siteScriptSubactionOptions: IDropdownOption[] = (siteScriptVerb === "createContentType"
    ? (["addSiteColumn", "removeSiteColumn"] as SiteScriptContentTypeSubactionVerb[])
    : SITE_SCRIPT_LIST_SUBACTION_VERBS).map((verb) => ({ key: verb, text: verb }));

  const selectSiteScriptVerb = (verb: SiteScriptVerb): void => {
    const text = JSON.stringify(createSiteScriptActionDefinition(verb), null, 2);
    setSiteScriptVerb(verb);
    setSiteScriptSubactionVerb(verb === "createContentType" ? "addSiteColumn" : "setTitle");
    setSiteScriptActionText(text);
    setSiteScriptActionError("");
  };
  const addSiteScriptSubaction = (): void => {
    try {
      const definition = parseSiteScriptActionDefinition(siteScriptActionText) as SiteScriptActionDefinition & { subactions?: unknown[] };
      const subaction = definition.verb === "createContentType"
        ? createSiteScriptContentTypeSubaction(siteScriptSubactionVerb as SiteScriptContentTypeSubactionVerb)
        : createSiteScriptListSubaction(siteScriptSubactionVerb as SiteScriptListSubactionVerb);
      definition.subactions = (definition.subactions || []).concat([subaction]);
      setSiteScriptActionText(JSON.stringify(definition, null, 2));
      setSiteScriptActionError("");
    } catch (error) {
      setSiteScriptActionError(error instanceof Error ? error.message : String(error));
    }
  };

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
    if (!endpoint.parameters.some((parameter) => parameter.name === "pageId" && parameter.location === "path")) return () => { cancelled = true; };
    const baseUrl = config.environment.siteUrl.replace(/\/+$/, "");
    if (!baseUrl) return () => { cancelled = true; };
    setDiscoveredChoices((current) => Object.assign({}, current, { pageId: [] }));
    const load = async (): Promise<void> => {
      try {
        const lists = await getAllSharePointRows<any>(props.context, baseUrl + "/_api/web/lists?$select=Id,Title,BaseTemplate&$top=500");
        const pagesList = lists.filter((list) => Number(list.BaseTemplate) === 119 || String(list.Title || "").toLowerCase() === "site pages")[0];
        if (!pagesList) throw new Error("The Site Pages library was not found on this site.");
        const listId = String(pagesList.Id).replace(/[{}]/g, "");
        const pages = await getAllSharePointRows<any>(props.context, baseUrl + "/_api/web/lists(guid'" + encodeURIComponent(listId) + "')/items?$select=Id,Title,FileRef,FileLeafRef,UniqueId&$top=500");
        if (cancelled) return;
        const choices = pages.map((page) => choice(page.Id, page.FileLeafRef || page.Title || page.FileRef || "Page #" + page.Id));
        setDiscoveredChoices((current) => Object.assign({}, current, { pageId: choices }));
        const testPage = String(((config.resources as any).pages || {}).testPage || "").replace(/\\/g, "/").replace(/^\/+/, "").toLowerCase();
        const testLeaf = testPage.split("/").pop() || "";
        const preferred = pages.filter((page) => {
          const fileRef = String(page.FileRef || "").replace(/\\/g, "/").toLowerCase();
          const leaf = String(page.FileLeafRef || "").toLowerCase();
          return !!testPage && (fileRef.endsWith("/" + testPage) || leaf === testLeaf);
        })[0];
        const currentId = values["path:pageId"];
        const currentExists = currentId && pages.some((page) => String(page.Id) === currentId);
        const defaultPage = currentExists ? undefined : preferred;
        if (defaultPage) setValues((current) => Object.assign({}, current, { "path:pageId": String(defaultPage.Id) }));
      } catch (error) {
        if (!cancelled) setChoiceNotice("Could not load page IDs from the Site Pages library: " + (error instanceof Error ? error.message : String(error)));
      }
    };
    void load();
    return () => { cancelled = true; };
  }, [props.context, config.environment.siteUrl, endpointId]);

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

  React.useEffect(() => {
    let cancelled = false;
    const listId = values["path:listId"] || values["path:listID"];
    const itemId = values["path:itemId"] || values["path:ID"];
    const baseUrl = config.environment.siteUrl.replace(/\/+$/, "");
    setDiscoveredChoices((current) => Object.assign({}, current, { commentId: [] }));
    if (endpoint.module !== "comments" || !listId || !itemId || !baseUrl) {
      setValues((current) => current["path:commentId"] ? Object.assign({}, current, { "path:commentId": "" }) : current);
      return () => { cancelled = true; };
    }
    const commentsUrl = baseUrl + "/_api/web/lists(guid'" + encodeURIComponent(listId.replace(/[{}]/g, "")) + "')/items(" + encodeURIComponent(itemId) + ")/comments";
    const load = async (): Promise<void> => {
      try {
        const comments = await getAllSharePointRows<any>(props.context, commentsUrl);
        if (cancelled) return;
        const choices = comments
          .filter((comment) => comment && (comment.id !== undefined || comment.Id !== undefined))
          .map((comment) => choice(comment.id !== undefined ? comment.id : comment.Id, comment.text || comment.Text || "Comment"));
        setDiscoveredChoices((current) => Object.assign({}, current, { commentId: choices }));
        setValues((current) => {
          const currentId = current["path:commentId"] || "";
          const selected = choices.some((option) => option.key === currentId) ? currentId : (choices[0] ? choices[0].key : "");
          return currentId === selected ? current : Object.assign({}, current, { "path:commentId": selected });
        });
        setChoiceNotice("");
      } catch (error) {
        if (!cancelled) setChoiceNotice("Could not load comment IDs for the selected item: " + (error instanceof Error ? error.message : String(error)));
      }
    };
    void load();
    return () => { cancelled = true; };
  }, [props.context, config.environment.siteUrl, endpoint.module, values["path:listId"], values["path:listID"], values["path:itemId"], values["path:ID"]]);

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
    setBatchSummary(null);
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
    setBatchSummary(null);
    setBuiltRequest("");
    setNotice("");
  };

  const selectAllPayloadCases = (): void => {
    const keys = payloadCases.map((candidate) => candidate.id);
    setPayloadKeys(keys);
    setPayloadId(keys[keys.length - 1] || "");
  };
  const clearPayloadCases = (): void => {
    setPayloadKeys([]);
    setPayloadId("");
  };
  const selectAllPayloadFamilies = (): void => {
    const families = endpoint.payloadFamilies;
    setPayloadFamilyKeys(families);
    setPayloadFamily(families[0] || "");
    setPayloadKeys([]);
    setPayloadId("");
  };

  const refreshSiteScriptCatalog = React.useCallback(async (): Promise<void> => {
    const base = config.environment.siteUrl.replace(/\/+$/, "");
    if (!base) return;
    setSiteScriptCatalogLoading(true);
    setSiteScriptCatalogError("");
    try {
      const [scriptResult, designResult] = await Promise.all([
        postSharePointJson(props.context, base + "/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts", {}),
        postSharePointJson(props.context, base + "/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns", {}),
      ]);
      const nextScripts = normalizeSiteScripts(scriptResult);
      const nextDesigns = normalizeSiteDesigns(designResult);
      setSiteScripts(nextScripts);
      setSiteDesigns(nextDesigns);
      setSelectedSiteScriptId((current) => nextScripts.some((script) => script.id === current) ? current : "");
      setSelectedSiteScriptIds((current) => current.filter((id) => nextScripts.some((script) => script.id === id)));
      setSelectedSiteDesignId((current) => nextDesigns.some((design) => design.id === current) ? current : "");
    } catch (error) {
      setSiteScriptCatalogError(error instanceof Error ? error.message : String(error));
    } finally {
      setSiteScriptCatalogLoading(false);
    }
  }, [props.context, config.environment.siteUrl]);

  React.useEffect(() => {
    if (selectedModule === "site-scripts" || selectedModule === "site-designs") void refreshSiteScriptCatalog();
  }, [selectedModule, refreshSiteScriptCatalog]);

  const loadSiteScriptDetails = async (id: string): Promise<void> => {
    if (!id) { setSiteScriptActionsText("[]"); return; }
    setSiteScriptCatalogLoading(true);
    setSiteScriptCatalogError("");
    try {
      const base = config.environment.siteUrl.replace(/\/+$/, "");
      const result = await postSharePointJson(props.context, base + "/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata", { id });
      const entity = dynamicEntity(result);
      const content = dynamicContent(result);
      const script = normalizeSiteScripts([Object.assign({}, entity, { Content: content })])[0];
      if (script) setSiteScripts((current) => current.some((item) => item.id === id)
        ? current.map((item) => item.id === id ? Object.assign({}, item, script, { content }) : item)
        : current.concat([script]));
      const actions = siteScriptActionsFrom({ Content: content });
      if (actions) {
        setSiteScriptActionsText(JSON.stringify(actions, null, 2));
        setSelectedSiteScriptActionIndex(actions.length ? 0 : -1);
        if (actions.length) setSiteScriptActionText(JSON.stringify(actions[0], null, 2));
      }
    } catch (error) {
      setSiteScriptCatalogError(error instanceof Error ? error.message : String(error));
    } finally {
      setSiteScriptCatalogLoading(false);
    }
  };

  const selectSiteScript = (id: string): void => {
    setSelectedSiteScriptId(id);
    void loadSiteScriptDetails(id);
  };

  const selectSiteDesign = async (id: string): Promise<void> => {
    setSelectedSiteDesignId(id);
    if (!id) return;
    const existing = siteDesigns.find((design) => design.id === id);
    if (existing && existing.siteScriptIds.length) setSelectedSiteScriptIds(existing.siteScriptIds);
    try {
      const base = config.environment.siteUrl.replace(/\/+$/, "");
      const result = await postSharePointJson(props.context, base + "/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata", { id });
      const raw = dynamicEntity(result);
      const design = normalizeSiteDesigns([raw])[0];
      if (design) {
        setSiteDesigns((current) => current.map((item) => item.id === id ? design : item));
        if (design.siteScriptIds.length) setSelectedSiteScriptIds(design.siteScriptIds);
      }
      setSiteScriptCatalogError("");
    } catch (error) {
      setSiteScriptCatalogError(error instanceof Error ? error.message : String(error));
    }
  };

  const addSiteScriptAction = (): void => {
    try {
      const parsed: unknown = JSON.parse(siteScriptActionsText || "[]");
      if (!Array.isArray(parsed)) throw new Error("Site script actions must be a JSON array.");
      const next = parsed.concat([createSiteScriptActionDefinition(siteScriptActionToAdd)]);
      setSiteScriptActionsText(JSON.stringify(next, null, 2));
      setSiteScriptCatalogError("");
    } catch (error) {
      setSiteScriptCatalogError(error instanceof Error ? error.message : String(error));
    }
  };

  const updateValue = (key: string, value?: string): void => setValues((current) => ({ ...current, [key]: value || "" }));

  const buildWith = (payloads: Payload[], overrides: Record<string, string> = {}): HttpRequestSpec => {
    const path: Record<string, string | number> = {};
    const query: Record<string, QueryValue> = {};
    const requestHeaders: Record<string, string> = {};
    const isOmitted = (location: string, name: string): boolean => payloads.some((payload) =>
      (payload.action === "omit" && payload.location === location && payload.parameter === name) ||
      (payload.mutations || []).some((mutation) =>
        mutation.action === "omit" && mutation.location === location && mutation.parameter === name));
    endpoint.parameters.forEach((parameter) => {
      if (parameter.location !== "path" && parameter.location !== "query") return;
      const key = parameter.location + ":" + parameter.name;
      if (parameter.location === "query" && !parameter.required && !enabledQueryFilters[key]) return;
      const raw = overrides[key] !== undefined ? overrides[key] : (values[key] || "");
      if (parameter.required && !raw.trim() && !isOmitted(parameter.location, parameter.name)) throw new Error("Enter the required " + parameter.location + " value: " + parameter.name + ". Use the Test config setup check above to discover tenant values.");
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
    if (endpoint.id === "siteScripts.executeAction") {
      const actionDefinition = parseSiteScriptActionDefinition(siteScriptActionText);
      body = { actionDefinition: JSON.stringify(actionDefinition) };
    } else if (endpoint.id === "siteScripts.create" || endpoint.id === "siteScripts.update") {
      let actions: unknown;
      try { actions = JSON.parse(siteScriptActionsText || "[]"); }
      catch (_error) { throw new Error("Site script actions must be valid JSON."); }
      if (!Array.isArray(actions)) throw new Error("Site script actions must be a JSON array.");
      actions.forEach((action) => validateSiteScriptActionDefinition(action));
      body = endpoint.id === "siteScripts.create"
        ? { actions }
        : { updateInfo: { Id: selectedSiteScriptId }, content: { actions } };
      if (endpoint.id === "siteScripts.update" && !selectedSiteScriptId) throw new Error("Select a site script loaded from SharePoint before updating it.");
    } else if (bodyText.trim()) {
      const contentType = endpoint.requestBody?.contentType || "application/json";
      if (contentType.indexOf("application/json") === 0) {
        try { body = JSON.parse(bodyText); }
        catch { throw new Error("Request body must be valid JSON."); }
      } else {
        body = bodyText;
      }
    }
    if (body && typeof body === "object" && !Array.isArray(body)) {
      const dynamicBody = Object.assign({}, body as Record<string, unknown>);
      if (endpoint.id === "siteScripts.metadata" || endpoint.id === "siteScripts.delete") {
        if (!selectedSiteScriptId) throw new Error("Select a site script loaded from SharePoint first.");
        dynamicBody.id = selectedSiteScriptId;
        body = dynamicBody;
      } else if (endpoint.id === "siteDesigns.getMetadata" || endpoint.id === "siteDesigns.delete" || endpoint.id === "siteDesigns.rights" || endpoint.id === "siteDesigns.grantRights" || endpoint.id === "siteDesigns.revokeRights") {
        if (!selectedSiteDesignId) throw new Error("Select a site design loaded from SharePoint first.");
        dynamicBody.id = selectedSiteDesignId;
        body = dynamicBody;
      } else if (endpoint.id === "siteDesigns.update") {
        if (!selectedSiteDesignId) throw new Error("Select a site design loaded from SharePoint first.");
        const updateInfo = Object.assign({}, (dynamicBody.updateInfo || {}) as Record<string, unknown>, { Id: selectedSiteDesignId });
        if (selectedSiteScriptIds.length) updateInfo.SiteScriptIds = selectedSiteScriptIds;
        dynamicBody.updateInfo = updateInfo;
        body = dynamicBody;
      } else if (endpoint.id === "siteDesigns.apply" || endpoint.id === "siteDesigns.addTask" || endpoint.id === "siteDesigns.addTaskCurrentWeb") {
        if (!selectedSiteDesignId) throw new Error("Select a site design loaded from SharePoint first.");
        dynamicBody.siteDesignId = selectedSiteDesignId;
        if (endpoint.id !== "siteDesigns.addTaskCurrentWeb") dynamicBody.webUrl = config.environment.siteUrl;
        body = dynamicBody;
      } else if (endpoint.id === "siteDesigns.getTask") {
        if (!selectedSiteDesignTaskId) throw new Error("Run a site design task first or enter its returned task ID.");
        dynamicBody.taskId = selectedSiteDesignTaskId;
        body = dynamicBody;
      } else if (endpoint.id === "siteDesigns.run") {
        if (!selectedSiteDesignId) throw new Error("Select a site design loaded from SharePoint first.");
        dynamicBody.siteDesignId = selectedSiteDesignId;
        dynamicBody.webUrl = config.environment.siteUrl;
        body = dynamicBody;
      } else if (endpoint.id === "siteDesigns.runStatus") {
        if (!selectedSiteDesignTaskId) throw new Error("Run a site design task first or enter its returned run ID.");
        dynamicBody.runId = selectedSiteDesignTaskId;
        dynamicBody.webUrl = config.environment.siteUrl;
        body = dynamicBody;
      } else if (endpoint.id === "siteDesigns.create") {
        const info = Object.assign({}, (dynamicBody.info || {}) as Record<string, unknown>);
        if (selectedSiteScriptIds.length) info.SiteScriptIds = selectedSiteScriptIds;
        dynamicBody.info = info;
        body = dynamicBody;
      }
    }
    const requestInput: CaseRequestInput = { path, query, headers: requestHeaders, body, method };
    payloads.forEach((payload) => applyPayloadToRequestInput(requestInput, payload, endpoint));
    // A method-override payload rewrites the effective verb to POST and carries
    // the intended verb in X-HTTP-Method, exactly as the exported plan does.
    const effectiveMethod = requestInput.method || method;
    const siteUrlOverride = payloads.map((payload) => payload.siteUrlOverride).filter(Boolean)[0];
    const spec = buildRequest(endpoint.id, {
      siteUrl: siteUrlOverride || config.environment.siteUrl,
      path: requestInput.path,
      query: requestInput.query,
      method: effectiveMethod,
      body: requestInput.body,
      headers: Object.assign(parsedHeaders, requestInput.headers),
      digestPlaceholder: true,
    });
    if (payloads.some((payload) => payload.destructive)) spec.destructive = true;
    if (["PUT", "PATCH", "DELETE", "MERGE"].indexOf(effectiveMethod) >= 0 ||
        (endpoint.methods.indexOf(effectiveMethod) < 0 && effectiveMethod === "POST")) spec.destructive = true;
    return spec;
  };
  const build = (overrides: Record<string, string> = {}): HttpRequestSpec => buildWith(selectedPayload ? [selectedPayload] : [], overrides);
  const requestFromPreview = (): HttpRequestSpec => builtRequest.trim()
    ? applyEditableRequest(build(), builtRequest, endpoint)
    : build();

  const previewRequest = (): void => {
    try {
      setBuiltRequest(formatPreview(build()));
      setNotice("");
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error));
    }
  };

  // Keep the preview in sync with endpoint, method, payload and value changes.
  // An incomplete form keeps the last valid preview; the explicit button
  // reports the validation error instead of flashing it on every keystroke.
  React.useEffect(() => {
    try {
      setBuiltRequest(formatPreview(build()));
    } catch (_error) {
      /* reported by previewRequest / execute */
    }
  }, [endpointId, method, payloadId, payloadKeys, values, bodyText, headers, enabledQueryFilters, payloadFamily, siteScriptActionText]);

  const createExecutor = (): RequestExecutor => {
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
      digest = makeDigestProvider(config.environment.siteUrl, config.safety.requestTimeoutMs);
    }
    return new RequestExecutor(transport, digest, {
      maxTransientRetries: 0,
      maxDigestRetries: 1,
      previewLength: config.safety.maxPayloadBytes > 0 ? Math.min(config.safety.maxPayloadBytes, 12000) : 12000,
    });
  };

  const waitRandomInterval = async (): Promise<void> => {
    if (lastExecutionAt.current === 0) return;
    const delayMs = 25 + Math.floor(Math.random() * 226);
    await new Promise<void>((resolve) => setTimeout(resolve, delayMs));
  };

  const rememberResponse = (label: string, spec: HttpRequestSpec, response: ExecutionRecord): void => {
    setResponseHistory((current) => current.concat([{
      id: Date.now() + current.length,
      label: label,
      request: spec.method + " " + spec.url,
      requestPayload: requestSnapshot(spec),
      record: response,
    }]));
  };

  const captureSiteScriptResponse = (endpointName: string, response: ExecutionRecord): void => {
    if (response.status < 200 || response.status >= 300 || !response.responseBody) return;
    let data: any;
    try { data = JSON.parse(response.responseBody); } catch (_error) { return; }
    if (endpointName === "siteScripts.get") {
      const rows = normalizeSiteScripts(data);
      if (rows.length) setSiteScripts(rows);
      return;
    }
    if (endpointName === "siteScripts.delete") {
      setSiteScripts((current) => current.filter((script) => script.id !== selectedSiteScriptId));
      setSelectedSiteScriptId("");
      setSiteScriptActionsText("[]");
      return;
    }
    if (endpointName === "siteScripts.fromList" || endpointName === "siteScripts.fromWeb") {
      const actions = siteScriptActionsFrom(data);
      if (actions) {
        setSiteScriptActionsText(JSON.stringify(actions, null, 2));
        setSelectedSiteScriptActionIndex(actions.length ? 0 : -1);
        if (actions.length) setSiteScriptActionText(JSON.stringify(actions[0], null, 2));
      }
      return;
    }
    if (endpointName === "siteScripts.create" || endpointName === "siteScripts.update" || endpointName === "siteScripts.metadata") {
      const entity = dynamicEntity(data);
      const content = dynamicContent(data);
      const id = dynamicId(data) || (endpointName === "siteScripts.update" || endpointName === "siteScripts.metadata" ? selectedSiteScriptId : "");
      const actions = siteScriptActionsFrom({ Content: content }) || (endpointName === "siteScripts.create" || endpointName === "siteScripts.update" ? (() => {
        try { const parsed = JSON.parse(siteScriptActionsText || "[]"); return Array.isArray(parsed) ? parsed : undefined; } catch (_error) { return undefined; }
      })() : undefined);
      if (actions) {
        setSiteScriptActionsText(JSON.stringify(actions, null, 2));
        setSelectedSiteScriptActionIndex(actions.length ? 0 : -1);
      }
      if (id) {
        const script: DynamicSiteScript = {
          id,
          title: String(entity.Title || entity.title || (siteScripts.find((item) => item.id === id) || {}).title || "Site script"),
          version: typeof entity.Version === "number" ? entity.Version : undefined,
          content: content || (actions ? { actions } : undefined),
        };
        setSiteScripts((current) => current.some((item) => item.id === id)
          ? current.map((item) => item.id === id ? Object.assign({}, item, script) : item)
          : current.concat([script]));
        setSelectedSiteScriptId(id);
      }
      return;
    }
    if (endpointName === "siteDesigns.get") {
      const rows = normalizeSiteDesigns(data);
      if (rows.length) setSiteDesigns(rows);
      return;
    }
    if (endpointName === "siteDesigns.delete") {
      setSiteDesigns((current) => current.filter((design) => design.id !== selectedSiteDesignId));
      setSelectedSiteDesignId("");
      return;
    }
    if (endpointName === "siteDesigns.create" || endpointName === "siteDesigns.getMetadata" || endpointName === "siteDesigns.update") {
      const entity = dynamicEntity(data);
      const id = dynamicId(data) || (endpointName === "siteDesigns.getMetadata" || endpointName === "siteDesigns.update" ? selectedSiteDesignId : "");
      const design = normalizeSiteDesigns([entity])[0];
      if (id) {
        const next = design || { id, title: String(entity.Title || "Site design"), siteScriptIds: selectedSiteScriptIds };
        setSiteDesigns((current) => current.some((item) => item.id === id)
          ? current.map((item) => item.id === id ? next : item)
          : current.concat([next]));
        setSelectedSiteDesignId(id);
        if (next.siteScriptIds.length) setSelectedSiteScriptIds(next.siteScriptIds);
      }
      return;
    }
    if (endpointName === "siteDesigns.addTask" || endpointName === "siteDesigns.addTaskCurrentWeb" || endpointName === "siteDesigns.apply" || endpointName === "siteDesigns.run" || endpointName === "siteDesigns.runStatus") {
      const returnedId = dynamicId(data);
      if (returnedId) setSelectedSiteDesignTaskId(returnedId);
    }
  };

  const runOne = async (spec: HttpRequestSpec, label: string, executor: RequestExecutor): Promise<ExecutionRecord> => {
    await waitRandomInterval();
    setBuiltRequest(spec.method + " " + spec.url);
    let result: ExecutionRecord;
    try {
      result = await executor.execute(spec);
    } catch (error) {
      result = {
        endpointId: spec.endpointId,
        method: spec.intendedMethod,
        url: spec.url,
        status: 0,
        responseClass: "network",
        durationMs: 0,
        attempts: [],
        responsePreview: "",
        responseBody: "",
        responseHeaders: {},
        startedAt: new Date().toISOString(),
        error: error instanceof Error ? error.message : String(error),
      };
    }
    lastExecutionAt.current = Date.now();
    setRecord(result);
    rememberResponse(label, spec, result);
    captureSiteScriptResponse(endpoint.id, result);
    setNotice(label + ": " + (result.error || (result.status >= 200 && result.status < 300 ? "completed" : "HTTP " + result.status)) + ".");
    return result;
  };

  const execute = async (): Promise<void> => {
    setNotice("");
    setRecord(null);
    setBusy(true);
    try {
      const spec = selectedPayloads.length > 1 ? build() : requestFromPreview();
      await runOne(spec, "Request", createExecutor());
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  };

  /**
   * Runs every selected case. Cases that write to different request cells share
   * one request; conflicting cases, controls, destructive cases and method
   * overrides are sent one at a time, in order.
   */
  const executeSelectedPayloads = async (): Promise<void> => {
    if (selectedPayloads.length === 0) {
      setNotice("Select at least one payload case first.");
      return;
    }
    const groups = selectedPayloads.map((payload) => [payload]);
    setNotice("");
    setRecord(null);
    setBatchSummary(null);
    setBusy(true);
    try {
      let attempted = 0;
      let succeeded = 0;
      let failed = 0;
      let notSent = 0;
      const notSentReasons: string[] = [];
      for (let i = 0; i < groups.length; i++) {
        const group = groups[i];
        const label = "Request " + (i + 1) + "/" + groups.length + " · " + PAYLOAD_FAMILIES[group[0].family].title + " / " + group[0].name;
        setNotice("Sending " + label + ". Remaining selected requests continue after individual failures.");
        // Let React paint progress before request construction or transport setup.
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
        let result: ExecutionRecord;
        try {
          const baseline = buildWith(group);
          const spec = groups.length === 1 && builtRequest.trim() ? applyEditableRequest(baseline, builtRequest, endpoint) : baseline;
          result = await runOne(spec, label, createExecutor());
        } catch (error) {
          notSent += 1;
          const failure = error instanceof Error ? error.message : String(error);
          if (notSentReasons.indexOf(failure) < 0) notSentReasons.push(failure);
          setBatchSummary({ attempted, succeeded, failed, notSent });
          setNotice(label + " was not sent: " + failure + ". Continuing with the next payload.");
          continue;
        }
        attempted += 1;
        if (!result.error && result.status >= 200 && result.status < 300) succeeded += 1;
        else {
          failed += 1;
        }
        setBatchSummary({ attempted, succeeded, failed, notSent });
        setNotice("Payload request " + (i + 1) + "/" + groups.length + " finished; continuing with the next selected request.");
      }
      setNotice(
        "Selected payload run complete: " + selectedPayloads.length + " case(s): " + attempted + " attempted, " + succeeded + " successful, " + failed + " failed after attempt, " + notSent + " not sent." +
        (notSentReasons.length ? " Not-sent reason(s): " + notSentReasons.slice(0, 3).join("; ") + (notSentReasons.length > 3 ? "; additional reasons omitted" : "") + "." : "")
      );
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
    setBatchSummary(null);
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
        await waitRandomInterval();
        const spec = build(plan.overrides);
        if (spec.destructive || spec.method !== "GET") throw new Error("Safety check stopped a non-read-only sweep request.");
        if (endpoint.family !== "microsoft-graph") assertAllowedUrl(spec.url, config);
        const result = await executor.execute(spec);
        lastExecutionAt.current = Date.now();
        rememberResponse(plan.label, spec, result);
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
  const primaryList = config.resources.lists.primary;
  const primaryListLabel = (primaryList.title || "not set") + (primaryList.id ? " · " + primaryList.id : "");
  const primaryItemId = String(config.resources.items.primaryItemId || "");
  const primaryItemChoice = primaryItemId
    ? (discoveredChoices.itemId || discoveredChoices.ID || []).filter((option) => option.key === primaryItemId)[0]
    : undefined;
  // Discovered choices read "Title · ID"; fall back to the bare ID when the
  // item has not been discovered yet.
  const primaryItemLabel = primaryItemId ? (primaryItemChoice ? primaryItemChoice.text : "#" + primaryItemId) : "not set";
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
    const responseText = record.responseBody || "";
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
    setNotice("Response body copied to clipboard.");
  };
  const responseStatusOptions: IDropdownOption[] = Array.from(new Set(responseHistory.map((entry) => entry.record.status)))
      .sort((a, b) => a - b)
      .map((status) => ({ key: String(status), text: String(status) }));
  const responseSizeOptions: IDropdownOption[] = Array.from(new Set(responseHistory.map((entry) => responseSizeBytes(entry.record))))
      .sort((a, b) => a - b)
      .map((size) => ({ key: String(size), text: size + " bytes" }));
  const visibleResponseHistory = responseHistory.filter((entry) =>
    (!responseStatusFilter || String(entry.record.status) === responseStatusFilter) &&
    (!responseSizeFilter || String(responseSizeBytes(entry.record)) === responseSizeFilter))
    .slice()
    .sort((left, right) => responseSizeBytes(right.record) - responseSizeBytes(left.record) || left.record.status - right.record.status || left.id - right.id);
  const copyRawRequest = async (): Promise<void> => {
    try {
      const spec = selectedPayloads.length > 1 ? build() : requestFromPreview();
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
          <p>Edits are saved in this browser for this webpart and SharePoint web. Parent and subsite settings stay separate. Apply reloads the runner; Reset restores the checked-in config.</p>
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
        <span>Primary list: {primaryListLabel}</span>
        <span>Primary item: {primaryItemLabel}</span>
        <span>Configured primary: {config.principals.users.primary.upn || config.principals.users.primary.loginName || "not set"}{(config.principals.users.primary.upn || config.principals.users.primary.loginName) ? (accountMatch(config.principals.users.primary.upn || config.principals.users.primary.loginName) ? " · matches this session" : " · different account") : ""}</span>
        <span>Configured secondary: {config.principals.users.secondary.upn || config.principals.users.secondary.loginName || "not set"}{(config.principals.users.secondary.upn || config.principals.users.secondary.loginName) ? (accountMatch(config.principals.users.secondary.upn || config.principals.users.secondary.loginName) ? " · matches this session" : " · different account") : ""}</span>
        <span>Requests use the signed-in account above; the configured identities are comparison references.</span>
      </div>
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
        <div className="researchApiPanel__payloads">
          <h3>Bug-family payload sets</h3>
          <p>Select one or more bug types, then select payload cases. Each selected case runs in its own request, one at a time.</p>
          <div className="researchApiPanel__payloadSelectors">
            <div>
              <Dropdown multiSelect label="Bug type" placeholder="Select payload families" options={payloadFamilyOptions} selectedKeys={payloadFamilyKeys} onChange={(_e, option) => { const keys = toggleDropdownKey(payloadFamilyKeys, option) as PayloadFamily[]; setPayloadFamilyKeys(keys); setPayloadFamily((option && option.selected ? String(option.key) : keys[keys.length - 1] || "") as PayloadFamily | ""); setPayloadId(""); setPayloadKeys([]); }} />
              <DefaultButton text="Select all types" onClick={selectAllPayloadFamilies} disabled={busy || endpoint.payloadFamilies.length === 0} />
            </div>
            <div>
              <Dropdown multiSelect label={"Payload case - " + payloadCases.length + " generated"} placeholder={payloadFamilyKeys.length > 0 ? "Select cases to apply" : "Select a bug type first"} options={payloadCaseOptions} selectedKeys={payloadKeys} disabled={payloadFamilyKeys.length === 0 && !payloadFamily} onChange={(_e, option) => { const keys = toggleDropdownKey(payloadKeys, option); setPayloadKeys(keys); setPayloadId(option && option.selected ? String(option.key) : keys[keys.length - 1] || ""); }} />
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <DefaultButton text="Select all cases" onClick={selectAllPayloadCases} disabled={busy || payloadCases.length === 0} />
                <DefaultButton text="Clear" onClick={clearPayloadCases} disabled={busy || payloadKeys.length === 0} />
              </Stack>
            </div>
          </div>
          {selectedPayloads.length > 0 && <p>{selectedPayloads.length} cases selected across {new Set(selectedPayloads.map((payload) => payload.family)).size} bug type(s). Each case runs separately.</p>}
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
            // Live discovery already knows the entity name for each ID, so it
            // wins over a configured variant that only has an ID.
            const discovered = discoveredChoices[parameter.name] || [];
            const discoveredKeys: Record<string, boolean> = {};
            discovered.forEach((option) => { discoveredKeys[option.key] = true; });
            const combined = discovered.concat(
              variants.filter((variant) => !discoveredKeys[variant.value]).map((variant) => ({ key: variant.value, text: variant.label + " · " + variant.value }))
            );
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
        {isDynamicSiteScriptsModule && <div className="researchApiPanel__siteScriptAction">
          <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
            <Dropdown label="Site script from SharePoint" placeholder="Select a script" options={dynamicSiteScriptOptions} selectedKey={selectedSiteScriptId || undefined} onChange={(_event, option) => selectSiteScript(option ? String(option.key) : "")} />
            <Dropdown label="Site design from SharePoint" placeholder="Select a design" options={dynamicSiteDesignOptions} selectedKey={selectedSiteDesignId || undefined} onChange={(_event, option) => { void selectSiteDesign(option ? String(option.key) : ""); }} />
            <DefaultButton text={siteScriptCatalogLoading ? "Loading…" : "Refresh IDs"} onClick={() => { void refreshSiteScriptCatalog(); }} disabled={siteScriptCatalogLoading} />
          </Stack>
          {isSiteDesignCreateOrUpdate && <Dropdown multiSelect label="Linked site scripts" placeholder="Select scripts for this design" options={dynamicSiteScriptOptions} selectedKeys={selectedSiteScriptIds} onChange={(_event, option) => setSelectedSiteScriptIds(toggleDropdownKey(selectedSiteScriptIds, option))} />}
          {(endpoint.id === "siteDesigns.getTask" || endpoint.id === "siteDesigns.runStatus") && <TextField ariaLabel="Task or run ID" placeholder="Task or run ID returned by SharePoint" value={selectedSiteDesignTaskId} onChange={(_event, value) => setSelectedSiteDesignTaskId(value || "")} />}
          {isSiteScriptCollectionEditor && <>
            <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
              <Dropdown label="Add action" options={siteScriptActionVerbOptions} selectedKey={siteScriptActionToAdd} onChange={(_event, option) => { if (option) setSiteScriptActionToAdd(String(option.key) as SiteScriptVerb); }} />
              <DefaultButton text="Add action" onClick={addSiteScriptAction} disabled={busy} />
            </Stack>
            <TextField ariaLabel="Site script actions JSON array" placeholder="Actions loaded from SharePoint or edit the action definitions" multiline rows={12} value={siteScriptActionsText} onChange={(_event, value) => setSiteScriptActionsText(value || "[]")} />
          </>}
          {siteScriptCatalogError && <MessageBar messageBarType={MessageBarType.error}>{siteScriptCatalogError}</MessageBar>}
        </div>}
        {endpoint.id === "siteScripts.executeAction" && <div className="researchApiPanel__siteScriptAction">
          {loadedSiteScriptActionOptions.length > 0 && <Dropdown label="Action from selected site script" placeholder="Choose a loaded action" options={loadedSiteScriptActionOptions} selectedKey={selectedSiteScriptActionIndex >= 0 ? selectedSiteScriptActionIndex : undefined} onChange={(_event, option) => {
            if (!option) return;
            const index = Number(option.key);
            const action = loadedSiteScriptActions[index];
            if (!action) return;
            setSelectedSiteScriptActionIndex(index);
            setSiteScriptActionText(JSON.stringify(action, null, 2));
            if (action && typeof action === "object" && "verb" in action && SITE_SCRIPT_VERBS.indexOf((action as any).verb) >= 0) {
              const verb = (action as any).verb as SiteScriptVerb;
              setSiteScriptVerb(verb);
              setSiteScriptSubactionVerb(verb === "createContentType" ? "addSiteColumn" : "setTitle");
            }
            setSiteScriptActionError("");
          }} />}
          <Dropdown label="Site script verb" options={siteScriptVerbOptions} selectedKey={siteScriptVerb} onChange={(_event, option) => { if (option) selectSiteScriptVerb(String(option.key) as SiteScriptVerb); }} />
          {(siteScriptVerb === "createSPList" || siteScriptVerb === "createContentType") && <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
            <Dropdown label="Subaction" options={siteScriptSubactionOptions} selectedKey={siteScriptSubactionVerb} styles={{ root: { minWidth: 340, width: 340 }, dropdown: { minWidth: 340 } }} onChange={(_event, option) => { if (option) setSiteScriptSubactionVerb(String(option.key)); }} />
            <DefaultButton text="Add subaction" onClick={addSiteScriptSubaction} disabled={busy} />
          </Stack>}
          {(siteScriptVerb === "createSPList" || siteScriptVerb === "createContentType") && <span>Subactions run within this parent action. SharePoint does not execute them as standalone actions.</span>}
          <TextField ariaLabel="Site script action definition JSON" placeholder="Edit the selected verb parameters as JSON" multiline rows={12} value={siteScriptActionText} onChange={(_event, value) => {
            const text = value || "";
            setSiteScriptActionText(text);
            try { const actionDefinition = parseSiteScriptActionDefinition(text); setSiteScriptVerb(actionDefinition.verb); setSiteScriptSubactionVerb(actionDefinition.verb === "createContentType" ? "addSiteColumn" : "setTitle"); setSiteScriptActionError(""); }
            catch (error) { setSiteScriptActionError(error instanceof Error ? error.message : String(error)); }
          }} />
          {siteScriptActionError && <MessageBar messageBarType={MessageBarType.error}>{siteScriptActionError}</MessageBar>}
        </div>}
        {bodyRequired && endpoint.id !== "siteScripts.executeAction" && !isSiteScriptCollectionEditor && <TextField ariaLabel="Request body" placeholder={"Request body · " + (endpoint.requestBody?.contentType || "application/json")} multiline rows={8} value={bodyText} onChange={(_e, value) => setBodyText(value || "")} />}
        {(endpoint.requiredHeaders || []).filter((name) => !/^(authorization|cookie|x-requestdigest)$/i.test(name)).map((name) => <TextField key={name} ariaLabel={"Header " + name} placeholder={name} value={headers[name] || ""} onChange={(_e, value) => setHeaders((current) => ({ ...current, [name]: value || "" }))} />)}
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <DefaultButton text="Preview request" onClick={previewRequest} disabled={busy || selectedPayloads.length > 1} />
          <DefaultButton text="Copy raw HTTP" onClick={() => { void copyRawRequest(); }} disabled={busy} />
          <PrimaryButton text="Execute API call" onClick={() => { void execute(); }} disabled={busy} />
          <PrimaryButton text={"Execute selected (" + selectedPayloads.length + " requests)"} onClick={() => { void executeSelectedPayloads(); }} disabled={busy || selectedPayloads.length === 0} />
          <DefaultButton text="Sweep available IDs (read only)" onClick={() => { void executeConfiguredIdSweep(); }} disabled={busy || method !== "GET" || endpoint.destructive || !pathAndQuery.some((p) => idVariantsFor(p.name).length > 0)} />
          {busy && <Spinner size={SpinnerSize.small} label="Executing" />}
        </Stack>
        {batchSummary && <div className="researchApiPanel__batchResults"><Label>Payload run · {batchSummary.attempted} attempted · {batchSummary.succeeded} successful · {batchSummary.failed} failed after attempt · {batchSummary.notSent} not sent</Label></div>}
        {sweepResults.length > 0 && <div><Label>ID sweep results · {sweepResults.length} requests</Label><pre className="researchApiPanel__output">{sweepResults.join("\n")}</pre></div>}
        {selectedPayloads.length <= 1 && <div><Label>Request preview · editable</Label><TextField ariaLabel="Editable request preview" placeholder="Preview request as JSON, then edit method, URL, headers and body" multiline rows={12} value={builtRequest} onChange={(_event, value) => setBuiltRequest(value || "")} /></div>}
        {record && <div>
          <Stack horizontal verticalAlign="center">
            <Label>Response · HTTP {record.status} · {record.responseClass} · {record.durationMs} ms</Label>
            <IconButton iconProps={{ iconName: "Copy" }} title="Copy complete response" ariaLabel="Copy complete response" onClick={() => { void copyResponse(); }} />
          </Stack>
          <pre className="researchApiPanel__output researchApiPanel__fullResponse">{record.responseBody || "(empty response body)"}</pre>
        </div>}
        <div className="researchApiPanel__responseHistory">
          <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 8 }}>
            <DefaultButton text={(responseHistoryOpen ? "Hide" : "View") + " all responses"} onClick={() => setResponseHistoryOpen((open) => !open)} disabled={responseHistory.length === 0} />
            <Dropdown ariaLabel="Filter responses by status code" placeholder="All statuses" options={responseStatusOptions} selectedKey={responseStatusFilter || undefined} onChange={(_event, option) => setResponseStatusFilter(option ? String(option.key) : "")} disabled={responseHistory.length === 0} />
            <Dropdown ariaLabel="Filter responses by body size" placeholder="All sizes" options={responseSizeOptions} selectedKey={responseSizeFilter || undefined} onChange={(_event, option) => setResponseSizeFilter(option ? String(option.key) : "")} disabled={responseHistory.length === 0} />
            <Label>{responseHistory.length} stored</Label>
            <DefaultButton text="Clear responses" onClick={() => { setResponseHistory([]); setExpandedRequestIds({}); setRecord(null); setResponseHistoryOpen(false); setResponseStatusFilter(""); setResponseSizeFilter(""); setSweepResults([]); setBatchSummary(null); }} disabled={responseHistory.length === 0 && !record} />
          </Stack>
          {responseHistoryOpen && <>
            {visibleResponseHistory.length === 0 && <span>No responses match this status code.</span>}
            {visibleResponseHistory.map((entry) => {
              const requestExpanded = expandedRequestIds[entry.id] === true;
              return <div className="researchApiPanel__historyEntry" key={entry.id}>
                <Label>{entry.label} · HTTP {entry.record.status} · {responseSizeBytes(entry.record)} bytes · {entry.record.responseClass}</Label>
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                  <IconButton iconProps={{ iconName: requestExpanded ? "ChevronDown" : "ChevronRight" }} title={requestExpanded ? "Hide sent request" : "View sent request"} ariaLabel={requestExpanded ? "Hide sent request" : "View sent request"} onClick={() => setExpandedRequestIds((current) => Object.assign({}, current, { [entry.id]: !current[entry.id] }))} />
                  <Label>Sent request · URL/query, headers and body</Label>
                </Stack>
                {requestExpanded && <pre className="researchApiPanel__output">{entry.requestPayload}</pre>}
                <Label>Response body</Label>
                <pre className="researchApiPanel__output researchApiPanel__fullResponse">{entry.record.responseBody || "(empty response body)"}</pre>
              </div>;
            })}
          </>}
        </div>
      </Stack>
      {(notice || choiceNotice) && <MessageBar messageBarType={record && record.status >= 200 && record.status < 300 ? MessageBarType.success : MessageBarType.warning}>{[notice, choiceNotice].filter(Boolean).join(" ")}</MessageBar>}
    </section>
    </>
  );
}
