/**
 * Site Script action definitions supported by Microsoft's
 * site-design-script-actions.schema.json. Kept in the existing research
 * request layer so the UI can validate and serialize ExecuteSiteScriptAction
 * inputs without adding another execution path.
 */

export const SITE_SCRIPT_VERBS = [
  "activateSPFeature", "addContentTypesFromHub", "associateExtension", "addPrincipalToSPGroup", "addNavLink",
  "applyTheme", "createContentType", "createSiteColumn", "createSiteColumnXml", "createSPList",
  "deactivateSPFeature", "installSolution", "joinHubSite", "removeNavLink", "setRegionalSettings",
  "setSiteBranding", "setSiteLogo", "setSiteLogoFromFile", "setSiteExternalSharingCapability", "triggerFlow", "setFooterLogo",
] as const;

export type SiteScriptVerb = typeof SITE_SCRIPT_VERBS[number];

export const SITE_SCRIPT_LIST_SUBACTION_VERBS = [
  "setTitle", "setDescription", "addSPField", "deleteSPField", "addSPFieldXml", "addSPLookupFieldXml",
  "addSiteColumn", "addSPView", "removeSPView", "addFile", "addFolder", "setDocumentTemplates",
  "addContentType", "removeContentType", "setClientFormCustomFormatter", "setSPFieldCustomFormatter",
  "associateFieldCustomizer", "associateListViewCommandSet",
] as const;

export type SiteScriptListSubactionVerb = typeof SITE_SCRIPT_LIST_SUBACTION_VERBS[number];
export type SiteScriptContentTypeSubactionVerb = "addSiteColumn" | "removeSiteColumn";
export type SiteScriptJsonObject = { [property: string]: unknown };

export interface SiteScriptListSubactionByVerb {
  setTitle: { verb: "setTitle"; title: string };
  setDescription: { verb: "setDescription"; description: string };
  addSPField: { verb: "addSPField"; fieldType: "Text" | "Note" | "Number" | "Boolean" | "User" | "DateTime"; displayName: string; internalName?: string; isRequired: boolean; addToDefaultView?: boolean; enforceUnique?: boolean; id?: string };
  deleteSPField: { verb?: "deleteSPField"; displayName: string };
  addSPFieldXml: { verb: "addSPFieldXml"; schemaXml: string; addToDefaultView?: boolean };
  addSPLookupFieldXml: { verb: "addSPLookupFieldXml"; schemaXml: string; targetListName?: string; targetListUrl?: string; addToDefaultView?: boolean };
  addSiteColumn: { verb: "addSiteColumn"; internalName: string; addToDefaultView?: boolean };
  addSPView: { verb?: "addSPView"; name: string; viewFields: string[]; query: string; rowLimit: number; isPaged: boolean; makeDefault?: boolean; scope?: "Default" | "Recursive" | "RecursiveAll" | "FilesOnly"; formatterJSON?: SiteScriptJsonObject; addLink?: string; viewType2?: string; viewDataXml?: string; replaceViewFields?: boolean; columnWidthXml?: string };
  removeSPView: { verb?: "removeSPView"; name: string };
  addFile: { verb: "addFile"; destination: string; source?: string; sourceUrl?: string };
  addFolder: { verb: "addFolder"; path: string };
  setDocumentTemplates: { verb: "setDocumentTemplates"; templateJSON: SiteScriptJsonObject };
  addContentType: { verb: "addContentType"; name: string; id?: string };
  removeContentType: { verb: "removeContentType"; name: string };
  setClientFormCustomFormatter: { verb: "setClientFormCustomFormatter"; clientFormCustomFormatter: SiteScriptJsonObject };
  setSPFieldCustomFormatter: { verb: "setSPFieldCustomFormatter"; fieldDisplayName: string; formatterJSON: SiteScriptJsonObject };
  associateFieldCustomizer: { verb: "associateFieldCustomizer"; internalName: string; clientSideComponentId: string; clientSideComponentProperties?: string };
  associateListViewCommandSet: { verb: "associateListViewCommandSet"; title: string; location: "ClientSideExtension.ListViewCommandSet.ContextMenu" | "ClientSideExtension.ListViewCommandSet.CommandBar" | "ClientSideExtension.ListViewCommandSet"; clientSideComponentId: string; clientSideComponentProperties?: string };
}

export type SiteScriptListSubactionDefinition = SiteScriptListSubactionByVerb[SiteScriptListSubactionVerb];
export type SiteScriptContentTypeSubactionDefinition =
  | { verb: "addSiteColumn"; internalName: string }
  | { verb: "removeSiteColumn"; internalName: string };

export interface SiteScriptActionByVerb {
  activateSPFeature: { verb: "activateSPFeature"; featureId: string; scope?: "web" | "site" };
  addContentTypesFromHub: { verb: "addContentTypesFromHub"; ids?: string[] };
  associateExtension: { verb: "associateExtension"; title: string; location: string; clientSideComponentId: string; clientSideComponentProperties?: string; registrationId?: string; registrationType?: "" | "List"; scope: "Web" | "Site" };
  addPrincipalToSPGroup: { verb: "addPrincipalToSPGroup"; principal: string; group: "Members" | "Owners" | "Visitors" };
  addNavLink: { verb: "addNavLink"; url: string; displayName: string; isWebRelative?: boolean; parentDisplayName?: string; parentUrl?: string; isParentUrlWebRelative?: boolean; navComponent?: "QuickLaunch" | "Hub" | "Footer" };
  applyTheme: { verb: "applyTheme"; themeName?: string; themeJson?: SiteScriptJsonObject };
  createContentType: { verb: "createContentType"; name: string; description?: string; parentName?: string; parentId?: string; id?: string; hidden: boolean; group?: string; subactions?: SiteScriptContentTypeSubactionDefinition[] };
  createSiteColumn: { verb: "createSiteColumn"; fieldType: "Text" | "Note" | "Number" | "Boolean" | "User" | "DateTime"; internalName: string; displayName?: string; isRequired: boolean; group?: string; enforceUnique?: boolean; id?: string };
  createSiteColumnXml: { verb: "createSiteColumnXml"; schemaXml: string; pushChanges?: boolean };
  createSPList: { verb: "createSPList"; listName: string; templateType: number; addNavLink?: boolean; hidden?: boolean; color?: string; icon?: string; fieldRefsXml?: string[]; output?: { webRelativeUrl?: string; absoluteUrl?: string }; subactions?: SiteScriptListSubactionDefinition[] };
  deactivateSPFeature: { verb: "deactivateSPFeature"; featureId: string; scope?: "web" | "site" };
  installSolution: { verb: "installSolution"; id: string; name: string };
  joinHubSite: { verb: "joinHubSite"; hubSiteId: string; name?: string };
  removeNavLink: { verb: "removeNavLink"; displayName: string; url?: string; isWebRelative?: boolean; navComponent?: "QuickLaunch" | "Hub" | "Footer" };
  setRegionalSettings: { verb: "setRegionalSettings"; sortOrder: number; timeZone?: number; locale?: number; hourFormat?: "12" | "24" };
  setSiteBranding: { verb: "setSiteBranding"; navigationLayout?: "Cascade" | "Megamenu"; headerLayout?: "Standard" | "Compact"; headerBackground?: "None" | "Neutral" | "Soft" | "Strong"; showFooter?: boolean };
  setSiteLogo: { verb: "setSiteLogo"; url: string };
  setSiteLogoFromFile: { verb: "setSiteLogoFromFile"; source?: string; sourceUrl?: string; templateImageUrl?: string };
  setSiteExternalSharingCapability: { verb: "setSiteExternalSharingCapability"; capability: "Disabled" | "ExistingExternalUserSharingOnly" | "ExternalUserSharingOnly" | "ExternalUserAndGuestSharing" };
  triggerFlow: { verb: "triggerFlow"; name: string; url: string; waitForReply?: boolean; parameters?: SiteScriptJsonObject };
  setFooterLogo: { verb: "setFooterLogo"; imageUrl: string; imageUrlType?: "WebRelativeUrl" | "TemplateImage" };
}

export type SiteScriptActionDefinition = SiteScriptActionByVerb[SiteScriptVerb];

type JsonType = "string" | "boolean" | "number" | "integer" | "object" | "array";
interface PropertyRule {
  type: JsonType;
  enum?: Array<string | number | boolean>;
  pattern?: string;
  items?: PropertyRule;
  properties?: { [name: string]: PropertyRule };
  additionalProperties?: boolean;
}
interface ActionRule {
  required: string[];
  properties: { [name: string]: PropertyRule };
  additionalProperties?: boolean;
  subactions?: "list" | "contentType";
}

const str = (pattern?: string, values?: string[]): PropertyRule => ({ type: "string", pattern, enum: values });
const bool: PropertyRule = { type: "boolean" };
const num: PropertyRule = { type: "number" };
const integer: PropertyRule = { type: "integer" };
const obj: PropertyRule = { type: "object" };
const stringArray: PropertyRule = { type: "array", items: str() };
const stringVerb = (verb: string): PropertyRule => str("^" + verb + "$", [verb]);
const subactionRule = (verb: string, required: string[], properties: { [name: string]: PropertyRule }): ActionRule => ({
  required,
  additionalProperties: false,
  properties: Object.assign({ verb: stringVerb(verb) }, properties),
});

const listSubactionRules: { [verb in SiteScriptListSubactionVerb]: ActionRule } = {
  setTitle: subactionRule("setTitle", ["title", "verb"], { title: str() }),
  setDescription: subactionRule("setDescription", ["description", "verb"], { description: str() }),
  addSPField: subactionRule("addSPField", ["displayName", "fieldType", "isRequired", "verb"], {
    fieldType: str(undefined, ["Text", "Note", "Number", "Boolean", "User", "DateTime"]), displayName: str(), internalName: str(),
    isRequired: bool, addToDefaultView: bool, enforceUnique: bool, id: str(),
  }),
  deleteSPField: subactionRule("deleteSPField", ["displayName"], { displayName: str() }),
  addSPFieldXml: subactionRule("addSPFieldXml", ["schemaXml", "verb"], { schemaXml: str(), addToDefaultView: bool }),
  addSPLookupFieldXml: subactionRule("addSPLookupFieldXml", ["schemaXml", "verb"], {
    schemaXml: str(), targetListName: str(), targetListUrl: str(), addToDefaultView: bool,
  }),
  addSiteColumn: subactionRule("addSiteColumn", ["internalName", "verb"], { internalName: str(), addToDefaultView: bool }),
  addSPView: subactionRule("addSPView", ["name", "viewFields", "query", "rowLimit", "isPaged"], {
    name: str(), viewFields: stringArray, query: str(), rowLimit: integer, isPaged: bool, makeDefault: bool,
    scope: str("^(Default|Recursive|RecursiveAll|FilesOnly)$"), formatterJSON: obj, addLink: str(), viewType2: str(),
    viewDataXml: str(), replaceViewFields: bool, columnWidthXml: str(),
  }),
  removeSPView: subactionRule("removeSPView", ["name"], { name: str() }),
  addFile: subactionRule("addFile", ["verb", "destination"], { destination: str(), source: str(), sourceUrl: str() }),
  addFolder: subactionRule("addFolder", ["verb", "path"], { path: str() }),
  setDocumentTemplates: subactionRule("setDocumentTemplates", ["verb", "templateJSON"], { templateJSON: obj }),
  addContentType: subactionRule("addContentType", ["name", "verb"], { name: str(), id: str() }),
  removeContentType: subactionRule("removeContentType", ["name", "verb"], { name: str() }),
  setClientFormCustomFormatter: subactionRule("setClientFormCustomFormatter", ["clientFormCustomFormatter", "verb"], { clientFormCustomFormatter: obj }),
  setSPFieldCustomFormatter: subactionRule("setSPFieldCustomFormatter", ["fieldDisplayName", "verb", "formatterJSON"], { fieldDisplayName: str(), formatterJSON: obj }),
  associateFieldCustomizer: subactionRule("associateFieldCustomizer", ["clientSideComponentId", "internalName", "verb"], {
    internalName: str(), clientSideComponentId: str(), clientSideComponentProperties: str(),
  }),
  associateListViewCommandSet: subactionRule("associateListViewCommandSet", ["clientSideComponentId", "title", "location", "verb"], {
    title: str(), location: str("^(ClientSideExtension\\.ListViewCommandSet\\.ContextMenu|ClientSideExtension\\.ListViewCommandSet\\.CommandBar|ClientSideExtension\\.ListViewCommandSet)$"),
    clientSideComponentId: str(), clientSideComponentProperties: str(),
  }),
};

const contentTypeSubactionRules: { [verb in SiteScriptContentTypeSubactionVerb]: ActionRule } = {
  addSiteColumn: subactionRule("addSiteColumn", ["internalName", "verb"], { internalName: str() }),
  removeSiteColumn: subactionRule("removeSiteColumn", ["internalName", "verb"], { internalName: str() }),
};

const listSubactionNames = SITE_SCRIPT_LIST_SUBACTION_VERBS as readonly string[];

const actionRules: { [verb in SiteScriptVerb]: ActionRule } = {
  activateSPFeature: { required: ["featureId", "verb"], properties: { verb: stringVerb("activateSPFeature"), featureId: str(), scope: str(undefined, ["web", "site"]) } },
  addContentTypesFromHub: { required: ["verb"], additionalProperties: false, properties: { verb: stringVerb("addContentTypesFromHub"), ids: stringArray } },
  associateExtension: { required: ["clientSideComponentId", "title", "location", "verb", "scope"], additionalProperties: false, properties: {
    verb: stringVerb("associateExtension"), title: str(), location: str("^ClientSideExtension\\..*$"), clientSideComponentId: str(),
    clientSideComponentProperties: str(), registrationId: str(), registrationType: str("^$|^List$"), scope: str("^(Web|Site)$"),
  } },
  addPrincipalToSPGroup: { required: ["principal", "group", "verb"], properties: {
    verb: stringVerb("addPrincipalToSPGroup"), principal: str(), group: str("^(Members|Owners|Visitors)$"),
  } },
  addNavLink: { required: ["displayName", "url", "verb"], additionalProperties: false, properties: {
    verb: stringVerb("addNavLink"), url: str(), displayName: str(), isWebRelative: bool, parentDisplayName: str(), parentUrl: str(),
    isParentUrlWebRelative: bool, navComponent: str("^(QuickLaunch|Hub|Footer)$"),
  } },
  applyTheme: { required: ["verb"], properties: { verb: stringVerb("applyTheme"), themeName: str(), themeJson: obj } },
  createContentType: { required: ["verb", "name", "hidden"], additionalProperties: false, subactions: "contentType", properties: {
    verb: stringVerb("createContentType"), name: str(), description: str(), parentName: str(), parentId: str(), id: str(), hidden: bool,
    group: str(), subactions: { type: "array", items: obj },
  } },
  createSiteColumn: { required: ["fieldType", "verb", "internalName", "isRequired"], additionalProperties: false, properties: {
    verb: stringVerb("createSiteColumn"), fieldType: str(undefined, ["Text", "Note", "Number", "Boolean", "User", "DateTime"]),
    internalName: str(), displayName: str(), isRequired: bool, group: str(), enforceUnique: bool, id: str(),
  } },
  createSiteColumnXml: { required: ["verb", "schemaXml"], additionalProperties: false, properties: {
    verb: stringVerb("createSiteColumnXml"), schemaXml: str(), pushChanges: bool,
  } },
  createSPList: { required: ["listName", "templateType", "verb"], additionalProperties: false, subactions: "list", properties: {
    verb: stringVerb("createSPList"), listName: str(), templateType: num, addNavLink: bool, hidden: bool, color: str(), icon: str(),
    fieldRefsXml: stringArray, output: { type: "object", additionalProperties: false, properties: { webRelativeUrl: str(), absoluteUrl: str() } },
    subactions: { type: "array", items: obj },
  } },
  deactivateSPFeature: { required: ["featureId", "verb"], properties: { verb: stringVerb("deactivateSPFeature"), featureId: str(), scope: str(undefined, ["web", "site"]) } },
  installSolution: { required: ["id", "verb", "name"], additionalProperties: false, properties: { verb: stringVerb("installSolution"), id: str(), name: str() } },
  joinHubSite: { required: ["hubSiteId", "verb"], additionalProperties: false, properties: { verb: stringVerb("joinHubSite"), hubSiteId: str(), name: str() } },
  removeNavLink: { required: ["displayName", "verb"], additionalProperties: false, properties: {
    verb: stringVerb("removeNavLink"), url: str(), displayName: str(), isWebRelative: bool, navComponent: str("^(QuickLaunch|Hub|Footer)$"),
  } },
  setRegionalSettings: { required: ["sortOrder", "verb"], additionalProperties: false, properties: {
    verb: stringVerb("setRegionalSettings"), timeZone: integer, locale: integer, sortOrder: integer, hourFormat: str("^(12|24)$"),
  } },
  setSiteBranding: { required: ["verb"], additionalProperties: false, properties: {
    verb: stringVerb("setSiteBranding"), navigationLayout: str("^(Cascade|Megamenu)$"), headerLayout: str("^(Standard|Compact)$"),
    headerBackground: str("^(None|Neutral|Soft|Strong)$"), showFooter: bool,
  } },
  setSiteLogo: { required: ["url", "verb"], properties: { verb: stringVerb("setSiteLogo"), url: str() } },
  setSiteLogoFromFile: { required: ["verb"], properties: { verb: stringVerb("setSiteLogoFromFile"), source: str(), sourceUrl: str(), templateImageUrl: str() } },
  setSiteExternalSharingCapability: { required: ["capability", "verb"], additionalProperties: false, properties: {
    verb: stringVerb("setSiteExternalSharingCapability"), capability: str(undefined, ["Disabled", "ExistingExternalUserSharingOnly", "ExternalUserSharingOnly", "ExternalUserAndGuestSharing"]),
  } },
  triggerFlow: { required: ["name", "url", "verb"], additionalProperties: false, properties: {
    verb: stringVerb("triggerFlow"), url: str(), name: str(), waitForReply: bool, parameters: obj,
  } },
  setFooterLogo: { required: ["verb", "imageUrl"], properties: {
    verb: stringVerb("setFooterLogo"), imageUrl: str(), imageUrlType: str(undefined, ["WebRelativeUrl", "TemplateImage"]),
  } },
};

const samples: { [verb in SiteScriptVerb]: SiteScriptActionDefinition } = {
  activateSPFeature: { verb: "activateSPFeature", featureId: "00000000-0000-0000-0000-000000000000", scope: "web" },
  addContentTypesFromHub: { verb: "addContentTypesFromHub", ids: ["0x0101"] },
  associateExtension: { verb: "associateExtension", title: "Research extension", location: "ClientSideExtension.ApplicationCustomizer", clientSideComponentId: "00000000-0000-0000-0000-000000000000", scope: "Web" },
  addPrincipalToSPGroup: { verb: "addPrincipalToSPGroup", principal: "research@example.invalid", group: "Members" },
  addNavLink: { verb: "addNavLink", displayName: "Research link", url: "/sites/research", isWebRelative: true },
  applyTheme: { verb: "applyTheme", themeName: "" },
  createContentType: { verb: "createContentType", name: "Research Content Type", hidden: false, subactions: [] },
  createSiteColumn: { verb: "createSiteColumn", fieldType: "Text", internalName: "ResearchColumn", displayName: "Research Column", isRequired: false },
  createSiteColumnXml: { verb: "createSiteColumnXml", schemaXml: "<Field Type=\"Text\" Name=\"ResearchColumn\" DisplayName=\"Research Column\" />", pushChanges: false },
  createSPList: { verb: "createSPList", listName: "Research Test List", templateType: 100, addNavLink: false, subactions: [] },
  deactivateSPFeature: { verb: "deactivateSPFeature", featureId: "00000000-0000-0000-0000-000000000000", scope: "web" },
  installSolution: { verb: "installSolution", id: "00000000-0000-0000-0000-000000000000", name: "Research Solution" },
  joinHubSite: { verb: "joinHubSite", hubSiteId: "00000000-0000-0000-0000-000000000000" },
  removeNavLink: { verb: "removeNavLink", displayName: "Research link" },
  setRegionalSettings: { verb: "setRegionalSettings", sortOrder: 500 },
  setSiteBranding: { verb: "setSiteBranding", headerLayout: "Standard" },
  setSiteLogo: { verb: "setSiteLogo", url: "https://example.invalid/logo.png" },
  setSiteLogoFromFile: { verb: "setSiteLogoFromFile", source: "SiteAssets/logo.png" },
  setSiteExternalSharingCapability: { verb: "setSiteExternalSharingCapability", capability: "Disabled" },
  triggerFlow: { verb: "triggerFlow", name: "Research Flow", url: "https://example.invalid/trigger", parameters: {} },
  setFooterLogo: { verb: "setFooterLogo", imageUrl: "", imageUrlType: "WebRelativeUrl" },
};

function isObject(value: unknown): value is { [key: string]: unknown } {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function validateRule(value: unknown, rule: PropertyRule, path: string): void {
  let valid = false;
  switch (rule.type) {
    case "string": valid = typeof value === "string"; break;
    case "boolean": valid = typeof value === "boolean"; break;
    case "number": valid = typeof value === "number" && isFinite(value); break;
    case "integer": valid = typeof value === "number" && isFinite(value) && Math.floor(value) === value; break;
    case "object": valid = isObject(value); break;
    case "array": valid = Array.isArray(value); break;
  }
  if (!valid) throw new Error(path + " must be " + rule.type + ".");
  if (rule.enum && rule.enum.indexOf(value as string | number | boolean) < 0) throw new Error(path + " must be one of: " + rule.enum.join(", ") + ".");
  if (rule.pattern && !(new RegExp(rule.pattern)).test(String(value))) throw new Error(path + " does not match the required pattern " + rule.pattern + ".");
  if (rule.type === "array" && rule.items) (value as unknown[]).forEach((item, index) => validateRule(item, rule.items as PropertyRule, path + "[" + index + "]"));
  if (rule.type === "object" && rule.properties) validateObject(value as { [key: string]: unknown }, { required: [], properties: rule.properties, additionalProperties: rule.additionalProperties }, path);
}

function validateObject(value: { [key: string]: unknown }, rule: ActionRule, path: string): void {
  rule.required.forEach((key) => {
    if (!Object.prototype.hasOwnProperty.call(value, key)) throw new Error(path + "." + key + " is required.");
  });
  Object.keys(value).forEach((key) => {
    const property = rule.properties[key];
    if (!property) {
      if (rule.additionalProperties === false) throw new Error(path + "." + key + " is not supported by the Microsoft site script schema.");
      return;
    }
    validateRule(value[key], property, path + "." + key);
  });
  if (rule.subactions) {
    const subactions = value.subactions;
    if (subactions !== undefined) {
      if (!Array.isArray(subactions)) throw new Error(path + ".subactions must be array.");
      subactions.forEach((subaction, index) => {
        if (!isObject(subaction)) throw new Error(path + ".subactions[" + index + "] must be object.");
        const verb = subaction.verb;
        if (rule.subactions === "list" && verb === undefined) {
          // The official schema does not mark `verb` required for deleteSPField,
          // addSPView, and removeSPView. Keep that schema behavior for nested
          // actions; definitions with a supplied verb still validate exactly.
          const possibleRules = listSubactionNames.map((name) => listSubactionRules[name as SiteScriptListSubactionVerb]);
          let matched = false;
          possibleRules.forEach((subRule) => {
            try { validateObject(subaction, subRule, path + ".subactions[" + index + "]"); matched = true; }
            catch (_error) { /* Try the other schema alternatives. */ }
          });
          if (!matched) throw new Error(path + ".subactions[" + index + "] does not match a supported list subaction.");
          return;
        }
        const subRule = rule.subactions === "list"
          ? (typeof verb === "string" && listSubactionNames.indexOf(verb) >= 0 ? listSubactionRules[verb as SiteScriptListSubactionVerb] : undefined)
          : (verb === "addSiteColumn" || verb === "removeSiteColumn" ? contentTypeSubactionRules[verb] : undefined);
        if (!subRule) throw new Error(path + ".subactions[" + index + "] has an unsupported or missing verb.");
        validateObject(subaction, subRule, path + ".subactions[" + index + "]");
      });
    }
  }
}

export function validateSiteScriptActionDefinition(value: unknown): asserts value is SiteScriptActionDefinition {
  if (!isObject(value)) throw new Error("Site script action definition must be a JSON object.");
  const verb = value.verb;
  if (typeof verb !== "string" || SITE_SCRIPT_VERBS.indexOf(verb as SiteScriptVerb) < 0) {
    throw new Error("Unsupported site script verb '" + String(verb || "") + "'.");
  }
  validateObject(value, actionRules[verb as SiteScriptVerb], "actionDefinition");
}

export function parseSiteScriptActionDefinition(text: string): SiteScriptActionDefinition {
  let value: unknown;
  try { value = JSON.parse(text); }
  catch (_error) { throw new Error("Site script action definition must be valid JSON."); }
  validateSiteScriptActionDefinition(value);
  return value;
}

export function serializeSiteScriptActionDefinition(value: unknown): string {
  validateSiteScriptActionDefinition(value);
  return JSON.stringify(value);
}

export function createSiteScriptActionDefinition<Verb extends SiteScriptVerb>(verb: Verb): SiteScriptActionByVerb[Verb] {
  return JSON.parse(JSON.stringify(samples[verb])) as SiteScriptActionByVerb[Verb];
}

export function createSiteScriptListSubaction<Verb extends SiteScriptListSubactionVerb>(verb: Verb): SiteScriptListSubactionByVerb[Verb] {
  const examples: { [subverb in SiteScriptListSubactionVerb]: { verb: string; [key: string]: unknown } } = {
    setTitle: { verb, title: "Research Test List" }, setDescription: { verb, description: "Research test list description" },
    addSPField: { verb, fieldType: "Text", displayName: "Research Field", internalName: "ResearchField", isRequired: false },
    deleteSPField: { verb, displayName: "Research Field" },
    addSPFieldXml: { verb, schemaXml: "<Field Type=\"Text\" Name=\"ResearchField\" DisplayName=\"Research Field\" />" },
    addSPLookupFieldXml: { verb, schemaXml: "<Field Type=\"Lookup\" Name=\"ResearchLookup\" DisplayName=\"Research Lookup\" />", targetListName: "Research Target" },
    addSiteColumn: { verb, internalName: "ResearchColumn", addToDefaultView: false },
    addSPView: { verb, name: "Research View", viewFields: ["Title"], query: "<Where></Where>", rowLimit: 30, isPaged: true },
    removeSPView: { verb, name: "Research View" }, addFile: { verb, destination: "Research/placeholder.txt", source: "placeholder.txt" },
    addFolder: { verb, path: "Research" }, setDocumentTemplates: { verb, templateJSON: {} },
    addContentType: { verb, name: "Item" }, removeContentType: { verb, name: "Item" },
    setClientFormCustomFormatter: { verb, clientFormCustomFormatter: {} },
    setSPFieldCustomFormatter: { verb, fieldDisplayName: "Research Field", formatterJSON: {} },
    associateFieldCustomizer: { verb, internalName: "ResearchField", clientSideComponentId: "00000000-0000-0000-0000-000000000000" },
    associateListViewCommandSet: { verb, title: "Research Command", location: "ClientSideExtension.ListViewCommandSet", clientSideComponentId: "00000000-0000-0000-0000-000000000000" },
  };
  return JSON.parse(JSON.stringify(examples[verb])) as SiteScriptListSubactionByVerb[Verb];
}

export function createSiteScriptContentTypeSubaction<Verb extends SiteScriptContentTypeSubactionVerb>(verb: Verb): Extract<SiteScriptContentTypeSubactionDefinition, { verb: Verb }> {
  return { verb, internalName: "ResearchColumn" } as Extract<SiteScriptContentTypeSubactionDefinition, { verb: Verb }>;
}
