import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DefaultButton, Dropdown, IDropdownOption, MessageBar, MessageBarType, Spinner, Stack } from "@fluentui/react";
import { loadTestConfig } from "../../config/loader";

interface ConfigSetupAssistantProps { context: WebPartContext; }
interface DiscoveredList {
  Id: string;
  Title: string;
  BaseTemplate: number;
  BaseType: number;
  ItemEntityTypeFullName?: string;
  RootFolder?: { ServerRelativeUrl?: string };
}
interface DiscoveredUser { Id: number; Title: string; LoginName: string; Email?: string; }
interface DiscoveredGroup { Id: number; Title: string; }
interface ListDetails {
  list: DiscoveredList;
  items: Array<{ Id: number; UniqueId: string; Title?: string }>;
  views: Array<{ Id: string; Title: string; DefaultView?: boolean }>;
  fields: Array<{ Id: string; Title: string; InternalName?: string }>;
  contentTypes: Array<{ StringId: string; Name: string }>;
}

function resultArray<T>(json: any): T[] {
  if (json && json.d && Array.isArray(json.d.results)) return json.d.results as T[];
  if (json && Array.isArray(json.value)) return json.value as T[];
  return [];
}

async function getJson(url: string, timeoutMs: number): Promise<any> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, {
      method: "GET",
      credentials: "same-origin",
      headers: { Accept: "application/json;odata=nometadata" },
      signal: controller.signal,
    });
    const text = await response.text();
    if (!response.ok) throw new Error("HTTP " + response.status + " while reading " + new URL(url).pathname + ".");
    return text ? JSON.parse(text) : {};
  } finally {
    clearTimeout(timer);
  }
}

function firstObject<T>(json: any): T {
  if (json && json.d) return json.d as T;
  return json as T;
}

function hasConfiguredTenant(siteUrl: string): boolean {
  try { return !!siteUrl && !new URL(siteUrl).hostname.toLowerCase().endsWith(".invalid"); }
  catch (_error) { return false; }
}

export default function ConfigSetupAssistant(props: ConfigSetupAssistantProps): React.ReactElement | null {
  const config = React.useMemo(() => loadTestConfig(), []);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>("");
  const [lists, setLists] = React.useState<DiscoveredList[]>([]);
  const [user, setUser] = React.useState<DiscoveredUser | null>(null);
  const [users, setUsers] = React.useState<DiscoveredUser[]>([]);
  const [secondaryUserId, setSecondaryUserId] = React.useState<string>(String(config.principals.users.secondary.id || ""));
  const [groups, setGroups] = React.useState<DiscoveredGroup[]>([]);
  const [primaryListId, setPrimaryListId] = React.useState<string>(config.resources.lists.primary.id || "");
  const [lookupListId, setLookupListId] = React.useState<string>(config.resources.lists.lookup ? config.resources.lists.lookup.id : "");
  const [selectedGroupId, setSelectedGroupId] = React.useState<string>(String(config.principals.groups.primaryGroupId || ""));
  const [details, setDetails] = React.useState<ListDetails | null>(null);
  const [sampleItemId, setSampleItemId] = React.useState<string>(String(config.resources.items.primaryItemId || ""));
  const [customContentTypeId, setCustomContentTypeId] = React.useState<string>(config.resources.contentTypes.customContentTypeId || "");
  const [loadingDetails, setLoadingDetails] = React.useState<boolean>(false);
  const [copied, setCopied] = React.useState<boolean>(false);
  const [attempted, setAttempted] = React.useState<boolean>(false);

  const contextInfo = (props.context as any).pageContext;
  const webUrl: string = contextInfo && contextInfo.web && contextInfo.web.absoluteUrl
    ? String(contextInfo.web.absoluteUrl).replace(/\/$/, "")
    : "";
  const siteUrl: string = contextInfo && contextInfo.site && contextInfo.site.absoluteUrl
    ? String(contextInfo.site.absoluteUrl).replace(/\/$/, "")
    : webUrl;
  const pageSiteId: string = contextInfo && contextInfo.site && contextInfo.site.id ? String(contextInfo.site.id) : "";
  const pageWebId: string = contextInfo && contextInfo.web && contextInfo.web.id ? String(contextInfo.web.id) : "";
  const contextTenantId: string = contextInfo && contextInfo.aadInfo && contextInfo.aadInfo.tenantId
    ? String(contextInfo.aadInfo.tenantId)
    : "";

  const configured = hasConfiguredTenant(config.environment.siteUrl);
  const currentHost = webUrl ? new URL(webUrl).hostname : "";
  const missing: string[] = [];
  if (!configured || !config.environment.tenantUrl) missing.push("environment.tenantUrl / environment.siteUrl");
  if (currentHost && config.safety.allowedHosts.indexOf(currentHost) < 0) missing.push("safety.allowedHosts (allow the current SharePoint host)");
  if (!config.environment.tenantId) missing.push("environment.tenantId");
  if (!config.environment.siteCollectionId) missing.push("environment.siteCollectionId");
  if (!config.environment.siteId) missing.push("environment.siteId");
  if (!config.environment.webId) missing.push("environment.webId");
  if (!config.resources.lists.primary.id || !config.resources.lists.primary.title) missing.push("resources.lists.primary (choose a list)");
  if (config.resources.lists.lookup && !config.resources.lists.lookup.id) missing.push("resources.lists.lookup (optional for lookup field calls)");
  ["documents", "sitePages", "siteAssets", "userInformationList"].forEach((key) => {
    if (!config.resources.lists[key] || !config.resources.lists[key].id) missing.push("resources.lists." + key + ".id");
  });
  if (!config.resources.items.primaryItemUniqueId) missing.push("resources.items.primaryItemUniqueId (choose a list item)");
  if (!config.resources.folders.serverRelativePath) missing.push("resources.folders.serverRelativePath (choose a test folder)");
  if (!config.resources.views.viewId || !config.resources.views.viewName) missing.push("resources.views.viewId / viewName");
  if (!config.resources.fields.titleFieldId) missing.push("resources.fields.titleFieldId");
  if (!config.resources.contentTypes.customContentTypeId) missing.push("resources.contentTypes.customContentTypeId (optional for content type calls)");
  if (!config.principals.users.primary.id || !config.principals.users.primary.upn) missing.push("principals.users.primary (identity calls)");
  if (!config.principals.users.secondary.id || !config.principals.users.secondary.upn) missing.push("principals.users.secondary (two-user tests; choose manually)");
  if (!config.principals.groups.primaryGroupId) missing.push("principals.groups.primaryGroupId (group calls)");
  ["ownerGroupId", "memberGroupId", "visitorGroupId"].forEach((key) => {
    if (!config.principals.groups[key]) missing.push("principals.groups." + key);
  });
  if (!config.principals.groups.shareWith.length) missing.push("principals.groups.shareWith (sharing calls)");
  if (!config.principals.groups.groupsToEnumerate.length) missing.push("principals.groups.groupsToEnumerate");

  React.useEffect(() => {
    if (attempted || missing.length === 0) return;
    setAttempted(true);
    if (!webUrl) {
      setError("Could not read the current SharePoint site URL from SPFx context. Open this web part on a SharePoint site to discover setup values.");
      return;
    }
    const currentHost = new URL(webUrl).hostname.toLowerCase();
    const configuredHost = configured ? new URL(config.environment.siteUrl).hostname.toLowerCase() : "";
    if (!configuredHost.endsWith(".invalid") && currentHost !== configuredHost) {
      setError("The current SharePoint site (" + currentHost + ") differs from environment.siteUrl (" + configuredHost + "). Discovery was skipped to avoid querying the wrong tenant.");
      return;
    }
    setLoading(true);
    const root = webUrl + "/_api/web";
    const pending: Array<Promise<void>> = [];
    if (!config.resources.lists.primary.id || !config.resources.lists.primary.title) {
      pending.push(getJson(root + "/lists?$select=Id,Title,BaseTemplate,BaseType,ItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder&$filter=Hidden%20eq%20false&$top=200", config.safety.requestTimeoutMs)
        .then((json) => { setLists(resultArray<DiscoveredList>(json)); })
        .catch((reason) => { setError((previous) => previous ? previous + " " + String(reason) : String(reason)); }));
    }
    if (!config.principals.users.primary.id || !config.principals.users.primary.upn) {
      pending.push(getJson(root + "/currentuser?$select=Id,Title,LoginName,Email", config.safety.requestTimeoutMs)
        .then((json) => { const current = firstObject<DiscoveredUser>(json); setUser(current); setUsers((previous) => previous.some((candidate) => candidate.Id === current.Id) ? previous : [current].concat(previous)); })
        .catch((reason) => { setError((previous) => previous ? previous + " " + String(reason) : String(reason)); }));
    }
    if (!config.principals.users.secondary.id || !config.principals.users.secondary.upn) {
      pending.push(getJson(root + "/siteusers?$select=Id,Title,LoginName,Email&$top=100", config.safety.requestTimeoutMs)
        .then((json) => { setUsers((previous) => {
          const found = resultArray<DiscoveredUser>(json);
          const combined = previous.concat(found);
          return combined.filter((candidate, index) => combined.findIndex((other) => other.Id === candidate.Id) === index);
        }); })
        .catch((reason) => { setError((previous) => previous ? previous + " " + String(reason) : String(reason)); }));
    }
    if (!config.principals.groups.primaryGroupId) {
      pending.push(getJson(root + "/sitegroups?$select=Id,Title&$top=200", config.safety.requestTimeoutMs)
        .then((json) => { setGroups(resultArray<DiscoveredGroup>(json)); })
        .catch((reason) => { setError((previous) => previous ? previous + " " + String(reason) : String(reason)); }));
    }
    Promise.all(pending).finally(() => setLoading(false));
  }, []);

  React.useEffect(() => {
    if (!primaryListId || !webUrl) { setDetails(null); return; }
    const id = primaryListId.replace(/[{}]/g, "");
    const root = webUrl + "/_api/web/lists(guid'" + id + "')";
    const timeoutMs = config.safety.requestTimeoutMs;
    setDetails(null);
    setSampleItemId("");
    setCustomContentTypeId("");
    setLoadingDetails(true);
    const discoveryWarnings: string[] = [];
    Promise.all([
      getJson(root + "?$select=Id,Title,BaseTemplate,BaseType,ItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder", timeoutMs),
      getJson(root + "/items?$select=Id,UniqueId,Title&$top=10", timeoutMs),
      getJson(root + "/views?$select=Id,Title,DefaultView&$top=100", timeoutMs),
      getJson(root + "/fields?$select=Id,Title,InternalName&$filter=Hidden%20eq%20false&$top=200", timeoutMs),
      getJson(root + "/contenttypes?$select=StringId,Name&$top=100", timeoutMs),
    ].map((request) => request.catch((reason) => {
      discoveryWarnings.push(reason instanceof Error ? reason.message : String(reason));
      return null;
    }))).then((responses) => {
      if (!responses[0]) throw new Error(discoveryWarnings.join(" "));
      const foundItems = responses[1] ? resultArray<ListDetails["items"][number]>(responses[1]) : [];
      if (foundItems.length) setSampleItemId(String(foundItems[0].Id));
      setDetails({
        list: firstObject<DiscoveredList>(responses[0]),
        items: foundItems,
        views: responses[2] ? resultArray<ListDetails["views"][number]>(responses[2]) : [],
        fields: responses[3] ? resultArray<ListDetails["fields"][number]>(responses[3]) : [],
        contentTypes: responses[4] ? resultArray<ListDetails["contentTypes"][number]>(responses[4]) : [],
      });
      setError(discoveryWarnings.join(" "));
    }).catch((reason) => setError(reason instanceof Error ? reason.message : String(reason)))
      .finally(() => setLoadingDetails(false));
  }, [primaryListId]);

  if (missing.length === 0) return null;

  const listOptions: IDropdownOption[] = lists.map((list) => ({ key: list.Id, text: list.Title + "  ·  " + list.Id }));
  const primary = details && (details.items.find((item) => String(item.Id) === sampleItemId) || details.items[0]);
  const secondary = details && details.items.find((item) => !primary || item.Id !== primary.Id);
  const titleField = details ? details.fields.filter((field) => field.InternalName === "Title" || field.Title === "Title")[0] : undefined;
  const defaultView = details ? details.views.filter((view) => view.DefaultView)[0] : undefined;
  const guessedGroups = {
    owner: groups.find((group) => /owner/i.test(group.Title)),
    member: groups.find((group) => /member/i.test(group.Title)),
    visitor: groups.find((group) => /visitor/i.test(group.Title)),
  };
  const chosenList = details && details.list;
  const patch: Record<string, unknown> = {};
  const environmentPatch: Record<string, unknown> = {};
  if ((!configured || !config.environment.tenantUrl) && siteUrl) environmentPatch.tenantUrl = new URL(siteUrl).origin;
  if (!configured && webUrl) environmentPatch.siteUrl = webUrl;
  if (!config.environment.tenantId && contextTenantId) environmentPatch.tenantId = contextTenantId;
  if (!config.environment.siteCollectionId && pageSiteId) environmentPatch.siteCollectionId = pageSiteId;
  if (!config.environment.siteId && pageSiteId) environmentPatch.siteId = pageSiteId;
  if (!config.environment.webId && pageWebId) environmentPatch.webId = pageWebId;
  if ((!config.environment.webUrl || config.environment.webUrl === "/") && webUrl) environmentPatch.webUrl = new URL(webUrl).pathname || "/";
  if (Object.keys(environmentPatch).length) patch.environment = environmentPatch;
  if (currentHost && (configured || !environmentPatch.tenantUrl)) {
    const existingHosts = (config.safety.allowedHosts || []).filter((host) => host !== "example.invalid");
    if (existingHosts.indexOf(currentHost) < 0) patch.safety = { allowedHosts: existingHosts.concat([currentHost]) };
  } else if (currentHost && environmentPatch.tenantUrl) {
    patch.safety = { allowedHosts: [currentHost] };
  }
  if (chosenList) {
    const listsPatch: Record<string, unknown> = {};
    listsPatch.primary = {
      id: chosenList.Id,
      title: chosenList.Title,
      template: chosenList.BaseTemplate,
      baseType: chosenList.BaseType,
      itemEntityType: chosenList.ItemEntityTypeFullName || "",
      serverRelativeUrl: chosenList.RootFolder && chosenList.RootFolder.ServerRelativeUrl || "",
    };
    const lookup = lists.find((list) => list.Id === lookupListId);
    if (lookup && lookup.Id !== chosenList.Id) {
      listsPatch.lookup = {
        id: lookup.Id,
        title: lookup.Title,
        template: lookup.BaseTemplate,
        baseType: lookup.BaseType,
        itemEntityType: lookup.ItemEntityTypeFullName || "",
        serverRelativeUrl: lookup.RootFolder && lookup.RootFolder.ServerRelativeUrl || "",
      };
    }
    const standardLists: Array<{ key: string; matches: (list: DiscoveredList) => boolean }> = [
      { key: "documents", matches: (list) => list.BaseTemplate === 101 && /\/Shared Documents$/i.test(list.RootFolder?.ServerRelativeUrl || "") },
      { key: "sitePages", matches: (list) => list.Title.toLowerCase() === "site pages" },
      { key: "siteAssets", matches: (list) => list.Title.toLowerCase() === "site assets" },
      { key: "userInformationList", matches: (list) => list.Title.toLowerCase() === "user information list" },
    ];
    standardLists.forEach((definition) => {
      const found = lists.find(definition.matches);
      const existing = config.resources.lists[definition.key];
      if (found && existing && !existing.id) {
        listsPatch[definition.key] = {
          id: found.Id,
          title: found.Title,
          template: found.BaseTemplate,
          baseType: found.BaseType,
          ...(found.ItemEntityTypeFullName ? { itemEntityType: found.ItemEntityTypeFullName } : {}),
          serverRelativeUrl: found.RootFolder && found.RootFolder.ServerRelativeUrl || "",
        };
      }
    });
    patch.resources = {
      lists: listsPatch,
      ...(primary ? { items: { primaryItemId: primary.Id, primaryItemUniqueId: primary.UniqueId, ...(secondary ? { secondaryItemId: secondary.Id } : {}) } } : {}),
      ...(chosenList.RootFolder?.ServerRelativeUrl && !config.resources.folders.serverRelativePath ? { folders: { serverRelativePath: chosenList.RootFolder.ServerRelativeUrl } } : {}),
      ...(titleField ? { fields: { titleFieldId: titleField.Id } } : {}),
      ...(defaultView ? { views: { viewId: defaultView.Id, viewName: defaultView.Title } } : {}),
      ...(customContentTypeId ? { contentTypes: { customContentTypeId } } : {}),
    };
  }
  const principalsPatch: Record<string, unknown> = {};
  const userConfig = (candidate: DiscoveredUser) => ({ id: candidate.Id, upn: candidate.Email || candidate.LoginName.split("|").pop() || "", loginName: candidate.LoginName });
  const selectedSecondary = users.find((candidate) => String(candidate.Id) === secondaryUserId);
  if (user || selectedSecondary) principalsPatch.users = {
    ...(user ? { primary: userConfig(user) } : {}),
    ...(selectedSecondary ? { secondary: userConfig(selectedSecondary) } : {}),
  };
  const groupsPatch: Record<string, unknown> = {};
  const selectedGroup = groups.find((group) => String(group.Id) === selectedGroupId);
  if (selectedGroup) {
    groupsPatch.primaryGroupId = selectedGroup.Id;
    groupsPatch.primaryGroupName = selectedGroup.Title;
  }
  if (guessedGroups.owner) groupsPatch.ownerGroupId = guessedGroups.owner.Id;
  if (guessedGroups.member) groupsPatch.memberGroupId = guessedGroups.member.Id;
  if (guessedGroups.visitor) groupsPatch.visitorGroupId = guessedGroups.visitor.Id;
  if (selectedGroup) groupsPatch.groupsToEnumerate = [selectedGroup.Id];
  if (Object.keys(groupsPatch).length) principalsPatch.groups = groupsPatch;
  if (user || selectedSecondary) {
    principalsPatch.shareWith = [user, selectedSecondary].filter(Boolean).map((candidate) => (candidate as DiscoveredUser).Email || (candidate as DiscoveredUser).LoginName.split("|").pop() || "");
    principalsPatch.principalIds = [user, selectedSecondary].filter(Boolean).map((candidate) => (candidate as DiscoveredUser).Id).concat(
      [selectedGroup, guessedGroups.owner, guessedGroups.member, guessedGroups.visitor].filter(Boolean).map((group) => (group as DiscoveredGroup).Id)
    );
  } else if (selectedGroup) {
    principalsPatch.principalIds = [selectedGroup.Id];
  }
  if (Object.keys(principalsPatch).length) patch.principals = principalsPatch;
  const json = JSON.stringify(patch, null, 2);
  const copyPatch = async (): Promise<void> => {
    await navigator.clipboard.writeText(json);
    setCopied(true);
    setTimeout(() => setCopied(false), 2500);
  };

  return (
    <section className="researchApiPanel researchApiPanel--setup">
      <h2>Test config setup check</h2>
      <p>Some settings in <code>src/config/test-config.json</code> are blank. I checked the current SharePoint site with read-only API calls and assembled values you can copy into the matching config sections. Merge this snippet into the file; it is not a replacement for the whole file.</p>
      {loading && <Spinner label="Discovering site, lists, user and groups" />}
      {error && <MessageBar messageBarType={MessageBarType.warning}>{error}</MessageBar>}
      <ul>{missing.map((item) => <li key={item}>{item}</li>)}</ul>
      {lists.length > 0 && <>
        <Dropdown label="Choose the primary list for list and item API calls" options={listOptions} selectedKey={primaryListId} onChange={(_event, option) => setPrimaryListId(option ? String(option.key) : "")} />
        <Dropdown label="Optional lookup list" options={([{ key: "", text: "None" }] as IDropdownOption[]).concat(listOptions)} selectedKey={lookupListId} onChange={(_event, option) => setLookupListId(option ? String(option.key) : "")} />
      </>}
      {groups.length > 0 && <Dropdown label="Optional primary group" options={groups.map((group) => ({ key: String(group.Id), text: group.Title + " · " + group.Id }))} selectedKey={selectedGroupId} onChange={(_event, option) => setSelectedGroupId(option ? String(option.key) : "")} />}
      {users.length > 1 && <Dropdown label="Optional secondary user for two-user tests" options={users.filter((candidate) => !user || candidate.Id !== user.Id).map((candidate) => ({ key: String(candidate.Id), text: candidate.Title + " · " + (candidate.Email || candidate.LoginName) }))} selectedKey={secondaryUserId} onChange={(_event, option) => setSecondaryUserId(option ? String(option.key) : "")} />}
      {details && <div className="researchApiPanel__meta"><b>{details.list.Title}</b><span>{details.items.length} sample items, {details.views.length} views, {details.fields.length} visible fields, {details.contentTypes.length} content types discovered.</span></div>}
      {loadingDetails && <Spinner label="Discovering selected list items, views, fields and content types" />}
      {details && details.items.length > 0 && <Dropdown label="Sample primary item for config" options={details.items.map((item) => ({ key: String(item.Id), text: "#" + item.Id + " · " + (item.Title || "(untitled)") + " · " + item.UniqueId }))} selectedKey={sampleItemId} onChange={(_event, option) => setSampleItemId(option ? String(option.key) : "")} />}
      {details && details.contentTypes.length > 0 && <Dropdown label="Optional custom content type" options={([{ key: "", text: "None" }] as IDropdownOption[]).concat(details.contentTypes.map((item) => ({ key: item.StringId, text: item.Name + " · " + item.StringId })))} selectedKey={customContentTypeId} onChange={(_event, option) => setCustomContentTypeId(option ? String(option.key) : "")} />}
      {groups.length > 0 && <details><summary>Available SharePoint groups ({groups.length})</summary><pre className="researchApiPanel__output">{JSON.stringify(groups.map((group) => ({ id: group.Id, title: group.Title })), null, 2)}</pre></details>}
      {details && <details><summary>Discovered item, view, field and content type values</summary><pre className="researchApiPanel__output">{JSON.stringify({ items: details.items, views: details.views, titleField: titleField || null, contentTypes: details.contentTypes }, null, 2)}</pre></details>}
      {json !== "{}" && <>
        <Stack horizontal tokens={{ childrenGap: 8 }}><DefaultButton text={copied ? "Copied" : "Copy config JSON"} onClick={() => { void copyPatch(); }} /></Stack>
        <pre className="researchApiPanel__output">{json}</pre>
      </>}
      <MessageBar messageBarType={MessageBarType.info}>Paste the suggested values into test-config.json, review the selected list and user, then reload the web part. Unavailable values are left for you to choose; no config is changed automatically.</MessageBar>
    </section>
  );
}
