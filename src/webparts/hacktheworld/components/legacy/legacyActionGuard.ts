import { TestConfig } from "../../config/loader";

export interface LegacyActionReadiness {
  missing: string[];
  destructive: boolean;
  changesPermissions: boolean;
}

/** Checks the shared configuration before a legacy button can run its probes. */
export function checkLegacyAction(label: string, config: TestConfig): LegacyActionReadiness {
  const text = label.toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
  const missing: string[] = [];
  let siteHost = "";
  try {
    const site = new URL(config.environment.siteUrl);
    const tenant = new URL(config.environment.tenantUrl);
    siteHost = site.hostname.toLowerCase();
    if (site.protocol !== "https:" || tenant.hostname.toLowerCase() !== siteHost || siteHost.endsWith(".invalid")) {
      missing.push("environment.tenantUrl and environment.siteUrl must point to your authorized HTTPS tenant site");
    }
  } catch (_error) {
    missing.push("environment.tenantUrl and environment.siteUrl must be valid HTTPS URLs");
  }
  if (siteHost && config.safety.allowedHosts.indexOf(siteHost) < 0) {
    missing.push("safety.allowedHosts must include " + siteHost);
  }

  const siteOnly = /^(site|getprincipalids)$/i.test(text);
  const needsList = !siteOnly;
  if (needsList && (!config.resources.lists.primary.id || !config.resources.lists.primary.title)) {
    missing.push("resources.lists.primary.id and title (use Test config setup check)");
  }
  if (/library|file|folder/.test(text) && !config.resources.lists.documents.id) {
    missing.push("resources.lists.documents.id (use Test config setup check)");
  }
  if (/listcolumn/.test(text) && !config.resources.lists.lookup?.id) {
    missing.push("resources.lists.lookup.id (choose an optional lookup list in Test config setup check)");
  }
  if (/listcolumn/.test(text) && !config.resources.fields.titleFieldId) {
    missing.push("resources.fields.titleFieldId (use Test config setup check)");
  }
  if (/content.?type/.test(text) && !config.resources.contentTypes.customContentTypeId) {
    missing.push("resources.contentTypes.customContentTypeId");
  }
  if (/view|folder info/.test(text) && !config.resources.views.viewId) {
    missing.push("resources.views.viewId (use Test config setup check)");
  }
  if (/folder/.test(text) && !config.resources.folders.serverRelativePath) {
    missing.push("resources.folders.serverRelativePath");
  }
  if (/^site$/.test(text) && !config.principals.groups.groupsToEnumerate.length) {
    missing.push("principals.groups.groupsToEnumerate (choose a group in Test config setup check)");
  }

  const needsIdentity = /share|sharing|haringinformation|role|permission|user/i.test(text) &&
    !/getshareid|getprincipalids/i.test(text);
  if (needsIdentity && (!config.principals.users.primary.upn || !config.principals.users.primary.loginName)) {
    missing.push("principals.users.primary.upn and loginName (use Test config setup check)");
  }

  const changesPermissions = /share list|share object|sharelist|rootfolder share|assign role|role assignment|roleinheritance|update.*role|remove.*role|with users|sharing link|share item/i.test(text);
  const readOnly = /^(get item information|get library item information|get listitem version information|list item versions|get list item folder info|site|getprincipalids|getshareid)$/i.test(text);
  const destructive = !readOnly && /add|update|delete|remove|assign|create|share list|withpassword|expiry|role6|rootfolder share|listcrud|file|folder|attachment|comment/i.test(text);
  if (destructive && !config.safety.allowDestructive) {
    missing.push("safety.allowDestructive must be true for this write operation");
  }
  if (changesPermissions && !config.safety.allowPermissionChanges) {
    missing.push("safety.allowPermissionChanges must be true for permission or sharing changes");
  }
  if (changesPermissions) {
    if (!config.principals.users.primary.upn || !config.principals.users.primary.loginName) {
      missing.push("principals.users.primary.upn and loginName");
    }
    if (!config.principals.groups.shareWith.length) {
      missing.push("principals.groups.shareWith");
    }
    if (!config.principals.groups.primaryGroupId) {
      missing.push("principals.groups.primaryGroupId");
    }
  }
  return { missing: missing, destructive: destructive, changesPermissions: changesPermissions };
}
