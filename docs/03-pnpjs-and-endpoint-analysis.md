# 03 - PnPjs 4.9.0 and endpoint sourcing

This document records how endpoint definitions were derived. It is the evidence
trail for the registry; it is not documentation copied from the internet.

## 1. Installed version

```text
node_modules/@pnp/sp/package.json  ->  "version": "4.9.0"
```

The dependency range in `package.json` is `^4.9.0`; no upgrade was performed.
Everything below was read from the installed source, not from current online
docs, because online docs may describe a newer version.

## 2. Files inspected

| File | Used for |
| --- | --- |
| `@pnp/sp/lists/types.js` | List CRUD, `getById`, `getByTitle`, `ensureSiteAssetsLibrary`, `ensureSitePagesLibrary`, `effectiveBasePermissions`, `eventReceivers`, `relatedFields`, IRM settings, update/merge, recycle, CAML query, render data, `renderListFormData`, `reserveListItemId`. |
| `@pnp/sp/security/funcs.js` | `breakRoleInheritance`, `resetRoleInheritance`, effective permissions, role assignment helpers. |
| `@pnp/sp/security/types.js` | Role assignment/inheritance method signatures. |
| `@pnp/sp/user-custom-actions/list.js` | List-scoped user custom actions. |
| `@pnp/sp/subscriptions/types.js` | List subscriptions and webhook add/remove. |
| `@pnp/sp/forms/types.js` | List forms collection. |
| `@pnp/sp/items/types.js` | Item CRUD, validate-update, attachments, comments, sharing. |
| `@pnp/sp/fields/types.js`, `content-types/types.js`, `views/types.js` | Field/content-type/view operations. |
| `@pnp/sp/files/*`, `folders/*` | File/folder operations and path handling. |
| `@pnp/sp/batching` | Batch endpoint shape. |
| `@pnp/sp/sites`, `webs`, `site-users`, `site-groups`, `security`, `sharing`, `recycle-bin`, `search`, `comments`, `attachments`, `approvals`, `taxonomy`, `webparts`, `navigation`, `features`, `profiles`, `regional-settings`, `related-items`, `site-designs`, `site-scripts`, `hubsites`, `appcatalog`, `clientside-pages`, `presets`, `social`, `sputilities`, `publishing-sitepageservice`, `groupsitemanager`, `context-info`, `column-defaults`, `utils` | Installed module inventory cross-check; only URL-backed operations are registered and sourced. |

The active registry currently has 316 unique endpoint IDs. In particular, it
includes 12 site, 8 web, 42 file, 17 folder, 33 list, 20 item, 23 field, 21
content-type, 33 permission and 14 view operations. Site/web, principal, role assignment, role definition,
search and batch routes were cross-checked against their installed PnPjs
methods; search query keys are included in both GET and POST forms. This count
is the current implemented registry, not a claim
that SharePoint has a finite or tenant-independent set of every possible route.
PnPjs modules can expose overloads, feature-dependent routes and non-REST
client APIs; only routes found in installed source, Microsoft documentation or
the project's legacy calls are included. Legacy document-library calls were
searched and mapped where they represent distinct URLs. Repeated calls to the
same URL are represented once; method and payload variations are handled by
the test controls.

Document libraries are SharePoint lists with `BaseTemplate` 101 rather than a
separate REST resource type. The Libraries button uses the shared list
collection endpoint with a library filter; list configuration remains under
Lists, while document operations are under Files, Folders, Items, Fields,
Content Types, Views and Permissions. Batch read-only and write changesets
share `/_api/$batch` but retain separate UI profiles so writes stay gated.

The UI offers a module button, endpoint selector, and payload-family/case
selectors. A selected payload modifies one request preview at a time. It does
not automatically execute a payload set. Destructive endpoints remain gated
by `safety.allowDestructive` and the request still requires tenant config and
the configured SharePoint origin.

## 3. Rules applied

1. A method is only `documented` if its URL appears in the installed source or
   Microsoft documentation.
2. A type definition without a URL is `pnpjs-typedef`, not `documented`.
3. Anything observed only in the legacy project code is `project-code` and is
   treated as a hypothesis until confirmed.
4. The registry never lists a method as supported unless the endpoint definition
   declares it. All other verbs are generated as `controlled-probe` or
   `hypothesis` by `variations.ts`.
5. URL templates mirror the PnPjs string exactly, including:
   - `guid'...'` for GUID path parameters;
   - `getByTitle('...')` with path encoding;
   - `breakroleinheritance(copyroleassignments=false, clearsubscopes=false)`;
   - `renderlistformdata(itemid=..., formid='...', mode='...')`.

## 4. Lists and document-library coverage

Lists, document libraries, files, folders, fields, content types, views,
sharing and item metadata are covered as separate UI modules. The list module
currently has 33 definitions. Prior doclib-specific calls (sharing info,
sharing links, document-sharing updates, stream rendering, path-based file and
folder operations, field/content-type/view mutations and upload/version
operations) were compared against the registry. See the generated inventory
for the canonical endpoint IDs and current counts.

### Pre-existing / core (13)

`libraries.collection` (the shared `/_api/web/lists` collection, filtered to
document libraries by default), `lists.collection.add`, `lists.getByTitle`,
`lists.getById`, `lists.update`, `lists.delete`, `lists.recycle`,
`lists.renderListDataAsStream`, `lists.renderListData`,
`lists.getItemsByCamlQuery`, `lists.addValidateUpdateItemUsingPath`, plus the
v2/Graph list variants tracked separately.

### Added from PnPjs 4.9.0 (16)

| Endpoint | Method | Evidence |
| --- | --- | --- |
| `lists.effectiveBasePermissions` | GET | `security/funcs.js`, `lists/types.js` |
| `lists.eventReceivers` | GET | `lists/types.js` |
| `lists.relatedFields` | GET, POST | `lists/types.js` (`getRelatedFields`) |
| `lists.irmSettings` | GET, PATCH | `lists/types.js` (`InformationRightsManagementSettings`) |
| `lists.renderListFormData` | POST | `lists/types.js` |
| `lists.reserveListItemId` | POST | `lists/types.js` |
| `lists.forms` | GET | `forms/types.js` |
| `lists.subscriptions` | GET | `subscriptions/types.js` |
| `lists.subscription.add` | POST | `subscriptions/types.js` |
| `lists.userCustomActions` | GET | `user-custom-actions/list.js` |
| `lists.defaultView` | GET | documented `SP.List` navigation property |
| `lists.roleAssignments` | GET | `security/types.js` |
| `lists.breakRoleInheritance` | POST | `security/funcs.js` |
| `lists.resetRoleInheritance` | POST | `security/funcs.js` |
| `lists.ensureSiteAssetsLibrary` | POST | `lists/types.js` |
| `lists.ensureSitePagesLibrary` | POST | `lists/types.js` |

Endpoint IDs are unique at module load; repeated legacy invocations of the
same URL are consolidated. List and document-library sharing routes retain
their distinct request payloads and safety classification.

## 5. Methods

Declared methods are limited to what the source shows:

- GET for reads and navigation properties;
- POST for actions (`renderlistdata`, `recycle`, `ensuresiteassetslibrary`,
  `breakroleinheritance`, ...);
- PATCH and MERGE for updates (`MERGE` is sent as POST + `X-HTTP-Method: MERGE`
  by default);
- DELETE for deletable resources.

HEAD, OPTIONS and PUT are **not** added as declared methods. They are generated
as probes so a 405 is expected and a 2xx is interesting.

## 6. Parameters

Each parameter records `name`, `location`, `type`, `required`, `example`,
`source` and `verification`. Examples are synthetic (`00000000-...`, `2`,
`Id,Title`) so they are safe to commit.

Known limitations:

- Not every PnPjs overload maps to a distinct endpoint; some are parameter
  combinations of the same URL.
- PnPjs sometimes concatenates a path client-side (`getById`, `getByTitle`);
  the registry records the resulting URL, not the TypeScript method name.
- Microsoft Graph endpoints are in a separate family with a bearer-token
  transport; callers provide the approved token acquisition flow.
- Microsoft Graph term store reads use the documented site-scoped routes.
  Workflow services are described as a CSOM client API, so no REST route is
  claimed. SharePoint Online webhook subscription operations are list scoped;
  no tenant-level subscription route is documented.

Primary references:

- [Determine SharePoint REST service endpoint URIs](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/determine-sharepoint-rest-service-endpoint-uris)
- [Work with folders and files using REST](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-folders-and-files-with-rest)
- [Make batch requests with SharePoint REST](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/make-batch-requests-with-the-rest-apis)
- [PnPjs source repository](https://github.com/pnp/pnpjs)
- [Microsoft Graph term store](https://learn.microsoft.com/en-us/graph/api/resources/termstore-store?view=graph-rest-1.0)
- [Microsoft Graph term store groups](https://learn.microsoft.com/en-us/graph/api/termstore-list-groups?view=graph-rest-1.0)
- [Microsoft Graph list sets](https://learn.microsoft.com/en-us/graph/api/termstore-group-list-sets?view=graph-rest-1.0)
- [Microsoft Graph list children](https://learn.microsoft.com/en-us/graph/api/termstore-term-list-children?view=graph-rest-1.0)
- [SharePoint list webhook operations](https://learn.microsoft.com/en-us/sharepoint/dev/apis/webhooks/lists/overview-sharepoint-list-webhooks)
- [SharePoint WorkflowServicesManager](https://learn.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.client.workflowservices.workflowservicesmanager?view=sharepoint-csom)

## 7. How to add the next endpoint

1. Find the URL in installed source or Microsoft docs.
2. Add an `endpoint({...})` entry with `sources` and per-parameter
   `source`/`verification`.
3. Give every path parameter a safe synthetic `example`; otherwise the coverage
   report will mark the baseline as not generated.
4. Set `destructive`, `requiresDigest`, `conditional` and `payloadFamilies`
   honestly.
5. Run:

```powershell
npm run research:test
npm run research:coverage
```

6. Add a test only if the endpoint has non-obvious URL/encoding behaviour.
7. Do not claim execution. A registry entry is `implemented`, not `verified`.
