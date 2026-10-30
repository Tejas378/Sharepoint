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

## 4. Lists milestone

The Lists module is the first milestone. It now has 29 definitions.

### Pre-existing / core (13)

`lists.collection.get`, `lists.collection.add`, `lists.getByTitle`,
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

No duplicate endpoint was added. `registry.test.ts` asserts at least 29 list
endpoints and unique ids.

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
- Microsoft Graph list endpoints are in the registry as a separate family but
  the transport is not implemented. They require a bearer token, not the
  SharePoint cookie + digest model.
- Taxonomy, workflows and tenant-level webhooks were not added because no
  installed PnPjs source in this project gives a stable URL for them. Adding
  guessed URLs would violate the registry rules.

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
