# 04 - Endpoint inventory

Generated from `src/webparts/hacktheworld/research/registry.ts` on 2026-10-30.
87 endpoints across 19 modules.

Legend:

- **Digest** - the endpoint needs `X-RequestDigest` for a write.
- **Destructive** - the definition includes an operation that creates, modifies
  or deletes state. Blocked by `safety.allowDestructive=false` until enabled.
- Presence in this table means **implemented**, not executed or verified.

Regenerate with:

```powershell
npm run research:coverage
```

## Summary

| Module | Endpoints |
| --- | --- |
| lists | 29 |
| items | 16 |
| permissions | 5 |
| attachments | 4 |
| files | 4 |
| changes | 3 |
| content-types | 3 |
| fields | 3 |
| search | 3 |
| views | 3 |
| comments | 2 |
| folders | 2 |
| groups | 2 |
| sharing | 2 |
| users | 2 |
| batch | 1 |
| context | 1 |
| recycle-bin | 1 |
| sites | 1 |

Declared methods: GET on 45 endpoints, POST on 43, DELETE on 9, PATCH on 7,
MERGE on 2. HEAD, OPTIONS and PUT are generated as variations only.

## Full inventory

### attachments

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `items.attachments` | GET | sharepoint-rest | no | no |
| `items.attachment.add` | POST | sharepoint-rest | yes | yes |
| `items.attachment.get` | GET | sharepoint-rest | no | no |
| `items.attachment.delete` | DELETE | sharepoint-rest | yes | yes |

### batch

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `batch.post` | POST | sharepoint-rest | yes | yes |

### changes

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `lists.getChanges` | POST | sharepoint-rest | yes | no |
| `lists.getItemChangesSinceToken` | POST | sharepoint-rest | yes | no |
| `web.getChanges` | POST | sharepoint-rest | yes | no |

### comments

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `items.comments` | GET | sharepoint-rest | no | no |
| `items.comment.add` | POST | sharepoint-rest | yes | yes |

### content-types

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `contentTypes.collection` | GET | sharepoint-rest | no | no |
| `contentTypes.byId` | GET, PATCH, DELETE | sharepoint-rest | yes | yes |
| `contentTypes.addAvailable` | POST | sharepoint-rest | yes | yes |

### context

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `rest.contextinfo` | POST | sharepoint-rest | no | no |

### fields

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `fields.collection` | GET, POST | sharepoint-rest | yes | yes |
| `fields.getByTitle` | GET, PATCH, DELETE | sharepoint-rest | yes | yes |
| `fields.createFieldAsXml` | POST | sharepoint-rest | yes | yes |

### files

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `files.byPath` | GET, DELETE | sharepoint-rest | yes | yes |
| `files.addUsingPath` | POST | sharepoint-rest | yes | yes |
| `files.versions` | GET | sharepoint-rest | no | no |
| `moveCopy.moveFileByPath` | POST | sharepoint-rest | yes | yes |

### folders

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `folders.byPath` | GET, DELETE | sharepoint-rest | yes | yes |
| `folders.add` | POST | sharepoint-rest | yes | yes |

### groups

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `siteGroups.collection` | GET, POST | sharepoint-rest | yes | yes |
| `siteGroups.users` | GET, POST | sharepoint-rest | yes | yes |

### items

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `items.collection.get` | GET | sharepoint-rest | no | no |
| `items.collection.add` | POST | sharepoint-rest | yes | yes |
| `items.getById` | GET | sharepoint-rest | no | no |
| `items.getByStringId` | GET | sharepoint-rest | no | no |
| `items.update` | PATCH, MERGE | sharepoint-rest | yes | yes |
| `items.delete` | DELETE | sharepoint-rest | yes | yes |
| `items.recycle` | POST | sharepoint-rest | yes | yes |
| `items.validateUpdateListItem` | POST | sharepoint-rest | yes | yes |
| `items.fieldValuesAsText` | GET | sharepoint-rest | no | no |
| `items.fieldValuesAsHTML` | GET | sharepoint-rest | no | no |
| `items.versions` | GET | sharepoint-rest | no | no |
| `items.version.get` | GET | sharepoint-rest | no | no |
| `v2.site.list.items` | GET | sharepoint-v2 | no | no |
| `v2.site.list.items.activities` | GET | sharepoint-v2 | no | no |
| `graph.sites.lists.items` | GET, POST | microsoft-graph | no | yes |
| `graph.sites.lists.items.get` | PATCH, DELETE | microsoft-graph | no | yes |

### lists

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `lists.collection.get` | GET | sharepoint-rest | no | no |
| `lists.collection.add` | POST | sharepoint-rest | yes | yes |
| `lists.getByTitle` | GET | sharepoint-rest | no | no |
| `lists.getById` | GET | sharepoint-rest | no | no |
| `lists.update` | PATCH, MERGE | sharepoint-rest | yes | yes |
| `lists.delete` | DELETE | sharepoint-rest | yes | yes |
| `lists.recycle` | POST | sharepoint-rest | yes | yes |
| `lists.renderListDataAsStream` | POST | sharepoint-rest | yes | no |
| `lists.renderListData` | POST | sharepoint-rest | yes | no |
| `lists.getItemsByCamlQuery` | POST | sharepoint-rest | yes | no |
| `lists.addValidateUpdateItemUsingPath` | POST | sharepoint-rest | yes | yes |
| `lists.effectiveBasePermissions` | GET | sharepoint-rest | no | no |
| `lists.eventReceivers` | GET | sharepoint-rest | no | no |
| `lists.relatedFields` | GET, POST | sharepoint-rest | yes | no |
| `lists.irmSettings` | GET, PATCH | sharepoint-rest | yes | no |
| `lists.renderListFormData` | POST | sharepoint-rest | yes | no |
| `lists.reserveListItemId` | POST | sharepoint-rest | yes | no |
| `lists.forms` | GET | sharepoint-rest | no | no |
| `lists.subscriptions` | GET | sharepoint-rest | no | no |
| `lists.subscription.add` | POST | sharepoint-rest | yes | yes |
| `lists.userCustomActions` | GET | sharepoint-rest | no | no |
| `lists.defaultView` | GET | sharepoint-rest | no | no |
| `lists.roleAssignments` | GET | sharepoint-rest | no | no |
| `lists.breakRoleInheritance` | POST | sharepoint-rest | yes | yes |
| `lists.resetRoleInheritance` | POST | sharepoint-rest | yes | yes |
| `lists.ensureSiteAssetsLibrary` | POST | sharepoint-rest | yes | yes |
| `lists.ensureSitePagesLibrary` | POST | sharepoint-rest | yes | yes |
| `v2.site.lists` | GET | sharepoint-v2 | no | no |
| `graph.sites.lists` | GET | microsoft-graph | no | no |

### permissions

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `items.sharingInformation` | POST | sharepoint-rest | yes | no |
| `items.roleAssignments` | GET | sharepoint-rest | no | no |
| `items.effectiveBasePermissions` | GET | sharepoint-rest | no | no |
| `roleAssignments.add` | POST | sharepoint-rest | yes | yes |
| `roleDefinitions.collection` | GET | sharepoint-rest | no | no |

### recycle-bin

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `web.recycleBin` | GET | sharepoint-rest | no | no |

### search

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `search.query` | GET, POST | sharepoint-rest | no | no |
| `search.postquery` | POST | sharepoint-rest | yes | no |
| `graph.search.query` | POST | microsoft-graph | no | no |

### sharing

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `items.shareLink` | POST | sharepoint-rest | yes | yes |
| `web.shareObject` | POST | sharepoint-rest | yes | yes |

### sites

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `graph.sites.get` | GET | microsoft-graph | no | no |

### users

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `siteUsers.collection` | GET | sharepoint-rest | no | no |
| `siteUsers.ensureUser` | POST | sharepoint-rest | yes | yes |

### views

| Endpoint | Methods | Family | Digest | Destructive |
| --- | --- | --- | --- | --- |
| `views.collection` | GET, POST | sharepoint-rest | yes | yes |
| `views.byId` | GET, PATCH, DELETE | sharepoint-rest | yes | yes |
| `views.renderAsHtml` | POST | sharepoint-rest | yes | no |

## Known gaps

- Taxonomy, workflow and tenant-level webhook endpoints are not registered.
- Microsoft Graph definitions exist but the transport is SharePoint
  cookie+digest only; Graph needs a separate bearer-token transport.
- `HEAD`, `OPTIONS`, `PUT` are probe-only, not declared.
- No endpoint is executed or verified. See `06-test-evidence.md`.
