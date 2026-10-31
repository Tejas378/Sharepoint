# Endpoint coverage report

Generated: 2026-10-31T18:53:12.598Z

| Metric | Value |
| --- | --- |
| Registered endpoints | 530 |
| Implemented | 530 |
| Baseline generated | 448 |
| Executed | 0 |
| Verified | 0 |

## Modules

| Module | Endpoints | Methods | Payload families |
| --- | --- | --- | --- |
| app-catalog | 20 | GET, POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| approvals | 5 | GET, POST | idor, input-validation, authz, business-logic, mass-assignment |
| attachments | 7 | GET, POST, DELETE | idor, authz, input-validation, path-traversal, xss, ssrf, method-confusion, business-logic, mass-assignment |
| batch | 3 | POST | method-confusion, authz, mass-assignment, input-validation, business-logic, xss |
| changes | 3 | POST | input-validation, business-logic, authz |
| client-pages | 12 | GET, POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| comments | 8 | GET, POST, DELETE | idor, xss, authz, sqli, mass-assignment, input-validation, business-logic, method-confusion |
| configuration | 8 | GET, POST | idor, authz, input-validation, odata-injection, business-logic, mass-assignment, method-confusion |
| content-types | 21 | GET, POST, PATCH, MERGE, DELETE | odata-injection, authz, mass-assignment, xss, input-validation, business-logic, idor, method-confusion |
| context | 1 | POST | input-validation |
| custom-actions | 12 | GET, POST, MERGE, DELETE | xss, authz, input-validation, business-logic, idor, mass-assignment |
| features | 4 | GET, POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| fields | 23 | GET, POST, PATCH, MERGE, DELETE | mass-assignment, xss, input-validation, business-logic, authz, idor, method-confusion, odata-injection |
| files | 43 | GET, DELETE, POST, MERGE | idor, path-traversal, authz, method-confusion, input-validation, ssrf, xss, mass-assignment, odata-injection, business-logic |
| folders | 17 | GET, DELETE, POST, MERGE | idor, path-traversal, authz, method-confusion, input-validation, odata-injection, business-logic, xss, mass-assignment |
| group-site-manager | 25 | GET, POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| groups | 17 | GET, POST, MERGE, DELETE | mass-assignment, xss, input-validation, business-logic, authz, idor, odata-injection, method-confusion |
| hubsites | 2 | GET | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| items | 27 | GET, POST, PATCH, MERGE, DELETE | idor, odata-injection, input-validation, authz, mass-assignment, xss, sqli, method-confusion, business-logic |
| libraries | 1 | GET | odata-injection, authz, input-validation |
| lists | 40 | GET, POST, PATCH, MERGE, DELETE | odata-injection, authz, input-validation, mass-assignment, xss, business-logic, method-confusion, idor |
| navigation | 16 | GET, POST, MERGE, DELETE | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| permissions | 33 | POST, GET, MERGE, DELETE | idor, authz, business-logic, method-confusion, input-validation, odata-injection, mass-assignment |
| profiles | 24 | GET, POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| publishing-site-page-service | 1 | GET | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| recycle-bin | 8 | GET, POST | odata-injection, idor, authz, path-traversal, business-logic |
| regional-settings | 24 | GET, POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| related-items | 6 | POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| search | 4 | GET, POST | odata-injection, xss, sqli, input-validation, authz |
| sharing | 11 | POST | idor, authz, business-logic, input-validation, ssrf, mass-assignment, xss |
| site-designs | 14 | POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| site-scripts | 8 | POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| sites | 12 | GET, POST, PATCH, MERGE | idor, authz, input-validation, ssrf, business-logic, mass-assignment, method-confusion |
| social | 9 | POST, GET | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| taxonomy | 4 | GET | authz, odata-injection, idor |
| users | 11 | GET, POST, MERGE, DELETE | odata-injection, idor, authz, input-validation, xss, business-logic, mass-assignment, method-confusion |
| utilities | 6 | POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| views | 14 | GET, POST, PATCH, MERGE, DELETE | odata-injection, xss, mass-assignment, input-validation, authz, idor, method-confusion, business-logic |
| webhooks | 5 | GET, POST, PATCH, DELETE | authz, input-validation, ssrf, business-logic, xss, idor |
| webparts | 12 | GET, POST | idor, authz, input-validation, business-logic, mass-assignment, method-confusion |
| webs | 9 | GET, POST, PATCH, MERGE, DELETE | idor, odata-injection, authz, mass-assignment, xss, input-validation, business-logic, method-confusion |

## Declared methods

- DELETE: 23 endpoint(s)
- GET: 189 endpoint(s)
- MERGE: 19 endpoint(s)
- PATCH: 10 endpoint(s)
- POST: 308 endpoint(s)

## Gaps

- web.createSubweb: destructive; blocked by safety.allowDestructive=false.
- web.updateConfiguration: destructive; blocked by safety.allowDestructive=false.
- lists.collection.add: destructive; blocked by safety.allowDestructive=false.
- lists.getById: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.update: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.update: destructive; blocked by safety.allowDestructive=false.
- lists.delete: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.delete: destructive; blocked by safety.allowDestructive=false.
- lists.recycle: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.recycle: destructive; blocked by safety.allowDestructive=false.
- lists.getChanges: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.getItemChangesSinceToken: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.renderListData: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.getItemsByCamlQuery: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.addValidateUpdateItemUsingPath: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.addValidateUpdateItemUsingPath: destructive; blocked by safety.allowDestructive=false.
- lists.effectiveBasePermissions: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.eventReceivers: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.eventReceivers.add: destructive; blocked by safety.allowDestructive=false.
- lists.relatedFields: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.irmSettings: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.renderListFormData: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.reserveListItemId: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.forms: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.subscriptions: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.subscription.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.subscription.add: destructive; blocked by safety.allowDestructive=false.
- lists.subscription.byId: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.subscription.update: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.subscription.update: destructive; blocked by safety.allowDestructive=false.
- lists.subscription.delete: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.subscription.delete: destructive; blocked by safety.allowDestructive=false.
- lists.defaultView: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.roleAssignments: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.breakRoleInheritance: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.breakRoleInheritance: destructive; blocked by safety.allowDestructive=false.
- lists.resetRoleInheritance: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.resetRoleInheritance: destructive; blocked by safety.allowDestructive=false.
- lists.ensureSiteAssetsLibrary: destructive; blocked by safety.allowDestructive=false.
- lists.ensureSitePagesLibrary: destructive; blocked by safety.allowDestructive=false.
- lists.validateUpdateFetchItem: destructive; blocked by safety.allowDestructive=false.
- lists.bulkValidateUpdateItems: destructive; blocked by safety.allowDestructive=false.
- items.collection.get: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.shareLink: destructive; blocked by safety.allowDestructive=false.
- lists.unshareLink: destructive; blocked by safety.allowDestructive=false.
- lists.updateDocumentSharingInfo: destructive; blocked by safety.allowDestructive=false.
- lists.roleAssignments.add: destructive; blocked by safety.allowDestructive=false.
- lists.roleAssignments.remove: destructive; blocked by safety.allowDestructive=false.
- items.collection.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.collection.add: destructive; blocked by safety.allowDestructive=false.
- items.getById: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.getByStringId: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.update: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.update: destructive; blocked by safety.allowDestructive=false.
- items.delete: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.delete: destructive; blocked by safety.allowDestructive=false.
- items.recycle: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.recycle: destructive; blocked by safety.allowDestructive=false.
- items.validateUpdateListItem: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.validateUpdateListItem: destructive; blocked by safety.allowDestructive=false.
- items.fieldValuesAsText: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.fieldValuesAsHTML: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.versions: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.version.get: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.attachments: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.version.delete: destructive; blocked by safety.allowDestructive=false.
- items.attachment.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.attachment.add: destructive; blocked by safety.allowDestructive=false.
- items.attachment.get: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.attachment.delete: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.attachment.delete: destructive; blocked by safety.allowDestructive=false.
- items.comments: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.comment.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.comment.add: destructive; blocked by safety.allowDestructive=false.
- items.comments.clear: destructive; blocked by safety.allowDestructive=false.
- items.comment.replies: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.comment.reply.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.comment.reply.add: destructive; blocked by safety.allowDestructive=false.
- items.comment.like: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.comment.like: destructive; blocked by safety.allowDestructive=false.
- items.comment.unlike: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.comment.unlike: destructive; blocked by safety.allowDestructive=false.
- items.comment.delete: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.comment.delete: destructive; blocked by safety.allowDestructive=false.
- items.sharingInformation: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.shareLink: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.shareLink: destructive; blocked by safety.allowDestructive=false.
- items.shareObject: destructive; blocked by safety.allowDestructive=false.
- items.unshareObject: destructive; blocked by safety.allowDestructive=false.
- items.deleteLinkByKind: destructive; blocked by safety.allowDestructive=false.
- items.unshareLink: destructive; blocked by safety.allowDestructive=false.
- items.roleAssignments: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.roleAssignments.add: destructive; blocked by safety.allowDestructive=false.
- items.roleAssignments.remove: destructive; blocked by safety.allowDestructive=false.
- items.effectiveBasePermissions: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.collection: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.add: destructive; blocked by safety.allowDestructive=false.
- fields.getByTitle: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.updateByTitle: destructive; blocked by safety.allowDestructive=false.
- fields.deleteByTitle: destructive; blocked by safety.allowDestructive=false.
- fields.addField: destructive; blocked by safety.allowDestructive=false.
- fields.createFieldAsXml: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.createFieldAsXml: destructive; blocked by safety.allowDestructive=false.
- fields.addDependentLookupField: destructive; blocked by safety.allowDestructive=false.
- fields.setShowInDisplayForm: destructive; blocked by safety.allowDestructive=false.
- fields.setShowInEditForm: destructive; blocked by safety.allowDestructive=false.
- fields.setShowInNewForm: destructive; blocked by safety.allowDestructive=false.
- contentTypes.collection: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.add: destructive; blocked by safety.allowDestructive=false.
- contentTypes.byId: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.updateById: destructive; blocked by safety.allowDestructive=false.
- contentTypes.deleteById: destructive; blocked by safety.allowDestructive=false.
- contentTypes.addAvailable: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.addAvailable: destructive; blocked by safety.allowDestructive=false.
- views.collection: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.add: destructive; blocked by safety.allowDestructive=false.
- views.byId: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.updateById: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.updateById: destructive; blocked by safety.allowDestructive=false.
- views.deleteById: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.deleteById: destructive; blocked by safety.allowDestructive=false.
- views.renderAsHtml: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.setViewXml: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.setViewXml: destructive; blocked by safety.allowDestructive=false.
- views.fields.collection: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.fields.schemaXml: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.fields.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.fields.add: destructive; blocked by safety.allowDestructive=false.
- views.fields.remove: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.fields.remove: destructive; blocked by safety.allowDestructive=false.
- views.fields.move: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.fields.move: destructive; blocked by safety.allowDestructive=false.
- views.fields.removeAll: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.fields.removeAll: destructive; blocked by safety.allowDestructive=false.
- views.htmlSchemaXml: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- files.deleteByPath: destructive; blocked by safety.allowDestructive=false.
- files.addUsingPath: destructive; blocked by safety.allowDestructive=false.
- files.addInFolder: destructive; blocked by safety.allowDestructive=false.
- files.checkOut: destructive; blocked by safety.allowDestructive=false.
- files.undoCheckOut: destructive; blocked by safety.allowDestructive=false.
- files.checkIn: destructive; blocked by safety.allowDestructive=false.
- files.publish: destructive; blocked by safety.allowDestructive=false.
- files.unpublish: destructive; blocked by safety.allowDestructive=false.
- files.approve: destructive; blocked by safety.allowDestructive=false.
- files.deny: destructive; blocked by safety.allowDestructive=false.
- files.copyTo: destructive; blocked by safety.allowDestructive=false.
- files.moveTo: destructive; blocked by safety.allowDestructive=false.
- files.recycle: destructive; blocked by safety.allowDestructive=false.
- files.deleteWithParameters: destructive; blocked by safety.allowDestructive=false.
- files.versions.deleteAll: destructive; blocked by safety.allowDestructive=false.
- files.versions.deleteById: destructive; blocked by safety.allowDestructive=false.
- files.versions.recycleById: destructive; blocked by safety.allowDestructive=false.
- files.versions.deleteByLabel: destructive; blocked by safety.allowDestructive=false.
- files.versions.recycleByLabel: destructive; blocked by safety.allowDestructive=false.
- files.versions.restoreByLabel: destructive; blocked by safety.allowDestructive=false.
- files.cancelUpload: destructive; blocked by safety.allowDestructive=false.
- files.startUpload: destructive; blocked by safety.allowDestructive=false.
- files.continueUpload: destructive; blocked by safety.allowDestructive=false.
- files.finishUpload: destructive; blocked by safety.allowDestructive=false.
- folders.deleteByPath: destructive; blocked by safety.allowDestructive=false.
- folders.addUsingPath: destructive; blocked by safety.allowDestructive=false.
- folders.addSubfolderUsingPath: destructive; blocked by safety.allowDestructive=false.
- folders.recycle: destructive; blocked by safety.allowDestructive=false.
- folders.add: destructive; blocked by safety.allowDestructive=false.
- moveCopy.moveFileByPath: destructive; blocked by safety.allowDestructive=false.
- moveCopy.copyFileByPath: destructive; blocked by safety.allowDestructive=false.
- moveCopy.moveFolderByPath: destructive; blocked by safety.allowDestructive=false.
- moveCopy.copyFolderByPath: destructive; blocked by safety.allowDestructive=false.
- siteUsers.ensureUser: destructive; blocked by safety.allowDestructive=false.
- siteUsers.removeById: destructive; blocked by safety.allowDestructive=false.
- siteGroups.add: destructive; blocked by safety.allowDestructive=false.
- siteGroups.removeById: destructive; blocked by safety.allowDestructive=false.
- siteGroups.setUserAsOwner: destructive; blocked by safety.allowDestructive=false.
- siteGroups.users.add: destructive; blocked by safety.allowDestructive=false.
- siteGroups.users.removeById: destructive; blocked by safety.allowDestructive=false.
- siteGroups.users.removeByLoginName: destructive; blocked by safety.allowDestructive=false.
- roleAssignments.add: destructive; blocked by safety.allowDestructive=false.
- roleAssignments.remove: destructive; blocked by safety.allowDestructive=false.
- web.shareObject: destructive; blocked by safety.allowDestructive=false.
- batch.post: destructive; blocked by safety.allowDestructive=false.
- batch.changeSet: destructive; blocked by safety.allowDestructive=false.
- recycleBin.item.restore: destructive; blocked by safety.allowDestructive=false.
- recycleBin.item.delete: destructive; blocked by safety.allowDestructive=false.
- recycleBin.item.moveToSecondStage: destructive; blocked by safety.allowDestructive=false.
- recycleBin.deleteAll: destructive; blocked by safety.allowDestructive=false.
- recycleBin.deleteAllSecondStageItems: destructive; blocked by safety.allowDestructive=false.
- recycleBin.moveAllToSecondStage: destructive; blocked by safety.allowDestructive=false.
- recycleBin.restoreAll: destructive; blocked by safety.allowDestructive=false.
- approvals.createItemRequest: destructive; blocked by safety.allowDestructive=false.
- approvals.cancelItemRequest: destructive; blocked by safety.allowDestructive=false.
- approvals.approveItemRequest: destructive; blocked by safety.allowDestructive=false.
- approvals.rejectItemRequest: destructive; blocked by safety.allowDestructive=false.
- v2.site.lists: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- v2.site.list.items: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- v2.site.list.items.activities: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- v2.favorites.followedListItems.add: destructive; blocked by safety.allowDestructive=false.
- graph.termStore.get: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- graph.termStore.groups: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- graph.termStore.sets: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- graph.termStore.children: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- graph.sites.get: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- graph.sites.lists: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- graph.sites.lists.items: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- graph.sites.lists.items: destructive; blocked by safety.allowDestructive=false.
- graph.sites.lists.items.get: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- graph.sites.lists.items.get: destructive; blocked by safety.allowDestructive=false.
- sites.createCommunicationSite: destructive; blocked by safety.allowDestructive=false.
- sites.createGroupSite: destructive; blocked by safety.allowDestructive=false.
- sites.deleteSiteCollection: destructive; blocked by safety.allowDestructive=false.
- sites.update: destructive; blocked by safety.allowDestructive=false.
- sites.setLogo: destructive; blocked by safety.allowDestructive=false.
- web.applyTheme: destructive; blocked by safety.allowDestructive=false.
- web.applyWebTemplate: destructive; blocked by safety.allowDestructive=false.
- web.applyTemplate: destructive; blocked by safety.allowDestructive=false.
- web.setStorageEntity: destructive; blocked by safety.allowDestructive=false.
- web.removeStorageEntity: destructive; blocked by safety.allowDestructive=false.
- web.delete: destructive; blocked by safety.allowDestructive=false.
- files.addTemplateFile: destructive; blocked by safety.allowDestructive=false.
- files.setContent: destructive; blocked by safety.allowDestructive=false.
- folders.update: destructive; blocked by safety.allowDestructive=false.
- folders.deleteWithParameters: destructive; blocked by safety.allowDestructive=false.
- siteUsers.add: destructive; blocked by safety.allowDestructive=false.
- siteUsers.update: destructive; blocked by safety.allowDestructive=false.
- siteUsers.removeByLoginName: destructive; blocked by safety.allowDestructive=false.
- siteUsers.deleteById: destructive; blocked by safety.allowDestructive=false.
- siteGroups.removeByLoginName: destructive; blocked by safety.allowDestructive=false.
- siteGroups.update: destructive; blocked by safety.allowDestructive=false.
- siteGroups.deleteById: destructive; blocked by safety.allowDestructive=false.
- permissions.web.breakInheritance: destructive; blocked by safety.allowDestructive=false.
- permissions.web.resetInheritance: destructive; blocked by safety.allowDestructive=false.
- permissions.item.breakInheritance: destructive; blocked by safety.allowDestructive=false.
- permissions.item.resetInheritance: destructive; blocked by safety.allowDestructive=false.
- web.fields.add: destructive; blocked by safety.allowDestructive=false.
- web.fields.addField: destructive; blocked by safety.allowDestructive=false.
- web.fields.createFieldAsXml: destructive; blocked by safety.allowDestructive=false.
- web.fields.addDependentLookup: destructive; blocked by safety.allowDestructive=false.
- web.fields.update: destructive; blocked by safety.allowDestructive=false.
- web.fields.delete: destructive; blocked by safety.allowDestructive=false.
- web.contentTypes.add: destructive; blocked by safety.allowDestructive=false.
- web.contentTypes.update: destructive; blocked by safety.allowDestructive=false.
- web.contentTypes.delete: destructive; blocked by safety.allowDestructive=false.
- web.contentTypes.addAvailable: destructive; blocked by safety.allowDestructive=false.
- roleDefinitions.add: destructive; blocked by safety.allowDestructive=false.
- roleDefinitions.update: destructive; blocked by safety.allowDestructive=false.
- roleDefinitions.delete: destructive; blocked by safety.allowDestructive=false.
- navigation.addQuicklaunch: destructive; blocked by safety.allowDestructive=false.
- navigation.moveAfter: destructive; blocked by safety.allowDestructive=false.
- navigation.quicklaunch.node.update: destructive; blocked by safety.allowDestructive=false.
- navigation.quicklaunch.node.delete: destructive; blocked by safety.allowDestructive=false.
- navigation.topNavigationBar.add: destructive; blocked by safety.allowDestructive=false.
- navigation.topNavigationBar.moveAfter: destructive; blocked by safety.allowDestructive=false.
- navigation.topNavigationBar.node.update: destructive; blocked by safety.allowDestructive=false.
- navigation.topNavigationBar.node.delete: destructive; blocked by safety.allowDestructive=false.
- profiles.setSingleValueProperty: destructive; blocked by safety.allowDestructive=false.
- profiles.setMultiValueProperty: destructive; blocked by safety.allowDestructive=false.
- profiles.hideSuggestion: destructive; blocked by safety.allowDestructive=false.
- profiles.setMyProfilePicture: destructive; blocked by safety.allowDestructive=false.
- profiles.createPersonalSitesBulk: destructive; blocked by safety.allowDestructive=false.
- profiles.createPersonalSite: destructive; blocked by safety.allowDestructive=false.
- profiles.shareAllSocialData: destructive; blocked by safety.allowDestructive=false.
- appCatalog.add: destructive; blocked by safety.allowDestructive=false.
- appCatalog.syncTeams: destructive; blocked by safety.allowDestructive=false.
- appCatalog.siteCollectionSyncTeams: destructive; blocked by safety.allowDestructive=false.
- appCatalog.siteCollection.add: destructive; blocked by safety.allowDestructive=false.
- appCatalog.siteCollection.app.deploy: destructive; blocked by safety.allowDestructive=false.
- appCatalog.siteCollection.app.retract: destructive; blocked by safety.allowDestructive=false.
- appCatalog.siteCollection.app.install: destructive; blocked by safety.allowDestructive=false.
- appCatalog.siteCollection.app.uninstall: destructive; blocked by safety.allowDestructive=false.
- appCatalog.siteCollection.app.upgrade: destructive; blocked by safety.allowDestructive=false.
- appCatalog.siteCollection.app.remove: destructive; blocked by safety.allowDestructive=false.
- appCatalog.app.deploy: destructive; blocked by safety.allowDestructive=false.
- appCatalog.app.retract: destructive; blocked by safety.allowDestructive=false.
- appCatalog.app.install: destructive; blocked by safety.allowDestructive=false.
- appCatalog.app.uninstall: destructive; blocked by safety.allowDestructive=false.
- appCatalog.app.upgrade: destructive; blocked by safety.allowDestructive=false.
- appCatalog.app.remove: destructive; blocked by safety.allowDestructive=false.
- webparts.import: destructive; blocked by safety.allowDestructive=false.
- webparts.definition.saveChanges: destructive; blocked by safety.allowDestructive=false.
- webparts.definition.move: destructive; blocked by safety.allowDestructive=false.
- webparts.definition.close: destructive; blocked by safety.allowDestructive=false.
- webparts.definition.open: destructive; blocked by safety.allowDestructive=false.
- webparts.definition.delete: destructive; blocked by safety.allowDestructive=false.
- clientPages.create: destructive; blocked by safety.allowDestructive=false.
- clientPages.checkout: destructive; blocked by safety.allowDestructive=false.
- clientPages.save: destructive; blocked by safety.allowDestructive=false.
- clientPages.publish: destructive; blocked by safety.allowDestructive=false.
- clientPages.discardCheckout: destructive; blocked by safety.allowDestructive=false.
- clientPages.schedulePublish: destructive; blocked by safety.allowDestructive=false.
- clientPages.saveAsTemplate: destructive; blocked by safety.allowDestructive=false.
- clientPages.promoteNews: destructive; blocked by safety.allowDestructive=false.
- clientPages.addRepost: destructive; blocked by safety.allowDestructive=false.
- clientPages.shareByEmail: destructive; blocked by safety.allowDestructive=false.
- siteDesigns.create: destructive; blocked by safety.allowDestructive=false.
- siteDesigns.update: destructive; blocked by safety.allowDestructive=false.
- siteDesigns.delete: destructive; blocked by safety.allowDestructive=false.
- siteDesigns.apply: destructive; blocked by safety.allowDestructive=false.
- siteDesigns.addTask: destructive; blocked by safety.allowDestructive=false.
- siteScripts.delete: destructive; blocked by safety.allowDestructive=false.
- siteScripts.update: destructive; blocked by safety.allowDestructive=false.
- siteScripts.executeAction: destructive; blocked by safety.allowDestructive=false.
- relatedItems.addSingleLink: destructive; blocked by safety.allowDestructive=false.
- relatedItems.addLinkToUrl: destructive; blocked by safety.allowDestructive=false.
- relatedItems.addLinkFromUrl: destructive; blocked by safety.allowDestructive=false.
- relatedItems.deleteSingleLink: destructive; blocked by safety.allowDestructive=false.
- features.add: destructive; blocked by safety.allowDestructive=false.
- features.remove: destructive; blocked by safety.allowDestructive=false.
- groupSiteManager.clearTeamsCache: destructive; blocked by safety.allowDestructive=false.
- groupSiteManager.create: destructive; blocked by safety.allowDestructive=false.
- groupSiteManager.createGroup: destructive; blocked by safety.allowDestructive=false.
- groupSiteManager.delete: destructive; blocked by safety.allowDestructive=false.
- groupSiteManager.ensureTeam: destructive; blocked by safety.allowDestructive=false.
- groupSiteManager.ensureTeamEx: destructive; blocked by safety.allowDestructive=false.
- groupSiteManager.pinToTeam: destructive; blocked by safety.allowDestructive=false.
- social.follow: destructive; blocked by safety.allowDestructive=false.
- social.stopFollowing: destructive; blocked by safety.allowDestructive=false.
- utilities.sendEmail: destructive; blocked by safety.allowDestructive=false.
- siteDesigns.grantRights: destructive; blocked by safety.allowDestructive=false.
- siteDesigns.revokeRights: destructive; blocked by safety.allowDestructive=false.
- siteDesigns.addTaskCurrentWeb: destructive; blocked by safety.allowDestructive=false.
- siteScripts.create: destructive; blocked by safety.allowDestructive=false.
- attachments.content.set: destructive; blocked by safety.allowDestructive=false.
- attachments.recycle: destructive; blocked by safety.allowDestructive=false.
- items.deleteWithParameters: destructive; blocked by safety.allowDestructive=false.
- items.uploadImage: destructive; blocked by safety.allowDestructive=false.
- items.setLike: destructive; blocked by safety.allowDestructive=false.
- items.rate: destructive; blocked by safety.allowDestructive=false.
- lists.userCustomActions.add: destructive; blocked by safety.allowDestructive=false.
- lists.userCustomActions.clear: destructive; blocked by safety.allowDestructive=false.
- lists.userCustomActions.update: destructive; blocked by safety.allowDestructive=false.
- lists.userCustomActions.delete: destructive; blocked by safety.allowDestructive=false.
- customActions.site.add: destructive; blocked by safety.allowDestructive=false.
- customActions.site.clear: destructive; blocked by safety.allowDestructive=false.
- customActions.site.update: destructive; blocked by safety.allowDestructive=false.
- customActions.site.delete: destructive; blocked by safety.allowDestructive=false.
- customActions.web.add: destructive; blocked by safety.allowDestructive=false.
- customActions.web.clear: destructive; blocked by safety.allowDestructive=false.
- customActions.web.update: destructive; blocked by safety.allowDestructive=false.
- customActions.web.delete: destructive; blocked by safety.allowDestructive=false.
- WorkflowServices coverage is not represented: the documented client API is CSOM, not a REST endpoint.

## Limitations

- No evidence entries: executed=0 and verified=0 are accurate, not a bug.
- Generated means a URL was constructed from documented examples; it does not prove the endpoint works.
- HEAD, OPTIONS, PUT and MERGE are generated as variations, not as declared endpoint methods.
- Microsoft Graph uses the separate bearer-token transport; a caller-supplied token provider is required and no live request is recorded.
- Coverage counts implementations, not vulnerabilities. No finding is confirmed without manual evidence.