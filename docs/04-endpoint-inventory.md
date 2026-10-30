# Endpoint coverage report

Generated: 2026-10-30T10:23:34.378Z

| Metric | Value |
| --- | --- |
| Registered endpoints | 316 |
| Implemented | 316 |
| Baseline generated | 188 |
| Executed | 0 |
| Verified | 0 |

## Modules

| Module | Endpoints | Methods | Payload families |
| --- | --- | --- | --- |
| approvals | 5 | GET, POST | idor, input-validation, authz, business-logic, mass-assignment |
| attachments | 4 | GET, POST, DELETE | idor, authz, input-validation, path-traversal, xss, ssrf, method-confusion |
| batch | 3 | POST | method-confusion, authz, mass-assignment, input-validation, business-logic, xss |
| changes | 3 | POST | input-validation, business-logic, authz |
| comments | 8 | GET, POST, DELETE | idor, xss, authz, sqli, mass-assignment, input-validation, business-logic, method-confusion |
| configuration | 8 | GET, POST | idor, authz, input-validation, odata-injection, business-logic, mass-assignment, method-confusion |
| content-types | 21 | GET, POST, PATCH, MERGE, DELETE | odata-injection, authz, mass-assignment, xss, input-validation, business-logic, idor, method-confusion |
| context | 1 | POST | input-validation |
| fields | 23 | GET, POST, PATCH, MERGE, DELETE | mass-assignment, xss, input-validation, business-logic, authz, idor, method-confusion, odata-injection |
| files | 42 | GET, DELETE, POST, MERGE | idor, path-traversal, authz, method-confusion, input-validation, ssrf, xss, mass-assignment, odata-injection, business-logic |
| folders | 17 | GET, DELETE, POST, MERGE | idor, path-traversal, authz, method-confusion, input-validation, odata-injection, business-logic, xss, mass-assignment |
| groups | 17 | GET, POST, MERGE, DELETE | mass-assignment, xss, input-validation, business-logic, authz, idor, odata-injection, method-confusion |
| items | 20 | GET, POST, PATCH, MERGE, DELETE | idor, odata-injection, input-validation, authz, mass-assignment, xss, sqli, method-confusion, business-logic |
| libraries | 1 | GET | odata-injection, authz, input-validation |
| lists | 33 | POST, GET, PATCH, MERGE, DELETE | mass-assignment, xss, input-validation, business-logic, odata-injection, authz, method-confusion, idor |
| permissions | 33 | POST, GET, MERGE, DELETE | idor, authz, business-logic, method-confusion, input-validation, odata-injection, mass-assignment |
| recycle-bin | 8 | GET, POST | odata-injection, idor, authz, path-traversal, business-logic |
| search | 4 | GET, POST | odata-injection, xss, sqli, input-validation, authz |
| sharing | 11 | POST | idor, authz, business-logic, input-validation, ssrf, mass-assignment, xss |
| sites | 12 | GET, POST, PATCH, MERGE | idor, authz, input-validation, ssrf, business-logic, mass-assignment, method-confusion |
| taxonomy | 4 | GET | authz, odata-injection, idor |
| users | 11 | GET, POST, MERGE, DELETE | odata-injection, idor, authz, input-validation, xss, business-logic, mass-assignment, method-confusion |
| views | 14 | GET, POST, PATCH, MERGE, DELETE | odata-injection, xss, mass-assignment, input-validation, authz, idor, method-confusion, business-logic |
| webhooks | 5 | GET, POST, PATCH, DELETE | authz, input-validation, ssrf, business-logic, xss, idor |
| webs | 8 | GET, POST, PATCH, MERGE, DELETE | idor, odata-injection, authz, mass-assignment, xss, input-validation, business-logic, method-confusion |

## Declared methods

- DELETE: 18 endpoint(s)
- GET: 135 endpoint(s)
- MERGE: 14 endpoint(s)
- PATCH: 10 endpoint(s)
- POST: 156 endpoint(s)

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
- lists.userCustomActions: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
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
- lists.forms.byPageType: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.forms.byId: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.form.byPageType: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.collection.get: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.shareLink: destructive; blocked by safety.allowDestructive=false.
- lists.unshareLink: destructive; blocked by safety.allowDestructive=false.
- lists.updateDocumentSharingInfo: destructive; blocked by safety.allowDestructive=false.
- lists.roleAssignments.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.roleAssignments.add: destructive; blocked by safety.allowDestructive=false.
- lists.roleAssignments.remove: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
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
- items.fieldValuesForEdit: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.effectiveBasePermissionsForUI: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.contentType: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.versions: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.version.get: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.attachments: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.version.delete: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.version.delete: destructive; blocked by safety.allowDestructive=false.
- items.attachment.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.attachment.add: destructive; blocked by safety.allowDestructive=false.
- items.attachment.get: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.attachment.delete: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.attachment.delete: destructive; blocked by safety.allowDestructive=false.
- items.comments: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.comment.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.comment.add: destructive; blocked by safety.allowDestructive=false.
- items.comments.clear: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
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
- items.shareObject: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.shareObject: destructive; blocked by safety.allowDestructive=false.
- items.sharingSettings: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.sharing.checkPermissions: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.unshareObject: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.unshareObject: destructive; blocked by safety.allowDestructive=false.
- items.deleteLinkByKind: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.deleteLinkByKind: destructive; blocked by safety.allowDestructive=false.
- items.unshareLink: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.unshareLink: destructive; blocked by safety.allowDestructive=false.
- items.roleAssignments: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.roleAssignments.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.roleAssignments.add: destructive; blocked by safety.allowDestructive=false.
- items.roleAssignments.remove: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- items.roleAssignments.remove: destructive; blocked by safety.allowDestructive=false.
- items.effectiveBasePermissions: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.collection: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.add: destructive; blocked by safety.allowDestructive=false.
- fields.getByTitle: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.updateByTitle: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.updateByTitle: destructive; blocked by safety.allowDestructive=false.
- fields.deleteByTitle: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.deleteByTitle: destructive; blocked by safety.allowDestructive=false.
- fields.getByInternalNameOrTitle: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.getById: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.addField: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.addField: destructive; blocked by safety.allowDestructive=false.
- fields.createFieldAsXml: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.createFieldAsXml: destructive; blocked by safety.allowDestructive=false.
- fields.addDependentLookupField: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.addDependentLookupField: destructive; blocked by safety.allowDestructive=false.
- fields.setShowInDisplayForm: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.setShowInDisplayForm: destructive; blocked by safety.allowDestructive=false.
- fields.setShowInEditForm: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.setShowInEditForm: destructive; blocked by safety.allowDestructive=false.
- fields.setShowInNewForm: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- fields.setShowInNewForm: destructive; blocked by safety.allowDestructive=false.
- contentTypes.collection: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.add: destructive; blocked by safety.allowDestructive=false.
- contentTypes.byId: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.updateById: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.updateById: destructive; blocked by safety.allowDestructive=false.
- contentTypes.deleteById: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.deleteById: destructive; blocked by safety.allowDestructive=false.
- contentTypes.addAvailable: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.addAvailable: destructive; blocked by safety.allowDestructive=false.
- contentTypes.fieldLinks: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.fields: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.parent: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.collection: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- views.add: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
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
- permissions.list.userEffective: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- permissions.item.userEffective: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- permissions.item.breakInheritance: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- permissions.item.breakInheritance: destructive; blocked by safety.allowDestructive=false.
- permissions.item.resetInheritance: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- permissions.item.resetInheritance: destructive; blocked by safety.allowDestructive=false.
- roleAssignments.list: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- roleAssignments.listBindings: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- roleAssignments.itemBindings: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- lists.rootFolder: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
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
- contentTypes.workflowAssociations: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- contentTypes.fieldLinks.byId: no complete documented example for every path placeholder, so a baseline cannot be generated automatically.
- WorkflowServices coverage is not represented: the documented client API is CSOM, not a REST endpoint.

## Limitations

- No evidence entries: executed=0 and verified=0 are accurate, not a bug.
- Generated means a URL was constructed from documented examples; it does not prove the endpoint works.
- HEAD, OPTIONS, PUT and MERGE are generated as variations, not as declared endpoint methods.
- Microsoft Graph uses the separate bearer-token transport; a caller-supplied token provider is required and no live request is recorded.
- Coverage counts implementations, not vulnerabilities. No finding is confirmed without manual evidence.