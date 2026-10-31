# PnPjs endpoints added in this update

Source: installed `@pnp/sp` 4.9.0 module source. These are request templates and do not represent live tenant verification.

| Endpoint ID | Method | Request path |
| --- | --- | --- |
| `navigation.quicklaunch` | GET | `/_api/web/navigation/quicklaunch` |
| `navigation.topNavigationBar` | GET | `/_api/web/navigation/topnavigationbar` |
| `navigation.addQuicklaunch` | POST | `/_api/web/navigation/quicklaunch` |
| `navigation.moveAfter` | POST | `/_api/web/navigation/MoveAfter` |
| `navigation.menuState` | POST | `/_api/navigation/MenuState` |
| `navigation.menuNodeKey` | POST | `/_api/navigation/MenuNodeKey` |
| `navigation.quicklaunch.node` | GET | `/_api/web/navigation/quicklaunch({nodeId:int})` |
| `navigation.quicklaunch.node.update` | POST, MERGE | `/_api/web/navigation/quicklaunch({nodeId:int})` |
| `navigation.quicklaunch.node.delete` | DELETE | `/_api/web/navigation/quicklaunch({nodeId:int})` |
| `navigation.topNavigationBar.add` | POST | `/_api/web/navigation/topnavigationbar` |
| `navigation.topNavigationBar.moveAfter` | POST | `/_api/web/navigation/topnavigationbar/MoveAfter` |
| `navigation.quicklaunch.node.children` | GET | `/_api/web/navigation/quicklaunch({nodeId:int})/children` |
| `navigation.topNavigationBar.node` | GET | `/_api/web/navigation/topnavigationbar({nodeId:int})` |
| `navigation.topNavigationBar.node.update` | POST, MERGE | `/_api/web/navigation/topnavigationbar({nodeId:int})` |
| `navigation.topNavigationBar.node.delete` | DELETE | `/_api/web/navigation/topnavigationbar({nodeId:int})` |
| `navigation.topNavigationBar.node.children` | GET | `/_api/web/navigation/topnavigationbar({nodeId:int})/children` |
| `profiles.myProperties` | GET | `/_api/sp.userprofiles.peoplemanager/getmyproperties` |
| `profiles.propertiesFor` | GET | `/_api/sp.userprofiles.peoplemanager/getpropertiesfor(@v)?@v={loginName:odataEncoded}` |
| `profiles.userProfileProperty` | GET | `/_api/sp.userprofiles.peoplemanager/getuserprofilepropertyfor(accountname=@v,propertyname='{propertyName:odata}')?@v={loginName:odataEncoded}` |
| `profiles.followers` | GET | `/_api/sp.userprofiles.peoplemanager/getmyfollowers` |
| `profiles.followersFor` | GET | `/_api/sp.userprofiles.peoplemanager/getfollowersfor(@v)?@v={loginName:odataEncoded}` |
| `profiles.peopleFollowedBy` | GET | `/_api/sp.userprofiles.peoplemanager/getpeoplefollowedby(@v)?@v={loginName:odataEncoded}` |
| `profiles.following` | GET | `/_api/sp.userprofiles.peoplemanager.isfollowing(possiblefolloweraccountname=@v,possiblefolloweeaccountname=@y)?@v={follower:odataEncoded}&@y={followee:odataEncoded}` |
| `profiles.followedTags` | GET | `/_api/sp.userprofiles.peoplemanager/getfollowedtags({maxCount:int})` |
| `profiles.setSingleValueProperty` | POST | `/_api/sp.userprofiles.peoplemanager/SetSingleValueProfileProperty` |
| `profiles.setMultiValueProperty` | POST | `/_api/sp.userprofiles.peoplemanager/SetMultiValuedProfileProperty` |
| `profiles.hideSuggestion` | POST | `/_api/sp.userprofiles.peoplemanager/hidesuggestion(@v)?@v={loginName:odataEncoded}` |
| `profiles.setMyProfilePicture` | POST | `/_api/sp.userprofiles.peoplemanager/setmyprofilepicture` |
| `profiles.createPersonalSitesBulk` | POST | `/_api/sp.userprofiles.profileloader.getprofileloader/createpersonalsiteenqueuebulk` |
| `profiles.ownerUserProfile` | POST | `/_api/sp.userprofiles.profileloader.getowneruserprofile` |
| `profiles.userProfile` | POST | `/_api/sp.userprofiles.profileloader.getprofileloader/getuserprofile` |
| `profiles.createPersonalSite` | POST | `/_api/sp.userprofiles.profileloader.getprofileloader/getuserprofile/createpersonalsiteenque({interactiveRequest:raw})` |
| `profiles.shareAllSocialData` | POST | `/_api/sp.userprofiles.profileloader.getprofileloader/getuserprofile/shareallsocialdata({share:raw})` |
| `profiles.peoplePickerResolve` | POST | `/_api/sp.ui.applicationpages.clientpeoplepickerwebserviceinterface.clientpeoplepickerresolveuser` |
| `profiles.peoplePickerSearch` | POST | `/_api/sp.ui.applicationpages.clientpeoplepickerwebserviceinterface.clientpeoplepickersearchuser` |
| `hubsites.collection` | GET | `/_api/hubsites` |
| `hubsites.getById` | GET | `/_api/hubsites/GetById?hubSiteId='{hubSiteId:odata}'` |
| `appCatalog.availableApps` | GET | `/_api/web/tenantappcatalog/AvailableApps` |
| `appCatalog.siteCollectionAvailableApps` | GET | `/_api/web/sitecollectionappcatalog/AvailableApps` |
| `appCatalog.add` | POST | `/_api/web/tenantappcatalog/AvailableApps/add(overwrite={overwrite:raw},url='{fileName:odata}')` |
| `appCatalog.syncTeams` | POST | `/_api/web/tenantappcatalog/SyncSolutionToTeams(id={appItemId:int})` |
| `appCatalog.getById` | GET | `/_api/web/tenantappcatalog/AvailableApps/getById('{appId:raw}')` |
| `appCatalog.siteCollectionSyncTeams` | POST | `/_api/web/sitecollectionappcatalog/SyncSolutionToTeams(id={appItemId:int})` |
| `appCatalog.siteCollection.getById` | GET | `/_api/web/sitecollectionappcatalog/AvailableApps/getById('{appId:raw}')` |
| `appCatalog.siteCollection.add` | POST | `/_api/web/sitecollectionappcatalog/AvailableApps/add(overwrite={overwrite:raw},url='{fileName:odata}')` |
| `appCatalog.siteCollection.app.deploy` | POST | `/_api/web/sitecollectionappcatalog/AvailableApps/getById('{appId:raw}')/Deploy({skipFeatureDeployment:raw})` |
| `appCatalog.siteCollection.app.retract` | POST | `/_api/web/sitecollectionappcatalog/AvailableApps/getById('{appId:raw}')/Retract` |
| `appCatalog.siteCollection.app.install` | POST | `/_api/web/sitecollectionappcatalog/AvailableApps/getById('{appId:raw}')/Install` |
| `appCatalog.siteCollection.app.uninstall` | POST | `/_api/web/sitecollectionappcatalog/AvailableApps/getById('{appId:raw}')/Uninstall` |
| `appCatalog.siteCollection.app.upgrade` | POST | `/_api/web/sitecollectionappcatalog/AvailableApps/getById('{appId:raw}')/Upgrade` |
| `appCatalog.siteCollection.app.remove` | POST | `/_api/web/sitecollectionappcatalog/AvailableApps/getById('{appId:raw}')/Remove` |
| `appCatalog.app.deploy` | POST | `/_api/web/tenantappcatalog/AvailableApps/getById('{appId:raw}')/Deploy({skipFeatureDeployment:raw})` |
| `appCatalog.app.retract` | POST | `/_api/web/tenantappcatalog/AvailableApps/getById('{appId:raw}')/Retract` |
| `appCatalog.app.install` | POST | `/_api/web/tenantappcatalog/AvailableApps/getById('{appId:raw}')/Install` |
| `appCatalog.app.uninstall` | POST | `/_api/web/tenantappcatalog/AvailableApps/getById('{appId:raw}')/Uninstall` |
| `appCatalog.app.upgrade` | POST | `/_api/web/tenantappcatalog/AvailableApps/getById('{appId:raw}')/Upgrade` |
| `appCatalog.app.remove` | POST | `/_api/web/tenantappcatalog/AvailableApps/getById('{appId:raw}')/Remove` |
| `webparts.getClientSideWebParts` | GET | `/_api/web/GetClientSideWebParts` |
| `webparts.definitions` | GET | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/WebParts?@pagePath={pagePath:odataEncoded}` |
| `webparts.definition.byId` | GET | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/WebParts/getbyid('{webPartId:raw}')?@pagePath={pagePath:odataEncoded}` |
| `webparts.definition.byControlId` | GET | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/WebParts/getByControlId('{controlId:raw}')?@pagePath={pagePath:odataEncoded}` |
| `webparts.definition.webpart` | GET | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/WebParts/getbyid('{webPartId:raw}')/webpart?@pagePath={pagePath:odataEncoded}` |
| `webparts.export` | POST | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/ExportWebPart?@pagePath={pagePath:odataEncoded}` |
| `webparts.import` | POST | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/ImportWebPart?@pagePath={pagePath:odataEncoded}` |
| `webparts.definition.saveChanges` | POST | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/WebParts/getById('{webPartId:raw}')/SaveWebPartChanges?@pagePath={pagePath:odataEncoded}` |
| `webparts.definition.move` | POST | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/WebParts/getById('{webPartId:raw}')/MoveWebPartTo(zoneID='{zoneId:odata}',zoneIndex={zoneIndex:int})?@pagePath={pagePath:odataEncoded}` |
| `webparts.definition.close` | POST | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/WebParts/getById('{webPartId:raw}')/CloseWebPart?@pagePath={pagePath:odataEncoded}` |
| `webparts.definition.open` | POST | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/WebParts/getById('{webPartId:raw}')/OpenWebPart?@pagePath={pagePath:odataEncoded}` |
| `webparts.definition.delete` | POST | `/_api/web/GetFileByServerRelativePath(DecodedUrl=@pagePath)/GetLimitedWebPartManager(scope=1)/WebParts/getById('{webPartId:raw}')/DeleteWebPart?@pagePath={pagePath:odataEncoded}` |
| `clientPages.get` | GET | `/_api/sitepages/pages({pageId:int})` |
| `clientPages.create` | POST | `/_api/sitepages/pages` |
| `clientPages.checkout` | POST | `/_api/sitepages/pages({pageId:int})/checkoutpage` |
| `clientPages.save` | POST | `/_api/sitepages/pages({pageId:int})/savepage` |
| `clientPages.publish` | POST | `/_api/sitepages/pages({pageId:int})/publish` |
| `clientPages.discardCheckout` | POST | `/_api/sitepages/pages({pageId:int})/discardPage` |
| `clientPages.schedulePublish` | POST | `/_api/sitepages/pages({pageId:int})/SchedulePublish` |
| `clientPages.saveAsTemplate` | POST | `/_api/sitepages/pages({pageId:int})/SavePageAsTemplate` |
| `clientPages.promoteNews` | POST | `/_api/sitepages/pages({pageId:int})/promoteToNews` |
| `clientPages.addRepost` | POST | `/_api/sitepages/pages/reposts` |
| `clientPages.shareByEmail` | POST | `/_api/SP.Publishing.RichSharing/SharePageByEmail` |
| `clientPages.ensureSitePagesLibrary` | POST | `/_api/lists/EnsureClientRenderedSitePagesLibrary` |
| `publishingSitePageService.currentUserMemberships` | GET | `/_api/SP.Publishing.SitePageService.GetCurrentUserMemberships` |
| `siteDesigns.get` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns` |
| `siteDesigns.create` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteDesign` |
| `siteDesigns.getMetadata` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignMetadata` |
| `siteDesigns.update` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteDesign` |
| `siteDesigns.delete` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteDesign` |
| `siteDesigns.apply` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.ApplySiteDesign` |
| `siteDesigns.addTask` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.AddSiteDesignTask` |
| `siteDesigns.getTask` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignTask` |
| `siteScripts.get` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScripts` |
| `siteScripts.metadata` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptMetadata` |
| `siteScripts.delete` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteScript` |
| `siteScripts.update` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.UpdateSiteScript` |
| `siteScripts.executeAction` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.ExecuteSiteScriptAction` |
| `relatedItems.get` | POST | `/_api/SP.RelatedItemManager.GetRelatedItems` |
| `relatedItems.getPageOne` | POST | `/_api/SP.RelatedItemManager.GetPageOneRelatedItems` |
| `relatedItems.addSingleLink` | POST | `/_api/SP.RelatedItemManager.AddSingleLink` |
| `relatedItems.addLinkToUrl` | POST | `/_api/SP.RelatedItemManager.AddSingleLinkToUrl` |
| `relatedItems.addLinkFromUrl` | POST | `/_api/SP.RelatedItemManager.AddSingleLinkFromUrl` |
| `relatedItems.deleteSingleLink` | POST | `/_api/SP.RelatedItemManager.DeleteSingleLink` |
| `features.collection` | GET | `/_api/web/features` |
| `features.getById` | GET | `/_api/web/features('{featureId:raw}')` |
| `features.add` | POST | `/_api/web/features/add` |
| `features.remove` | POST | `/_api/web/features/remove` |
| `groupSiteManager.canCreateGroup` | GET | `/_api/groupsitemanager/CanUserCreateGroup` |
| `groupSiteManager.clearTeamsCache` | POST | `/_api/groupsitemanager/ClearCurrentUserTeamsCache` |
| `groupSiteManager.create` | POST | `/_api/groupsitemanager/Create` |
| `groupSiteManager.createGroup` | POST | `/_api/groupsitemanager/CreateGroup` |
| `groupSiteManager.delete` | POST | `/_api/groupsitemanager/Delete` |
| `groupSiteManager.ensureTeam` | POST | `/_api/groupsitemanager/EnsureTeamForGroup` |
| `groupSiteManager.ensureTeamEx` | POST | `/_api/groupsitemanager/EnsureTeamForGroupEx` |
| `groupSiteManager.orgLabels` | POST | `/_api/groupsitemanager/GetAllOrgLabels` |
| `groupSiteManager.joinedTeams` | POST | `/_api/groupsitemanager/GetCurrentUserJoinedTeams` |
| `groupSiteManager.sharedChannels` | POST | `/_api/groupsitemanager/GetCurrentUserSharedChannelMemberGroups` |
| `groupSiteManager.teamGroups` | POST | `/_api/groupsitemanager/GetCurrentUserTeamConnectedMemberGroups` |
| `groupSiteManager.creationContext` | GET | `/_api/groupsitemanager/GetGroupCreationContext` |
| `groupSiteManager.conversionData` | GET | `/_api/groupsitemanager/GetGroupSiteConversionData` |
| `groupSiteManager.siteStatus` | POST | `/_api/groupsitemanager/GetSiteStatus` |
| `groupSiteManager.teamChannelFilesUrl` | POST | `/_api/groupsitemanager/GetTeamChannelFilesUrl` |
| `groupSiteManager.teamChannels` | POST | `/_api/groupsitemanager/GetTeamChannels` |
| `groupSiteManager.teamChannelsDirect` | POST | `/_api/groupsitemanager/GetTeamChannelsDirect` |
| `groupSiteManager.teamChannelsEx` | POST | `/_api/groupsitemanager/GetTeamChannelsEx` |
| `groupSiteManager.teamChannelsByUrl` | POST | `/_api/groupsitemanager/GetTeamChannelsWithSiteUrl` |
| `groupSiteManager.userSharedChannels` | POST | `/_api/groupsitemanager/GetUserSharedChannelMemberGroups` |
| `groupSiteManager.userTeamGroups` | POST | `/_api/groupsitemanager/GetUserTeamConnectedMemberGroups` |
| `groupSiteManager.validSiteUrl` | POST | `/_api/groupsitemanager/GetValidSiteUrlFromAlias` |
| `groupSiteManager.teamifyHidden` | POST | `/_api/groupsitemanager/IsTeamifyPromptHidden` |
| `groupSiteManager.notebook` | POST | `/_api/groupsitemanager/Notebook` |
| `groupSiteManager.pinToTeam` | POST | `/_api/groupsitemanager/PinToTeam` |
| `regionalSettings.settings` | GET | `/_api/web/regionalsettings` |
| `regionalSettings.installedLanguages` | GET | `/_api/web/regionalsettings/installedlanguages` |
| `regionalSettings.timeZones` | GET | `/_api/web/timezones` |
| `regionalSettings.timeZoneById` | POST | `/_api/web/timezones/GetById({timeZoneId:int})` |
| `regionalSettings.utcToLocal` | POST | `/_api/web/timezone/utctolocaltime('{dateTime:odata}')` |
| `regionalSettings.localToUtc` | POST | `/_api/web/timezone/localtimetoutc('{dateTime:odata}')` |
| `social.follow` | POST | `/_api/social.following/follow` |
| `social.isFollowed` | POST | `/_api/social.following/isfollowed` |
| `social.stopFollowing` | POST | `/_api/social.following/stopfollowing` |
| `social.followedSitesUri` | GET | `/_api/social.following/FollowedSitesUri` |
| `social.followedDocumentsUri` | GET | `/_api/social.following/FollowedDocumentsUri` |
| `social.my.followed` | GET | `/_api/social.following/my/followed(types={types:int})` |
| `social.my.followedCount` | GET | `/_api/social.following/my/followedcount(types={types:int})` |
| `social.my.followers` | GET | `/_api/social.following/my/followers` |
| `social.my.suggestions` | GET | `/_api/social.following/my/suggestions` |
| `utilities.currentUserEmails` | POST | `/_api/SP.Utilities.Utility.GetCurrentUserEmailAddresses` |
| `utilities.resolvePrincipal` | POST | `/_api/SP.Utilities.Utility.ResolvePrincipalInCurrentContext` |
| `utilities.searchPrincipals` | POST | `/_api/SP.Utilities.Utility.SearchPrincipalsUsingContextWeb` |
| `utilities.expandGroups` | POST | `/_api/SP.Utilities.Utility.ExpandGroupsToPrincipals` |
| `utilities.sendEmail` | POST | `/_api/SP.Utilities.Utility.SendEmail` |
| `utilities.createInvitationEmailBody` | POST | `/_api/SP.Utilities.Utility.CreateEmailBodyForInvitation` |
| `profiles.editProfileLink` | GET | `/_api/sp.userprofiles.peoplemanager/EditProfileLink` |
| `profiles.myPeopleListPublic` | GET | `/_api/sp.userprofiles.peoplemanager/IsMyPeopleListPublic` |
| `profiles.amIFollowedBy` | GET | `/_api/sp.userprofiles.peoplemanager/amifollowedby(@v)?@v={loginName:odataEncoded}` |
| `profiles.amIFollowing` | GET | `/_api/sp.userprofiles.peoplemanager/amifollowing(@v)?@v={loginName:odataEncoded}` |
| `profiles.trendingTags` | GET | `/_api/sp.userprofiles.peoplemanager.gettrendingtags` |
| `siteDesigns.rights` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRights` |
| `siteDesigns.grantRights` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GrantSiteDesignRights` |
| `siteDesigns.revokeRights` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.RevokeSiteDesignRights` |
| `siteDesigns.addTaskCurrentWeb` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.AddSiteDesignTaskToCurrentWeb` |
| `siteDesigns.run` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRun` |
| `siteDesigns.runStatus` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesignRunStatus` |
| `siteScripts.fromList` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteScriptFromList` |
| `siteScripts.fromWeb` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.getSiteScriptFromWeb` |
| `siteScripts.create` | POST | `/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.CreateSiteScript(Title=@title,Description=@desc)?@title={title:odataEncoded}&@desc={description:odataEncoded}` |
| `attachments.content` | GET | `/_api/web/lists({listId:guid})/items({itemId:int})/AttachmentFiles({fileName:odata})/$value` |
| `attachments.content.set` | POST (X-HTTP-Method: PUT) | `/_api/web/lists({listId:guid})/items({itemId:int})/AttachmentFiles({fileName:odata})/$value` |
| `attachments.recycle` | POST (X-HTTP-Method: DELETE) | `/_api/web/lists({listId:guid})/items({itemId:int})/AttachmentFiles({fileName:odata})/recycleObject` |
| `items.deleteWithParameters` | POST | `/_api/web/lists({listId:guid})/items({itemId:int})/DeleteWithParameters` |
| `items.wopiFrameUrl` | POST | `/_api/web/lists({listId:guid})/items({itemId:int})/getWOPIFrameUrl(@action)?@action={action:int}` |
| `items.uploadImage` | POST | `/_api/web/UploadImage(listTitle=@a1,imageName=@a2,listId=@a3,itemId=@a4)?@a1={listTitle:odataEncoded}&@a2={imageName:odataEncoded}&@a3={listId:odataEncoded}&@a4={itemId:int}` |
| `items.likedBy` | POST | `/_api/web/lists({listId:guid})/items({itemId:int})/likedBy` |
| `items.likedByInformation` | GET | `/_api/web/lists({listId:guid})/items({itemId:int})/likedByInformation` |
| `items.setLike` | POST | `/_api/Microsoft.Office.Server.ReputationModel.Reputation.SetLike(listID=@a1,itemID=@a2,like=@a3)` |
| `items.rate` | POST | `/_api/Microsoft.Office.Server.ReputationModel.Reputation.SetRating(listID=@a1,itemID=@a2,rating=@a3)` |

## Follow-up audit: collection and custom-action routes

The list collection GET was absent from the Lists module even though the
Libraries alias used the same route with a document-library filter. It is now
available as an unfiltered collection request. PnPjs 4.9.0 also exposes form
lookup by ID and user custom action CRUD at list, web and site scope.

| Endpoint ID | Method | Request path |
| --- | --- | --- |
| `lists.collection.get` | GET | `/_api/web/lists` |
| `lists.forms.byId` | GET | `/_api/web/lists({listId:guid})/forms('{formId:raw}')` |
| `lists.userCustomActions.add` | POST | `/_api/web/lists({listId:guid})/usercustomactions` |
| `lists.userCustomActions.clear` | POST | `/_api/web/lists({listId:guid})/usercustomactions/clear` |
| `lists.userCustomActions.byId` | GET | `/_api/web/lists({listId:guid})/usercustomactions('{actionId:raw}')` |
| `lists.userCustomActions.update` | MERGE | `/_api/web/lists({listId:guid})/usercustomactions('{actionId:raw}')` |
| `lists.userCustomActions.delete` | DELETE | `/_api/web/lists({listId:guid})/usercustomactions('{actionId:raw}')` |
| `customActions.site` | GET | `/_api/site/usercustomactions` |
| `customActions.site.add` | POST | `/_api/site/usercustomactions` |
| `customActions.site.clear` | POST | `/_api/site/usercustomactions/clear` |
| `customActions.site.byId` | GET | `/_api/site/usercustomactions('{actionId:raw}')` |
| `customActions.site.update` | MERGE | `/_api/site/usercustomactions('{actionId:raw}')` |
| `customActions.site.delete` | DELETE | `/_api/site/usercustomactions('{actionId:raw}')` |
| `customActions.web` | GET | `/_api/web/usercustomactions` |
| `customActions.web.add` | POST | `/_api/web/usercustomactions` |
| `customActions.web.clear` | POST | `/_api/web/usercustomactions/clear` |
| `customActions.web.byId` | GET | `/_api/web/usercustomactions('{actionId:raw}')` |
| `customActions.web.update` | MERGE | `/_api/web/usercustomactions('{actionId:raw}')` |
| `customActions.web.delete` | DELETE | `/_api/web/usercustomactions('{actionId:raw}')` |
| `web.applyWebTemplate` | POST | `/_api/web/applywebtemplate(webTemplate='{template:odataEncoded}')` |

## Localized resource lookups from the installed type declarations

`regional-settings/types.d.ts` declares `IUserResources`, and the adjacent
module declaration files attach those methods to webs, lists, fields, content
types and user custom actions. `getValueForUICultureBinder` in `funcs.js`
confirms each call is a POST with `{ cultureName }`. All attached resource
routes are now represented:

| Endpoint ID | Request path |
| --- | --- |
| `regionalSettings.web.titleResource` | `/_api/web/titleResource/getValueForUICulture` |
| `regionalSettings.web.descriptionResource` | `/_api/web/descriptionResource/getValueForUICulture` |
| `regionalSettings.list.titleResource` | `/_api/web/lists({listId:guid})/titleResource/getValueForUICulture` |
| `regionalSettings.list.descriptionResource` | `/_api/web/lists({listId:guid})/descriptionResource/getValueForUICulture` |
| `regionalSettings.field.web.titleResource` | `/_api/web/fields('{fieldId:raw}')/titleResource/getValueForUICulture` |
| `regionalSettings.field.web.descriptionResource` | `/_api/web/fields('{fieldId:raw}')/descriptionResource/getValueForUICulture` |
| `regionalSettings.field.list.titleResource` | `/_api/web/lists({listId:guid})/fields('{fieldId:raw}')/titleResource/getValueForUICulture` |
| `regionalSettings.field.list.descriptionResource` | `/_api/web/lists({listId:guid})/fields('{fieldId:raw}')/descriptionResource/getValueForUICulture` |
| `regionalSettings.contentType.web.titleResource` | `/_api/web/contentTypes('{contentTypeId:raw}')/nameResource/getValueForUICulture` |
| `regionalSettings.contentType.web.descriptionResource` | `/_api/web/contentTypes('{contentTypeId:raw}')/descriptionResource/getValueForUICulture` |
| `regionalSettings.contentType.list.titleResource` | `/_api/web/lists({listId:guid})/contentTypes('{contentTypeId:raw}')/nameResource/getValueForUICulture` |
| `regionalSettings.contentType.list.descriptionResource` | `/_api/web/lists({listId:guid})/contentTypes('{contentTypeId:raw}')/descriptionResource/getValueForUICulture` |
| `regionalSettings.customAction.site.titleResource` | `/_api/site/usercustomactions('{actionId:raw}')/titleResource/getValueForUICulture` |
| `regionalSettings.customAction.site.descriptionResource` | `/_api/site/usercustomactions('{actionId:raw}')/descriptionResource/getValueForUICulture` |
| `regionalSettings.customAction.web.titleResource` | `/_api/web/usercustomactions('{actionId:raw}')/titleResource/getValueForUICulture` |
| `regionalSettings.customAction.web.descriptionResource` | `/_api/web/usercustomactions('{actionId:raw}')/descriptionResource/getValueForUICulture` |
| `regionalSettings.customAction.list.titleResource` | `/_api/web/lists({listId:guid})/usercustomactions('{actionId:raw}')/titleResource/getValueForUICulture` |
| `regionalSettings.customAction.list.descriptionResource` | `/_api/web/lists({listId:guid})/usercustomactions('{actionId:raw}')/descriptionResource/getValueForUICulture` |

The PnPjs methods `getParentInfos`, `getById`, `getByTitle`, and specialized
field adders are also represented by their underlying collection/item routes
and payload controls; they do not each create a new server URL. Composite
helpers such as chunked uploads are represented by their component requests.

## Column-defaults composite helper routes

The installed `column-defaults/list.js` and `folder.js` methods perform several
requests rather than one REST operation. Their file-content read already maps
to `files.download`; the write path and conditional receiver registration were
missing and are now explicit:

| Endpoint ID | Method | Request path |
| --- | --- | --- |
| `files.addInFolder` | POST | `/_api/web/GetFolderByServerRelativePath(DecodedUrl=@folderPath)/Files/add(overwrite={overwrite:raw},url='{fileName:odata}')?@folderPath={folderPath:odataEncoded}` |
| `lists.eventReceivers.add` | POST | `/_api/web/lists({listId:guid})/EventReceivers/add` |

The second route is destructive because it registers a server-side event
receiver; both write routes are gated by `safety.allowDestructive`.

Custom action add, clear, update and delete operations are destructive and
remain disabled unless `safety.allowDestructive` is enabled.
