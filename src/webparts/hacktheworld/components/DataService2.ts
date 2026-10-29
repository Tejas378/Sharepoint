//        xhr.setRequestHeader("X-RequestDigest", localStorage.getItem('reqd')??"*");
/* eslint-disable */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-async-promise-executor */
/* eslint-disable @typescript-eslint/no-empty-function*/
/* eslint-disable max-lines */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable react/self-closing-comp */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @rushstack/security/no-unsafe-regexp */
/* eslint-disable max-lines */
//* eslint-disable no-useless-concat */
/* eslint-disable no-useless-escape */
/* eslint-disable @typescript-eslint/no-non-null-assertion */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable eqeqeq */
/* eslint-disable react/no-unescaped-entities */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable prefer-const */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable react/no-unescaped-entities */
import {  FieldTypes, PermissionKind, SPFI, spfi, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base"
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/folders";
import "@pnp/sp/folders/web";
import { createBatch } from "@pnp/sp/batching";
import "@pnp/sp/batching";
import "@pnp/sp/items/list";
import { random } from "@microsoft/sp-lodash-subset";
// import Constants from "./Constants";
let sp: SPFI;
let webObj: IWeb;
import {CONFIG}  from'./config'
const { lid,itemId,groupIds,groupId,groupName,principleIds,userLoginName,kinds,roles2,listName,LibraryName,viewName,subFolderName,SITE_URL,personGroup } = CONFIG; 
const HTTPMethods = ["GET", "POST", "PUT,", "PATCH", "MERGE", "DELETE"];
// const delay = (ms:number) => new Promise((resolve) => setTimeout(resolve, ms));
export default class DataService2 {
 
  public constructor(context: WebPartContext) {
    sp = spfi().using(SPFx(context))
    webObj = Web([sp.web, SITE_URL]);

  }
Test() {
    console.log("Test")
  }

delay = (ms:number) => new Promise(resolve => setTimeout(resolve, ms));
getShareLinkIds2 = async(): Promise<any[]> => {
  return new Promise((resolve, reject) => {
    let ShareIds: any[] = [];
    webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).getSharingInformation({"maxPrincipalsToReturn":100,"populateInheritedLinks":true,"clientSupportedFeatures":"share"}, ["permissionsInformation"], ["permissionsInformation","anyoneLinkTrackUsers"])
    .then((res) => {
        ShareIds = res.permissionsInformation.links.map((link: any) => link.linkDetails.ShareId);
        console.log(ShareIds);
        resolve(ShareIds); // Resolve the promise with ShareIds
      })
      .catch((error) => reject(error)); // Reject the promise in case of an error
  });
}
listInfo =async () =>{  

  try{
      this.sendRequest("POST",`https://tmsobb2.sharepoint.com/_api/web/Lists(@a1)/GetSharingInformation?@a1=%27${lid}%27&$Expand=pickerSettings,permissionsInformation`,JSON.stringify({"request":{"maxPrincipalsToReturn":100,"maxLinkMembersToReturn":100}}));
     
      await webObj.lists.getById(lid).descriptionResource("test").catch(first=>console.log(first));
      await webObj.lists.getById(lid).renderListData("<View><RowLimit>5000</RowLimit></View>").catch(first=>console.log(first));
      await webObj.lists.getById(lid).renderListDataAsStream({AddAllFields:true}).catch(first=>console.log(first));


      await webObj.lists.getById(lid).renderListFormData(itemId,'',1).catch(first=>console.log(first));
      await webObj.lists.getById(lid).renderListFormData(itemId,'',1).catch(first=>console.log(first));
      await webObj.lists.getById(lid).renderListFormData(itemId,'',1).catch(first=>console.log(first));


      await webObj.lists.getById(lid).getChanges({Add:true,Field:true,File:true}).catch(first=>console.log(first));
      await webObj.lists.getById(lid).getChanges({Add:true,Field:true,File:true}).catch(first=>console.log(first));
      await webObj.lists.getById(lid).getDefaultColumnValues().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getItemsByCAMLQuery({ViewXml:"<View><Query><Where><Contains><FieldRef Name='Title'/><Value Type='Text'>Test</Value></Contains></Where></Query></View>"}).catch(first=>console.log(first));
      await webObj.lists.getById(lid).getItemsByCAMLQuery({AllowIncrementalResults:true}).catch(first=>console.log(first));
      await webObj.lists.getById(lid).getListItemChangesSinceToken({ChangeToken:"*",RowLimit:"100"}).catch(first=>console.log(first));
      await webObj.lists.getById(lid).getListItemChangesSinceToken({RowLimit:"100"}).catch(first=>console.log(first));
      await webObj.lists.getById(lid).getListItemChangesSinceToken({}).catch(first=>console.log(first));
      await webObj.lists.getById(lid).getParentInfos().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getSiteScript().catch(first=>console.log(first));
      await webObj.lists.getById(lid).select("Title,ID")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).expand("Title")().catch(first=>console.log(first))
      await webObj.lists.getById(lid).eventReceivers().catch(first=>console.log(first))
      await webObj.lists.getById(lid).eventReceivers.select("Title")().catch(first=>console.log(first))
      await webObj.lists.getById(lid).eventReceivers.expand("Title")().catch(first=>console.log(first))
      await webObj.lists.getById(lid).eventReceivers.filter("Title eq Test")().catch(first=>console.log(first))
      await webObj.lists.getById(lid).eventReceivers.getContextInfo().catch(first=>console.log(first))
      await webObj.lists.getById(lid).eventReceivers.orderBy("Title")().catch(first=>console.log(first))
      await webObj.lists.getById(lid).firstUniqueAncestorSecurableObject().catch(first=>console.log(first))
      await webObj.lists.getById(lid).firstUniqueAncestorSecurableObject.expand()().catch(first=>console.log(first))
      await webObj.lists.getById(lid).firstUniqueAncestorSecurableObject.select("Title")().catch(first=>console.log(first))
      await webObj.lists.getById(lid).firstUniqueAncestorSecurableObject.getContextInfo("Title").catch(first=>console.log(first))
      await webObj.lists.getById(lid).informationRightsManagementSettings().catch(first=>console.log(first));
      await webObj.lists.getById(lid).informationRightsManagementSettings.select("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).informationRightsManagementSettings.expand("Title")().catch(first=>console.log(first));
      await this.sendRequest3("GET",`https://tmsobb2.sharepoint.com/_api/v2.0/sites/tmsobb2.sharepoint.com,bcb8ee81b-f387-41be-aa6f-a78488637407,27341d3a-24dc-48fc-ae4e-0a439cf859fd/lists/${lid}/Activities?%24expand=listItem(%24select%3Did%2Cname%2CwebUrl)&%24top=25`)
      await this.sendRequest("POST",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresiteassetslibrary/items(1)`,{}).catch(first=>console.log(first))
      await this.sendRequest("DELETE",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresiteassetslibrary/items(1)`,{}).catch(first=>console.log(first))
      await this.sendRequest("PUT",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresiteassetslibrary/items(1)`,JSON.stringify({"Title":"TTT"})).catch(first=>console.log(first))
      await this.sendRequest("PATCH",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresiteassetslibrary/items(1)`,JSON.stringify({"Title":"TTT"})).catch(first=>console.log(first))
      await this.sendRequest("MERGE",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresiteassetslibrary/items(1)`,JSON.stringify({"Title":"TTT"})).catch(first=>console.log(first))
      await this.sendRequest("POST",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresiteassetslibrary`,{}).catch(first=>console.log(first))
      await this.sendRequest("POST",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresitepageslibrary/items(1)`,{}).catch(first=>console.log(first))
      await this.sendRequest("PUT",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresitepageslibrary/items(1)`,{"Title":"TTT"}).catch(first=>console.log(first))
      await this.sendRequest("PATCH",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresitepageslibrary/items(1)`,{"Title":"TTT"}).catch(first=>console.log(first))
      await this.sendRequest("MERGE",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresitepageslibrary/items(1)`,{"Title":"TTT"}).catch(first=>console.log(first))
      await this.sendRequest("DELETE",`https://tmsobb2.sharepoint.com/_api/web/lists/ensuresitepageslibrary`,{}).catch(first=>console.log(first))
      await webObj.lists.getById(lid).update({Title:'TestTitle'},"*").catch(first=>console.log(first));
      await webObj.lists.getById(lid).subscriptions().catch(first=>console.log(first));
      await webObj.lists.getById(lid).rootFolder().catch(first=>console.log(first));
      await webObj.lists.getById(lid).relatedFields().catch(first=>console.log(first));
      await webObj.lists.getById(lid).reserveListItemId().catch(first=>console.log(first));
      await webObj.lists.getById(lid).setDefaultColumnValues([]).catch(first=>console.log(first));
}
  catch(first)
  {
    console.log(first)
  }
  }
ListCrudContentType= async() =>{

    try{
      let contenttype:any='';
      let contenttypebody:any='';
      await  webObj.lists.getById(lid).contentTypes().then((ctypes:any)=>{ contenttype=ctypes[0].StringId;}).catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype)().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype)().then(body=>{contenttypebody=body;console.log(body)}).catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype).fields().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype).fields.getContextInfo().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype).fields.expand("Title")().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype).fieldLinks().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype).fieldLinks.getById('AA')().catch(first=>console.log(first));;
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype).fieldLinks.expand('Title')().catch(first=>console.log(first));;
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype).fieldLinks.getContextInfo('Title').catch(first=>console.log(first));;
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype).descriptionResource('mm').catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.add("TestAddedContentType","Test0","AAA").catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.addAvailableContentType(contenttype).catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype).parent().catch(first=>console.log(first));;
      await  webObj.lists.getById(lid).contentTypes.getById(contenttype).update({Description:"sss"}).catch(first=>console.log(first));
      // await  webObj.lists.getById(lid).contentTypes.getById(contenttype).delete().catch(first=>console.log(first))
      await  webObj.lists.getById(lid).contentTypes.orderBy("Title",true)().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.filter("Id eq item")().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).contentTypes.select("Id eq item")().catch(first=>console.log(first));
     HTTPMethods.forEach((req:any)=>{
      this.sendRequest3(req,`https://tmsobb2.sharepoint.com/_api/web/Lists(guid'${lid}')/ContentTypes('${contenttype}')`,JSON.stringify(contenttypebody))

     })
  
    }
    catch(first)
    {
      console.log(first)
    }
  
  }
ListColumnAddUpdateDelete= async () => { //Tested Working
  try{
  let fieldsID:string='d80fa84b-02e7-430d-817f-66e88ebef823';
  await webObj.lists.getById(lid).fields.getByTitle('Title')().then(res=>res).then(res=>{fieldsID=res.Id;console.log(fieldsID)}).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.getByTitle('Title')().catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.getById(fieldsID)().catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.getByInternalNameOrTitle('Title')().catch(first=>console.log(first));;;
  await webObj.lists.getById(lid).fields.getByInternalNameOrTitle('ID')().catch(first=>console.log(first));;;
  await webObj.lists.getById(lid).fields.filter("Title eq Title")().catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.select("Title")().catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.expand("Title")().catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.expand()().catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.top(1)().catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.orderBy("Title")().catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.add("My Field1", FieldTypes.Text, { FieldTypeKind: 3, Group: "My Group" }).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.addText("My Field2", { MaxLength: 255, Group: "My Group" }).catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.addLookup("My Lookup_two", {LookupListId: '6381a604-2666-4571-9da7-9aa8b92a935e', LookupFieldName: "Title"}).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.addBoolean("Temp Field", { Group: "My Group" }).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.addField("Lookup-test",7,{LookupListId: '6381a604-2666-4571-9da7-9aa8b92a935e', LookupFieldName: "Title"}).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.addField("test",2).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.getByTitle("Title").select("Title")().catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.getByTitle("Title").expand("Title")().catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.skip(2).add("My Field1ddfg", FieldTypes.Text, { FieldTypeKind: 3, Group: "My Group" }).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.getByTitle("Title").update({ Description: "My Description" }).catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.getByTitle("Title").update({ ValidationMessage: "My Description" }).catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.getByTitle("Title").setShowInDisplayForm(true).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.getByTitle("Title").setShowInEditForm(false).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.getByTitle("Title").setShowInNewForm(false).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.addDependentLookupField('sclookup2','04c883c0-9465-4167-a53e-e1ede03e7d66','MyField').catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.addDependentLookupField('sclookup2','04c883c0-9465-4167-a53e-e1ede03e7d66','MyField').catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.getByTitle("My Field1").update({DefaultValue:null,EnforceUniqueValues:false,Required:false,ReadOnlyField:false},"SP.Field").catch(first=>console.log(first));;
  await webObj.lists.getById(lid).fields.getByTitle("My Field1").update({DefaultValue:"xxxxxxxxxx",EnforceUniqueValues:false,Required:false,ReadOnlyField:true},"SP.Field").catch(first=>console.log(first));; 
  await webObj.lists.getById(lid).fields.getByTitle("My Field1").update({Description:"Testtttttt",Sealed:true},"SP.Field").catch(first=>console.log(first));; 
  await this.sendRequest("MERGE",`https://tmsobb2.sharepoint.com/_api/web/GetList(@a1)/Fields/GetByInternalNameOrTitle(@a2)?@a1=%27%2FLists%2F${listName}%27&@a2=%27test%27`,JSON.stringify({"__metadata":{"type":"SP.Field"},"SchemaXml":"<Field ClientSideComponentId='00000000-0000-0000-0000-000000000000' DisplayName='test' Format='Dropdown' IsModern='TRUE' MaxLength='255' Name='test' Title='test' Type='Text'><Default>aaaaa</Default></Field>","Title":"test","Description":""}));
  const xmlSchema = `
  <Field 
      Name="CustomField" 
      DisplayName="Custom Field" 
      Type="Text" 
      Required="FALSE" 
      MaxLength="255"
      Group="Custom Columns" 
      StaticName="CustomField"
  />
  `;   
  const fieldCreationInfo = {
  SchemaXml: xmlSchema,
  Options: 2 // AddFieldOptions (Optional: specify how the field is added)
  };
  await webObj.lists.getById(lid).fields.createFieldAsXml(fieldCreationInfo).catch(first=>console.log(first));
  await webObj.lists.getById(lid).fields.getByTitle("test").delete().catch(first=>console.log(first));

this.sendRequest2("DELETE",`https://tmsobb2.sharepoint.com/_api/web/GetList(@a1)/fields('fa564e0f-0c70-4ab9-b863-0177e6ddd247')?@a1=%27%2FLists%2F${listName}%27`)

this.sendRequest2("POST",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/fields/adddependentlookupfield(displayname='Project',primarylookupfieldid='a3de9a09-1800-4cbc-b89a-058c719dc44e',showfield='Title')`)?.catch(first=>console.log(first))
this.sendRequest2("POST",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/fields/createfieldasxml`,JSON.stringify(`{ 'parameters': {     'SchemaXml': '<Field Type="UserMulti" Title="Stakeholders" DisplayName="Stakeholders"    Required="FALSE" UserSelectionMode="PeopleAndGroups" UserSelectionScope="0" Mult="TRUE" />' } }`))?.catch(first=>console.log(first))




  // await webObj.lists.getById(lid).fields.getByTitle("kkkkkk").delete().catch(first=>console.log(first));
}
    catch(first)
    {
      console.log(first)
    }
  }
ListCrudviewsandgetviews= async () => { //Tested Working
    try{
        const newView = {
          Title: "TestFormView", // Required: the title of the view
          RowLimit: 30, // Required: the row limit for view pagination
          ViewType: "Grid", // Required: specify the type of view, e.g., "Grid"
          ViewQuery: "<Query><Where><IsNotNull><FieldRef Name='Title' /></IsNotNull></Where></Query>", // Example CAML query to filter items
          AssociatedContentTypeId: "0x0101", // Optional: example content type ID
          CalendarViewStyles: "<CalendarViewStyles />", // Optional: example XML string
          CustomFormatter: "<CustomFormatter />", // Optional: custom JSON or XML format string
          DefaultView: false,
          DefaultViewForContentType: true,
          EditorModified: false,
          Formats: "<Formats />", // Optional: format string
          Hidden: false,
          HtmlSchemaXml: "<HtmlSchemaXml />", // Optional: example XML schema
          Id: "12345abcde", // Optional: example GUID
          ImageUrl: "/sites/demo/SiteAssets/view-icon.png", // Optional: relative URL to an image
          IncludeRootFolder: true,
          JSLink: "/_layouts/15/clienttemplates.js", // Optional: path to JavaScript file for client templates
          ListViewXml: "<ListViewXml />", // Optional: view XML configuration
          Method: "GET", // Optional: HTTP method, if applicable
          MobileDefaultView: false,
          MobileView: true,
          ModerationType: "Moderated", // Optional: moderation type
          NewDocumentTemplates: "<NewDocumentTemplates />", // Optional: document template configuration
          OrderedView: true,
          Paged: true,
          PersonalView: false,
          ReadOnlyView: false,
          RequiresClientIntegration: true,
          Scope:1, // Optional: specify view scope
          ServerRelativePath: {
              DecodedUrl: "/sites/demo/Lists/ExampleList",
          },
          ServerRelativeUrl: "/sites/demo/Lists/ExampleList/ViewName",
          StyleId: "100", // Optional: style ID
          TabularView: true,
          Threaded: false,
          Toolbar: "<Toolbar />", // Optional: toolbar configuration XML
          ToolbarTemplateName: "CustomToolbarTemplate", // Optional: toolbar template name
          ViewData: "<ViewData />", // Optional: data configuration
          ViewJoins: "<ViewJoins />", // Optional: joins XML
          ViewProjectedFields: {
              SchemaXml: "<ViewProjectedFields />",
          },
          ViewType2: "COMPACTLIST", // Optional: one of "KANBAN" | "TILES" | "COMPACTLIST" | "MODERNCALENDAR" or null
          VisualizationInfo: null, // Optional: additional visualization settings
      }; 
      let viewId:string = '';
      let vids:string[]=[];
      await webObj.lists.getById(lid).views().then(views=>views).then(view=>{view.forEach(viii=>viii.Title!=="AllItems"? vids.push(viii.Id):'');console.log(vids)}).catch(first=>console.log(first));
      let body2:any =`{"__metadata":{"type":"SP.View"},"Title":"${viewName}","CustomFormatter":${JSON.stringify('<![CDATA[{"header":{"title":"TestForm"},"formId":"${fid}","listId":"{${lid}}"}]]>')},"Hidden":true}`

      await this.sendRequest2("PATCH",`https://tmsobb2.sharepoint.com/_api/web/GetList(@a1)/Views(@a2)?@a1=%27%2FLists%2F${listName}%27&@a2=%27${vids[0]}%27`,body2)?.catch(er=>console.log(er));
      // const newView2 = {"parameters":{"__metadata":{"type":"SP.ViewCreationInformation"},"Title":"Test Form","ViewFields":{"__metadata":{"type":"Collection(Edm.String)"},"results":[]},"Query":"","Paged":true,"PersonalView":false,"RowLimit":30,"ViewTypeKind":1,"ViewType2":"FORMS"}}   
      await webObj.lists.getById(lid).views().catch(first=>console.log(first))
      await webObj.lists.getById(lid).views.add("Test",true).then((resp:any)=>resp).then(res=>{viewId=res.Id;console.log(viewId)}).catch(first=>console.log(first))
      await webObj.lists.getById(lid).views.add("Test123",true,{...newView}).then((resp:any)=>resp).then(res=>{viewId=res.Id;console.log(viewId)}).catch(first=>console.log(first))
      await webObj.lists.getById(lid).views.getByTitle(viewName)().then((resp:any)=>resp).then(res=>{viewId=res.Id;console.log(viewId)}).catch(first=>console.log(first))
      await webObj.lists.getById(lid).views.getByTitle(viewName)().catch(first=>console.log(first));;
      await webObj.lists.getById(lid).views.getByTitle(viewName).select("Title")().catch(first=>console.log(first));;
      await webObj.lists.getById(lid).views.getByTitle(viewName).renderAsHtml().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).update({"Title":"Test"}).catch(first=>console.log(first)); //update view name
      await webObj.lists.getById(lid).views.getByTitle(viewName).setViewXml("<View><RowLimit>5000</RowLimit></View>").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).expand("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).fields.move("ID",3).catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).fields.add("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).fields.add("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).fields.orderBy("Title",true)().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).fields.select("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).fields.filter("Title eq Test")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).fields.remove("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).fields.removeAll().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).fields.getSchemaXml().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).fields.top(1)().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getByTitle(viewName).delete().catch(first=>console.log(first));
      viewId = viewId==""?vids.length>0?vids[0]:viewId:viewId;
      await webObj.lists.getById(lid).views.getById(viewId)().then((resp:any)=>resp).then(res=>{viewId=res.Id;console.log(viewId)}).catch(first=>console.log(first))
      await webObj.lists.getById(lid).views.getById(viewId)().catch(first=>console.log(first));;
      await webObj.lists.getById(lid).views.getById(viewId).select("Title")().catch(first=>console.log(first));;
      await webObj.lists.getById(lid).views.getById(viewId).renderAsHtml().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).update({"Title":"Test"}).catch(first=>console.log(first)); //update view name
      await webObj.lists.getById(lid).views.getById(viewId).setViewXml("<View><RowLimit>5000</RowLimit></View>").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).expand("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.move("ID",3).catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.add("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.add("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.orderBy("Title",true)().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.select("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.filter("Title eq Test")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.remove("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.removeAll().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.getSchemaXml().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.top(1)().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).delete().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).expand("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).select("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).getContextInfo().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).renderAsHtml().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).update({"Title":"Test"}).catch(first=>console.log(first));;
      await webObj.lists.getById(lid).getView(viewId).setViewXml("<View><RowLimit>5000</RowLimit></View>").catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.getSchemaXml().catch(first=>console.log(first));;
      await webObj.lists.getById(lid).getView(viewId).fields.select("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.expand("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.expand()().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.filter("Title eq Test")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.move("ID",0).catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.orderBy("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.add("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.remove("Title").catch(first=>console.log(first));;
      await webObj.lists.getById(lid).getView(viewId).fields.remove("ID").catch(first=>console.log(first));;
      await webObj.lists.getById(lid).getView(viewId).fields.removeAll().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).delete().catch(first=>console.log(first));
      await vids.forEach(vid=>{webObj.lists.getById(lid).getView(vid).delete().catch(first=>console.log(first)); })
      this.sendRequest("GET",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)
      this.sendRequest("POST",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)
      this.sendRequest("PUT",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)
      this.sendRequest("PATCH",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)
      this.sendRequest("MERGE",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)
      this.sendRequest("DELETE",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)

      // await webObj.lists.getById(lid).views.add("Test",false).catch(first=>console.log(first));

    }
    catch(first)
    {
      console.log(first)
    }
  }
ListCrudFormviewsandgetviews= async () => { //Tested Working
    try{
    
      let viewId:string = '';
      let vids:string[]=[];
      await webObj.lists.getById(lid).views().then(views=>views).then(view=>{view.forEach((viii:any)=>viii.ViewType2=="FORMS"? vids.push(viii.Id):'');console.log(vids)}).catch(first=>console.log(first));
      viewId = vids.length>0?vids[0]:viewId;
      await webObj.lists.getById(lid).views.getById(viewId)().then((resp:any)=>resp).then(res=>{viewId=res.Id;console.log(viewId)}).catch(first=>console.log(first))
      await webObj.lists.getById(lid).views.getById(viewId)().catch(first=>console.log(first));;
      await webObj.lists.getById(lid).views.getById(viewId).select("Title")().catch(first=>console.log(first));;
      await webObj.lists.getById(lid).views.getById(viewId).renderAsHtml().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).update({"Title":"Test"}).catch(first=>console.log(first)); //update view name
      await webObj.lists.getById(lid).views.getById(viewId).setViewXml("<View><RowLimit>5000</RowLimit></View>").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).expand("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.move("ID",3).catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.add("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.add("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.orderBy("Title",true)().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.select("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.filter("Title eq Test")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.remove("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.removeAll().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.getSchemaXml().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).fields.top(1)().catch(first=>console.log(first));
      await webObj.lists.getById(lid).views.getById(viewId).delete().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).expand("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).select("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).getContextInfo().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).renderAsHtml().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).update({"Title":"Test"}).catch(first=>console.log(first));;
      await webObj.lists.getById(lid).getView(viewId).setViewXml("<View><RowLimit>5000</RowLimit></View>").catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.getSchemaXml().catch(first=>console.log(first));;
      await webObj.lists.getById(lid).getView(viewId).fields.select("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.expand("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.expand()().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.filter("Title eq Test")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.move("ID",0).catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.orderBy("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.add("Title").catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).fields.remove("Title").catch(first=>console.log(first));;
      await webObj.lists.getById(lid).getView(viewId).fields.remove("ID").catch(first=>console.log(first));;
      await webObj.lists.getById(lid).getView(viewId).fields.removeAll().catch(first=>console.log(first));
      await webObj.lists.getById(lid).getView(viewId).delete().catch(first=>console.log(first));
      this.sendRequest("GET",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)?.catch(first=>console.log(first))   
      this.sendRequest("POST",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)?.catch(first=>console.log(first))   
      this.sendRequest("PUT",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)?.catch(first=>console.log(first))   
      this.sendRequest("PATCH",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)?.catch(first=>console.log(first))   
      this.sendRequest("MERGE",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)?.catch(first=>console.log(first))   
      this.sendRequest("DELETE",`https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/views('${viewId}')/HtmlSchemaXml`)?.catch(first=>console.log(first))   
      await webObj.lists.getById(lid).views.add("Test",false).catch(first=>console.log(first));

    }
    catch(first)
    {
      console.log(first)
    }
  }
ListCruddefaultView= async(contenttype:string="1") =>{ //Tested Working

    try{
      await  webObj.lists.getById(lid).defaultView().catch(first=>{console.log(first)})
      await  webObj.lists.getById(lid).defaultView.renderAsHtml().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.select("ID")().catch(first=>console.log(first));;
      await  webObj.lists.getById(lid).defaultView.select("Id")().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.expand()().catch(first=>console.log(first));;
      await  webObj.lists.getById(lid).defaultView.expand("Title")().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.update({Title:"Test111",ReadOnlyView:true,DefaultView:true}).catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.update({ RowLimit: 20}).catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.setViewXml('<View><RowLimit>5000</RowLimit></View>').catch(first=>console.log(first))
      await  webObj.lists.getById(lid).defaultView.fields().catch(first=>console.log(first))
      await  webObj.lists.getById(lid).defaultView.fields.expand("Title")().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.fields.filter("Title eq test")().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.fields.select("Id")().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.fields.add("Title").catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.fields.add("Modified").catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.fields.getContextInfo().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.fields.getSchemaXml().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.fields.move("Title",0).catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.fields.orderBy("Title")().catch(first=>console.log(first));; 
      await  webObj.lists.getById(lid).defaultView.fields.removeAll().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.fields.getSchemaXml().catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.fields.remove("Title").catch(first=>console.log(first));
      await  webObj.lists.getById(lid).defaultView.delete().catch(first=>console.log(first));
    }
    catch(first)
    {
      console.log(first)
    }
  
  }
ListCrudForm= async(contenttype:string="1") =>{ //Tested Working
    try{
      // let viewId:string='';
      const newView2 = {"parameters":{"__metadata":{"type":"SP.ViewCreationInformation"},"Title":"Test Form","ViewFields":{"__metadata":{"type":"Collection(Edm.String)"},"results":[]},"Query":"","Paged":true,"PersonalView":false,"RowLimit":30,"ViewTypeKind":1,"ViewType2":"FORMS"}}   
      let vids:string[]=[];
      await webObj.lists.getById(lid).views().then(views=>views).then(view=>{view.forEach((viii:any)=>viii.ViewType2=="FORMS"? vids.push(viii.Id):'');console.log(vids)}).catch(first=>console.log(first));
      let fid:any=vids.length>0?vids[0]:'7e57806a-f589-47c1-8ff8-7ad3632030a5'//vids[0]??'';
      await this.sendRequest("POST",`https://tmsobb2.sharepoint.com/_api/web/GetList(@a1)/Views/Add?@a1=%27%2FLists%2F${listName}%27`,JSON.stringify(newView2)).then((res:any)=>res).then((res:any)=>{console.log(res.d.Id);console.log(fid)}).catch((first:any)=>console.log(first))
      let body1:any =`{"__metadata":{"type":"SP.View"},"Title":${viewName},"CustomFormatter":${JSON.stringify('<![CDATA[{"header":{"title":"TestForm"},"formId":"${fid}","listId":"{${lid}}"}]]>')},"Hidden":true}`
      let body2:any =`{"__metadata":{"type":"SP.View"},"Title":${viewName},"Hidden":false}`
      await this.sendRequest2("PATCH",`https://tmsobb2.sharepoint.com/_api/web/GetList(@a1)/Views(@a2)?@a1=%27%2FLists%2F${listName}%27&@a2=%27${fid}%27`,body1)?.catch(first=>console.log(first))   
      await this.sendRequest2("PATCH",`https://tmsobb2.sharepoint.com/_api/web/GetList(@a1)/Views(@a2)?@a1=%27%2FLists%2F${listName}%27&@a2=%27${fid}%27`,body2)?.catch(first=>console.log(first))   
      await webObj.lists.getById(lid).forms().catch(first=>console.log(first));
      await webObj.lists.getById(lid).forms.getById(fid)().catch(first=>console.log(first));
      await webObj.lists.getById(lid).forms.getById(fid).expand("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).forms.getById(fid).select("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).forms.filter(`Title eq ${viewName}`)().catch(first=>console.log(first));
      await webObj.lists.getById(lid).forms.expand("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).forms.select("Title")().catch(first=>console.log(first));
      await webObj.lists.getById(lid).forms.getContextInfo(`/Lists/${listName}`).catch(first=>console.log(first));
      await webObj.lists.getById(lid).forms.orderBy("TimeCreated")().catch(first=>console.log(first));
      await this.sendRequest2("POST",`https://tmsobb2.sharepoint.com/_api/web/lists/getByTitle('${listName}')/forms/getByPageType`,JSON.stringify({"formType":4}))?.catch(first=>console.log(first))   
      await this.sendRequest2("GET",`https://tmsobb2.sharepoint.com/_api/web/lists/getByTitle('${listName}')/forms/getById('${fid}')`)?.catch(first=>console.log(first))   
      await this.sendRequest2("GET",`https://tmsobb2.sharepoint.com/_api/web/lists/getByTitle('${listName}')/forms('${fid}')/id`)?.catch(first=>console.log(first))   
      await this.sendRequest2("GET",`https://tmsobb2.sharepoint.com/_api/web/lists/getByTitle('${listName}')/forms('${fid}')/formType`)?.catch(first=>console.log(first))   
      await this.sendRequest2("POST",`https://tmsobb2.sharepoint.com/_api/web/lists/getByTitle('${listName}')/forms/getById('${fid}')`)?.catch(first=>console.log(first))   
      await this.sendRequest2("POST",`https://tmsobb2.sharepoint.com/_api/web/lists/getByTitle('${listName}')/forms('${fid}')/id`)?.catch(first=>console.log(first))   
      await this.sendRequest2("POST",`https://tmsobb2.sharepoint.com/_api/web/lists/getByTitle('${listName}')/forms('${fid}')/formType`)?.catch(first=>console.log(first))   
      let roles = [0,1,22,3,4,5,6,7,8,9]
      kinds.forEach(kid=>{
        roles.forEach((rid)=>{
        var xhr = new XMLHttpRequest();
        xhr.open("POST", `https://tmsobb2.sharepoint.com/_api/web/Lists(@a1)/ShareLink?@a1=%27%7b${lid}%7D%27`, true);
        xhr.setRequestHeader("content-type", "application\/json;odata=verbose");
        xhr.setRequestHeader("accept", "application\/json; odata=verbose");
        xhr.setRequestHeader("X-RequestDigest", localStorage.getItem('reqd')??"*");
        xhr.setRequestHeader("accept-language", "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7");
        xhr.withCredentials = true;
        var body = JSON.stringify({"request":{"createLink":true,"settings":{"linkKind":kid,"allowAnonymousAccess":false,"role":rid,"expiration":null,"restrictShareMembership":false,"nav": btoa(fid) }}})
                var aBody = new Uint8Array(body.length);
        for (var i = 0; i < aBody.length; i++)
          aBody[i] = body.charCodeAt(i); 
        xhr.send(new Blob([aBody]));
       })
     });

     await this.ListCrudFormviewsandgetviews();

    }
    catch(first)
    {
      console.log(first)
    }
  
  }
ListLibraryFileFileversions = async () => {  //Tested Working
  try {
    let versionIds:any =[]; 
    let versionURLS:any =[]; 
    await webObj.lists.getByTitle('Test').rootFolder.files.getByUrl("Attackerfile.txt").versions().then( async (versions: any) => {versions.forEach((version: any) =>{versionIds.push(version.ID); versionURLS.push(version.Url)});console.log(versionIds);}).catch(first=>console.log(first)); 

      let vid1:any = versionIds[0];
      versionURLS.forEach((versionurl:string)=>{
        this.sendRequest2("GET",`https://tmsobb2.sharepoint.com/${versionurl}`)
      });
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1);
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1)().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).select("ID")().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.orderBy('ID',true)().catch(first=>console.log(first));;
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.select('ID')().catch(first=>console.log(first));;
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.orderBy("ID",true)().catch(first=>console.log(first));;
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.filter(`ID eq ${vid1}`)().catch(first=>console.log(first));;
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getBlob().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getBuffer().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getJSON().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getStream().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getText().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getContextInfo().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.recycleByID(vid1).catch(first=>console.log(first))
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.recycleByLabel(`${vid1}`).catch(first=>console.log(first))
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.deleteById(vid1).catch(first=>console.log(first))
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.deleteByLabel(`${vid1}`).catch(first=>console.log(first))
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.restoreByLabel(`${vid1}`).catch(first=>console.log(first))
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).delete().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.skip(3).deleteAll().catch(first=>console.log(first))
  vid1 = versionIds[0]+1;
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1);
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getBlob().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getBuffer().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getJSON().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getStream().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getText().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).getContextInfo().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.orderBy('ID',true)();
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.select('ID')();
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.recycleByID(vid1).catch(first=>console.log(first))
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.recycleByLabel(`${vid1}`).catch(first=>console.log(first))
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.deleteById(vid1).catch(first=>console.log(first))
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.deleteByLabel(`${vid1}`).catch(first=>console.log(first))
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.restoreByLabel(`${vid1}`).catch(first=>console.log(first))
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.getById(vid1).delete().catch(first=>console.log(first));
  await  webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").versions.skip(1).deleteAll().catch(first=>console.log(first))
    
  }
  catch (error) 
  {
      console.log(error);
  }
}; 
ListCrudrootFolderFiles = async (fileInput:any) => {//Tested Working
  try {
      const file =await fileInput.files[0];
      const fileUrl =  `${file.name}`; // Relative URL for the file in the SharePoint list
      const chunkSize = 10485760; // Chunk size (10 MB in this example)
      const props = {
       progress: () => { console.log(`progress`); },
          ChunkSize: chunkSize,  // Optional: controls the chunk size
          Overwrite: true,       // Optional: controls whether to overwrite the file if it exists
      };
      const props2 = {
        Overwrite: true,  // Optional: controls the chunk size
        AutoCheckoutOnInvalidData: false,   
        EnsureUniqueFileName:false
            // Optional: controls whether to overwrite the file if it exists
      };
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addTemplateFile('/Test/xyz.txt',1).catch(first=>console.log(first));;
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addTemplateFile('/xywz.txt',1).catch(first=>console.log(first));;
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addChunked(fileUrl, file, props2).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addChunked("new.txt",file , { progress: data => { console.log(`progress`); }, Overwrite: true }).catch(first=>console.log(first));;
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addChunked("new.txt",file , { progress: data => { console.log(`progress`); }, Overwrite: true }).catch(first=>console.log(first));;
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addUsingPath(fileUrl, file, props2).catch(first=>console.log(first));  
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").approve("TestApprove").catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").cancelUpload("b531397f-24f9-40ed-a983-a42f03e17927").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").checkSharingPermissions([{email:"tmsobb@tmsobb2.onmicrosoft.com"}]).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").checkout().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").checkin("Test").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").checkin("Test",1).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").checkin("Test",2).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").checkin("Test",0).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").deleteSharingLinkByKind(1).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").deny("Denied").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").exists().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getBlob().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getBuffer().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getItem().catch(first=>console.log(first));
     await (await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getItem()).comments.add("Test").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).items.getItemByStringId(`'${itemId}`).comments.add("Test").catch(first=>console.log(first));
     await (await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getItem()).update({Title:"sssss"}).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getItem("Title", "Modified").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getItem<{ Id: number, Title: string }>("Id", "Title").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getJSON().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getLimitedWebPartManager();
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getLockedByUser().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getObjectSharingSettings().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getShareLink(1).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getSharingInformation().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getStream().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").getText().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").listItemAllFields().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").listItemAllFields.expand("Title")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").listItemAllFields.select("Title")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").listItemAllFields.getContextInfo().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").copyByPath("https://tmsobb2.sharepoint.com/Movec/ssss.txt",true,{KeepBoth:true,ShouldBypassSharedLocks:true,ResetAuthorAndCreatedOnCopy:true}).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").copyTo("/Movec/sss.txt",true).catch(first=>console.log(first)).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").copyByPath("https://tmsobb2.sharepoint.com/Movec/ssss.txt",true,{KeepBoth:true,ShouldBypassSharedLocks:true,ResetAuthorAndCreatedOnCopy:true}).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").moveByPath('https://tmsobb2.sharepoint.com/Movec/sss2222.txt',false,true).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").moveByPath('https://tmsobb2.sharepoint.com/Movec/sss2222.txt',true,true).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").moveByPath('https://tmsobb2.sharepoint.com/Movec/egret.txt',true,true).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addUsingPath(fileUrl, file, props2).catch(first=>console.log(first)); 
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").publish().catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").publish("Published").catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").unpublish().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").unpublish("AAAA").catch(first=>console.log(first)); 
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").setContent("ssss").catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").setContentChunked(file,{progress:data=>console.log("starting")})
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").checkout().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").undoCheckout().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").delete().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").deleteWithParams({BypassSharedLock:true}).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").deleteWithParams({BypassSharedLock:true}).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.getByUrl("Attackerfile.txt").recycle().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.orderBy("Title")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.select("Title")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.files.filter("Title eq sssss")().catch(first=>console.log(first));
  
  } catch (error) {
      console.error("Error uploading file:", error);
  }
};
ListCrudrootFolderSubFolderFiles = async (fileInput:any) => {//Tested Working
  try {
      const file =await fileInput.files[0];
      const fileUrl =  `${file.name}`; // Relative URL for the file in the SharePoint list
      const chunkSize = 10485760; // Chunk size (10 MB in this example)
      const props = {
       progress: () => { console.log(`progress`); },
          ChunkSize: chunkSize,  // Optional: controls the chunk size
          Overwrite: true,       // Optional: controls whether to overwrite the file if it exists
      };
      const props2 = {
        Overwrite: true,  // Optional: controls the chunk size
        AutoCheckoutOnInvalidData: false,   
        EnsureUniqueFileName:false
            // Optional: controls whether to overwrite the file if it exists
      };   
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addUsingPath(fileUrl, file, props).catch(first=>console.log(first));  
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addUsingPath(fileUrl, file, props2).catch(first=>console.log(first));  
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addTemplateFile('/Test/xyz.txt',1).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addTemplateFile('/xywz.txt',1).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addChunked(fileUrl, file, props2).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addChunked("new.txt",file , { progress: data => { console.log(`progress`); }, Overwrite: true }).catch(first=>console.log(first));;
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addChunked("new.txt",file , { progress: data => { console.log(`progress`); }, Overwrite: true }).catch(first=>console.log(first));;
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addUsingPath(fileUrl, file, props2).catch(first=>console.log(first));  
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").approve("TestApprove").catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").cancelUpload("b531397f-24f9-40ed-a983-a42f03e17927").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").checkSharingPermissions([{email:"tmsobb@tmsobb2.onmicrosoft.com"}]).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").checkout().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").checkin("Test").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").checkin("Test",1).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").checkin("Test",2).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").checkin("Test",0).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").deleteSharingLinkByKind(1).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").deny("Denied").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").exists().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getBlob().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getBuffer().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getItem().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getItem("Title", "Modified").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getItem<{ Id: number, Title: string }>("Id", "Title").catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getJSON().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getLimitedWebPartManager();
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getLockedByUser().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getObjectSharingSettings().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getShareLink(1).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getSharingInformation().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getStream().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getText().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").listItemAllFields().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").listItemAllFields.expand("Title")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").listItemAllFields.select("Title")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").listItemAllFields.getContextInfo().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").copyByPath("https://tmsobb2.sharepoint.com/Movec/ssss.txt",true,{KeepBoth:true,ShouldBypassSharedLocks:true,ResetAuthorAndCreatedOnCopy:true}).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").copyTo("/Movec/sss.txt",true).catch(first=>console.log(first)).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").copyByPath("https://tmsobb2.sharepoint.com/Movec/ssss.txt",true,{KeepBoth:true,ShouldBypassSharedLocks:true,ResetAuthorAndCreatedOnCopy:true}).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").moveByPath('https://tmsobb2.sharepoint.com/Movec/sss2222.txt',false,true).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").moveByPath('https://tmsobb2.sharepoint.com/Movec/sss2222.txt',true,true).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").moveByPath('https://tmsobb2.sharepoint.com/Movec/egret.txt',true,true).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addUsingPath(fileUrl, file, props2).catch(first=>console.log(first)); 
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt"). publish().catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt"). publish("Published").catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").unpublish().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").unpublish("AAAA").catch(first=>console.log(first)); 
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").setContent("ssss").catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").setContentChunked(file,{progress:data=>console.log("starting")})
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").checkout().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").undoCheckout().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").delete().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").deleteWithParams({BypassSharedLock:true}).catch(first=>console.log(first))
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").deleteWithParams({BypassSharedLock:false}).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").recycle().catch(first=>console.log(first));
     
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.addChunked(fileUrl, file, props).catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.orderBy("Title")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.select("Title")().catch(first=>console.log(first));
     await webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.filter("Title eq sssss")().catch(first=>console.log(first));

  } catch (error) {
      console.error("Error uploading file:", error);
  }
};
ListCrudrootFolderShare= async() =>{

  try{
    await  webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).getSharingInformation({"maxPrincipalsToReturn":100,"populateInheritedLinks":true,"clientSupportedFeatures":"share"}, ["permissionsInformation"], ["permissionsInformation","anyoneLinkTrackUsers"]).catch(first=>console.log(first));
    await  webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getSharingInformation({"maxPrincipalsToReturn":100,"populateInheritedLinks":true,"clientSupportedFeatures":"share"}, ["permissionsInformation"], ["permissionsInformation","anyoneLinkTrackUsers"]).catch(first=>console.log(first));
    await  kinds.map(async(kid:any)=>{
    await  webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).getShareLink(kid).catch(first=>console.log(first));
    await  webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").getShareLink(kid).catch(first=>console.log(first));
      })
    for (const role of roles2) {
      for (const person of personGroup) {
        try {
          await  webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).shareWith(person, role, true, {
            subject: "dddd",
            body: "dddd"
          }).catch(first=>console.log(first));
          await  webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").shareWith(person, role, true, {
            subject: "dddd",
            body: "dddd"
          }).catch(first=>console.log(first));
        } catch (error) {
          console.log(error);
        }
      }
    }
    for (const role of roles2) {
      for (const person of personGroup) {
        try {
          await  webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).shareWith(person, role, false, {
            subject: "dddd",
            body: "dddd"
          }).catch(first=>console.log(first));
          await  webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).files.getByUrl("Attackerfile.txt").shareWith(person, role, false, {
            subject: "dddd",
            body: "dddd"
          }).catch(first=>console.log(first));
        } catch (error) {
          console.log(error);
        }
      }
    }


    // let shareId:any[]= await this.getShareLinkIds2(itemId);
    // await  kinds.map(async(kid:any)=>{
    //   await shareId.forEach(async (sid:any)=>{
    //     await  webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).unshareLink(kid,sid).catch(first=>console.log(first));
    //   })
    // })
    // await  kinds.map(async(kid:any)=>{
    //   await  webObj.lists.getByTitle(LibraryName).rootFolder.folders.getByUrl(subFolderName).deleteSharingLinkByKind(kid).catch(first=>console.log(first));
    // })
  }
  catch(first)
  {
    console.log(first)
  }

}

ListCrudrootFolderSubFolder= async(fileInput:any="1") =>{

  try{
  
//  const file = fileInput.files[0];
      //   const fileUrl = `${file.name}`; // Relative URL for the file in the SharePoint list
      // const chunkSize = 10485760; // Chunk size (10 MB in this example)
      // const props = {
      //    progress: (data:any) => { console.log(`progress`); },
      //       ChunkSize: chunkSize,  // Optional: controls the chunk size
      //       Overwrite: true,       // Optional: controls whether to overwrite the file if it exists
      //   };
      // const props2 = {
      //     Overwrite: true,  // Optional: controls the chunk size
      //     AutoCheckoutOnInvalidData: false,   
      //     EnsureUniqueFileName:false
      //         // Optional: controls whether to overwrite the file if it exists
      //   };
     await  webObj.lists.getById(lid).rootFolder.folders().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName);
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).addSubFolderUsingPath("testk"+random(1,100,false)).catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).copyByPath("/Lists/emp/TestFolder2",true).catch(first=>console.log(first));
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).expand("Title")().catch(first=>console.log(first));
     await  this.ListCrudrootFolderSubFolderFiles(fileInput);
     await  kinds.map(async(kid:any)=>{
       await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).getShareLink(kid);
       })
     await  kinds.map(async(kid:any)=>{
         await  webObj.lists.getById(lid).rootFolder.folders.getByUrl("Test").deleteSharingLinkByKind(kid);
       })
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).getDefaultColumnValues();
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).clearDefaultColumnValues();
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).getItem().catch(first=>console.log(first));
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).getItem("Title", "Modified").catch(first=>console.log(first));
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).getItem<{ Id: number, Title: string }>("Id", "Title");
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).getObjectSharingSettings().catch(first=>console.log(first));
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).getSharingInformation().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).listItemAllFields().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).listItemAllFields.select("ID")().catch(first=>console.log(first));
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).listItemAllFields.expand("ID")().catch(first=>console.log(first));
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).moveByPath('https://tmsobb2.sharepoint.com/Movec/sss222a',true).catch(first=>console.log(first))
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).moveByPath('https://tmsobb2.sharepoint.com/Movec/sss222a',false).catch(first=>console.log(first))
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).parentFolder().catch(first=>console.log(first));
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).properties.expand()().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).properties.expand("Title")().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).properties.select("Title")().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).select("Id")().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).select("Title")().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).setDefaultColumnValues([{name:"Title",value:"aaa"}]).catch(first=>console.log(first))
     principleIds.forEach(async (pid:any)=>{
       await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).shareWith(pid).catch(first=>console.log(first))
     })
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).storageMetrics().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).storageMetrics.select("Id")().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).storageMetrics.select("LastModified")().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).storageMetrics.expand("Id")().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).storageMetrics.expand("LastModified")().catch(first=>console.log(first));;
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).unshare().catch(first=>console.log(first));;
     await  kinds.map(async(kid:any)=>{
           await  webObj.lists.getById(lid).rootFolder.folders.getByUrl("Test").unshareLink(kid);
         })
     await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).update({Name:"sss"}).catch(first=>console.log(first));;;
    await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).delete().catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).deleteWithParams({BypassSharedLock:false}).catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).deleteWithParams({BypassSharedLock:true}).catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).delete().catch(first=>console.log(first));;
    await  webObj.lists.getById(lid).rootFolder.folders.getByUrl(subFolderName).recycle().catch(first=>console.log(first));;
    await  webObj.lists.getById(lid).rootFolder.folders.addUsingPath('C',true).catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.folders.expand("Title")().catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.folders.select("Title")().catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.folders.getContextInfo().catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.folders.orderBy("Title")().catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.folders.top(1)().catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.folders.filter("Title eq Test")().catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.folders.skip(2)().catch(first=>console.log(first));
    await webObj.lists.getByTitle('Test').rootFolder.files.getByUrl("xyz.txt").delete().catch(first=>console.log(first));
  }
  catch(first)
  {
    console.log(first)
  }

}
ListCrudrootFolder= async(fileInput:any="1") =>{
  try{
    // const file = fileInput.files[0];
    await webObj.lists.getById(lid).rootFolder.getItem("Title").catch(first=>console.log(first));  
     const item = await webObj.lists.getByTitle(listName).rootFolder.getItem("Attackerfile.txt");

     await item.update({
      Title: "A Title",
      OtherField: "My Other Value"
    });

    await webObj.lists.getById(lid).rootFolder.addSubFolderUsingPath("testk"+random(1,100,false)).catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.checkSharingPermissions([{email:"tmsobb@tmsobb2.onmicrosoft.com"}])
    await  webObj.lists.getById(lid).rootFolder.clearDefaultColumnValues()
    await  webObj.lists.getById(lid).rootFolder.copyByPath(`/Lists/${listName}/TestFolder2`,true).catch(first=>console.log(first));
    await  webObj.lists.getById(lid).rootFolder.delete("*")
    await  webObj.lists.getById(lid).rootFolder.deleteSharingLinkByKind(1)
    await  webObj.lists.getById(lid).rootFolder.deleteWithParams({BypassSharedLock:true,ETagMatch:"*",DeleteIfEmpty:true})
    await  webObj.lists.getById(lid).rootFolder.expand("Title")
    await  webObj.lists.getById(lid).rootFolder.select("Title")
    this.ListCrudrootFolderFiles(fileInput) //rootFolder.files.
  }
  catch(first)
  {
    console.log(first)
  }

}

sendRequest = (method:string, url:string, body:any="",kind:any="",role:any="",scope:any="",sid:any=""):any => {

    try{
      return new Promise((resolve, reject) => {
          var xhr = new XMLHttpRequest();
          xhr.open(method, url, true);
          // Common headers
          xhr.setRequestHeader("content-type", "application\/json;odata=verbose");
          xhr.setRequestHeader("accept", "application\/json;odata=verbose");
          xhr.setRequestHeader("accept-language", "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7");
          xhr.setRequestHeader("X-RequestDigest", localStorage.getItem('reqd')??"*");
          xhr.withCredentials = true; 
          xhr.onload = function() {
              if (xhr.status >= 200 && xhr.status < 300) {
                  resolve(JSON.parse(xhr.responseText)); // Resolve the promise with the response text
              } else {
                  reject(new Error(xhr.statusText)); // Reject the promise with the status text
              }
          };
  
          xhr.onerror = function() {
              reject(new Error(xhr.statusText)); // Reject the promise on network error
          };
  
          // Convert body to Blob if provided
          if (body) {
              var aBody = new Uint8Array(body.length);
              for (var i = 0; i < aBody.length; i++) {
                  aBody[i] = body.charCodeAt(i);
              }
              xhr.send(new Blob([aBody]));
          } else {
              xhr.send();
          }
      });
    }
    catch(first)
    {
      console.log(first)
    }
  }
sendRequest2 = (method:string, url:string, body:any="",kind:any="",role:any="",scope:any="",sid:any="") => {

    try{
      return new Promise((resolve, reject) => {
          var xhr = new XMLHttpRequest();
          xhr.open(method, url, true);
          // Common headers
          xhr.setRequestHeader("content-type", "application\/json;odata=verbose");
          xhr.setRequestHeader("accept", "application\/json;odata=verbose");
          xhr.setRequestHeader("accept-language", "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7");
          xhr.setRequestHeader("X-RequestDigest", localStorage.getItem('reqd')??"*");
          xhr.withCredentials = true; 
          xhr.onload = function() {
              if (xhr.status >= 200 && xhr.status < 300) {
                  resolve(xhr.responseText); // Resolve the promise with the response text
              } else {
                  reject(new Error(xhr.statusText)); // Reject the promise with the status text
              }
          };
  
          xhr.onerror = function() {
              reject(new Error(xhr.statusText)); // Reject the promise on network error
          };
  
          // Convert body to Blob if provided
          if (body) {
              var aBody = new Uint8Array(body.length);
              for (var i = 0; i < aBody.length; i++) {
                  aBody[i] = body.charCodeAt(i);
              }
              xhr.send(new Blob([aBody]));
          } else {
              xhr.send();
          }
      });
    }
    catch(first)
    {
      console.log(first)
    }
  }
sendRequest3 = (method:string, url:string, body:any="",kind:any="",role:any="",scope:any="",sid:any="") => {

    try{
      return new Promise((resolve, reject) => {
          var xhr = new XMLHttpRequest();
          xhr.open(method, url, true);
          // Common headers
          xhr.setRequestHeader("content-type", "application\/json");
          xhr.setRequestHeader("Accept", "application/json;odata=verbose");
          xhr.setRequestHeader("accept-language", "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7");
          xhr.setRequestHeader("X-RequestDigest", localStorage.getItem('reqd')??"*");
          xhr.withCredentials = true; 
          xhr.onload = function() {
              if (xhr.status >= 200 && xhr.status < 300) {
                  resolve(xhr.responseText); // Resolve the promise with the response text
              } else {
                  reject(new Error(xhr.statusText)); // Reject the promise with the status text
              }
          };
  
          xhr.onerror = function() {
              reject(new Error(xhr.statusText)); // Reject the promise on network error
          };
  
          // Convert body to Blob if provided
          if (body) {
              var aBody = new Uint8Array(body.length);
              for (var i = 0; i < aBody.length; i++) {
                  aBody[i] = body.charCodeAt(i);
              }
              xhr.send(new Blob([aBody]));
          } else {
              xhr.send();
          }
      });
    }
    catch(first)
    {
      console.log(first)
    }
  }
  //to get data from the given list
_getDataFromList(listName: string, selectColumns: string, expandColumns: string, filterCondition: string, sortByColumnName: string, isAscending: boolean): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try{
      webObj.lists.getById(lid).items
        .select('Title');
    }catch(error){
      console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDataFromList \n Error-" + error.message);
          this._logError("DataService.tsx", "DataService", "_getDataFromList()", error.message);

          reject(error);
    }});
  }

  //to get group details of logged in user
_getCurrentUserGroups(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try{
      webObj.currentUser.groups()
        .then((val) => {
          resolve(val);
        }).catch((error: { message: string; }) => {
          console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getCurrentUserGroups \n Error-" + error.message);
          this._logError("DataService.tsx", "DataService", "_getCurrentUserGroups()", error.message);

          reject(error);
        });
    }catch(error){
      console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getCurrentUserGroups \n Error-" + error.message);
      this._logError("DataService.tsx", "DataService", "_getCurrentUserGroups()", error.message);

      reject(error);
    }});
  }
  //to get options of given choice column

  //to get current user details
public _getCurrentLoggedInUserData(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try{
      webObj.currentUser()
        .then((val) => {
          resolve(val);
        }).catch((error: { message: string; }) => {
          this._logError("DataService.tsx", "DataService", "_getCurrentLoggedInUserData()", error.message);

          console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getCurrentLoggedInUserData \n Error-" + error.message);
          reject(error);
        });
    }catch(error){
      this._logError("DataService.tsx", "DataService", "_getCurrentLoggedInUserData()", error.message);

          console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getCurrentLoggedInUserData \n Error-" + error.message);
          reject(error);
    }
  });
  }
  //upload image to document library
public _uploadImageToDocLibrary = (documentLibName: string, fileName: any, file: any, overwrite: boolean) => {
    return new Promise<any>((resolve, reject) => {
      try{
      webObj.getFolderByServerRelativePath(documentLibName).files.addUsingPath(fileName, file, { Overwrite: overwrite }).then(val => {
        resolve(val);
      }).catch((error) => {
        reject(error);
        this._logError("DataService.tsx", "DataService", "_uploadImageToDocLibrary()", error.message);
        console.log("Exception in DataService : \n Component Name-dataService \n Function Name- _uploadImageToDocLibrary \n Error-" + error.message);

      });
    }catch(error){
      reject(error);
      this._logError("DataService.tsx", "DataService", "_uploadImageToDocLibrary()", error.message);
      console.log("Exception in DataService : \n Component Name-dataService \n Function Name- _uploadImageToDocLibrary \n Error-" + error.message);
    }})
  }
public async _addMultipleListItems(listName: string, itemInfo: any,) {
    return new Promise<any>(async (resolve, reject) => {
      try {
        const list = webObj.lists.getById(lid);
        const [batchedSP, execute] = createBatch(list);
        list.using(batchedSP);
        itemInfo.forEach((item: any) => {
          list.items.add(item).catch((error) => reject(error))
        });
        await execute().then((response) => resolve(true)).catch((e: any) => {
          reject(e);
        });
      }
      catch (error) {
        console.log(error)
        this._logError("DataService.tsx", "DataService", "_addMultipleListItems()", error.message);
        console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _addMultipleListItems \n Error-" + error.message);
        reject(error)
      }

    });
  }
public async _updateListItemInFolder1(listName: string, itemInfo: any) {
    return new Promise<any>(async (resolve, reject) => {
      try{
      const [batchedSP, execute] = sp.batched();
      const list = batchedSP.web.lists.getById(lid);
      const res: any[] = [];
      itemInfo.forEach((item: any) => {
        let id = 0;
        id = item !== undefined && item.length > 0 && item.filter((field: any) => field.FieldName === 'Id').length > 0 ? item.filter((field: any) => field.FieldName === 'Id')[0].FieldValue : 0;
        item = item.filter((obj: any) => {
          return obj.FieldName !== "Id";
        });
        list.items.getById(id).validateUpdateListItem(
          item
        ).then(r => res.push(r)).catch(r => res.push(r));
      });
      const errorCount: any[] = [];
      await execute().then((s) => {
        for (let i = 0; i < res.length; i++) {
          const errors = res[i].filter((field: any) => field.ErrorMessage !== null);
          if (errors.length > 0) {
            errorCount.push(i);
          }
          if (i === (res.length - 1)) {
            resolve(errorCount);
          }
        }
      }).catch((e) => {
        this._logError("DataService.tsx", "DataService", "_updateListItemInFolder1()", e.message);
        console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _updateListItemInFolder1 \n Error-" + e.message);
        reject(errorCount);
      });
    }catch(error){
      this._logError("DataService.tsx", "DataService", "_updateListItemInFolder1()", error.message);
        console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _updateListItemInFolder1 \n Error-" + error.message);
        reject(error);
    }});
  }
public _updateFolderOfListByTitle = (listName: string, title: string) => {
    return new Promise<any>(async (resolve, reject) => {
      try {
        const item = await webObj.lists.getById(lid).items.filter(`Title eq 'xyz'`)();
        if (item.length === 0) {
          reject(new Error(`Item with title '${title}' not found in the list '${listName}'.`));
          return;
        }
        const itemId = item[0].Id;
        await webObj.lists.getById(lid).items.getById(itemId).update({
          Title: "iiiiiiiiiiiiii"
        });
        resolve("Item updated successfully.");
      } catch (error) {
        this._logError("DataService.tsx", "DataService", "_updateFolderOfListByTitle()", error.message);
        console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _updateFolderOfListByTitle \n Error-" + error.message);
        reject(error);
      }
    });
  }
public _createList = (listName: string): Promise<any> => {
    return new Promise<any>(async (resolve, reject) => {
      try {
        // const choices = [`Completed`, `Delayed`, `Need Attention`, `On Track`];
        // ensure that a list exists. If it doesn't it will be created with the provided title (the rest of the settings will be default):
        const listCreationResponse = await webObj.lists.ensure(listName, '', undefined, undefined, { Hidden: true });
        const list = await listCreationResponse.list
        const [batchedSP, execute] = createBatch(list);
        list.using(batchedSP);
    
        list.fields.addMultilineText("OverallComments", { RichText: false, RestrictedMode: false, AppendOnly: false, AllowHyperlink: true, Hidden: false })
        list.fields.addMultilineText("ChangeManagementComments", { RichText: false, RestrictedMode: false, AppendOnly: false, AllowHyperlink: true, Hidden: false })
        list.fields.addMultilineText("FinanceComments", { RichText: false, RestrictedMode: false, AppendOnly: false, AllowHyperlink: true, Hidden: false })
        await execute();
        resolve(listCreationResponse.created)
      } catch (error) {
        this._logError("DataService.tsx", "DataService", "_createList()", error.message);
        console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _createList \n Error-" + error.message);
        reject(error);
      }
    });
  };
public _deleteList = (listName: string): Promise<any> => {
    return new Promise<void>(async (resolve, reject) => {
      try{
      await webObj.lists.getById(lid).delete().then((res) => resolve(res)).catch((error) => {
        this._logError("DataService.tsx", "DataService", "_deleteList()", error.message);
        console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _deleteList \n Error-" + error.message);
        reject(error)
      });
    }catch(error){
      this._logError("DataService.tsx", "DataService", "_deleteList()", error.message);
        console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _deleteList \n Error-" + error.message);
        reject(error)
    }})
  }
  public _getCommentsData = (listName: string, id: number) => {
    const items: any = []
    return new Promise<any>((resolve, reject) => {
      try{
      webObj.lists.getById(lid)
        .items.getById(id).versions().then((versions) => {
          versions.forEach((item: any, index: number) => {
            if (item.Comments !== null && item.Comments !== '') {
              items.push({
                versionId: item.VersionId,
                comments: item.Comments,
                commentDateTime: item.Created,
                commentBy: item.Editor.LookupValue
              });
            }
          });
          resolve(items);

        }).catch((error: { message: string; }) => {
          console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getCommentsData \n Error-" + error.message);
          this._logError("DataService.tsx", "DataService", "_getCommentsData()", error.message);
          reject(error);
        });
    }
    catch(error){
      this._logError("DataService.tsx", "DataService", "_createList()", error.message);
      console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _getCommentsData() \n Error-" + error.message);
      reject(error);
    }
  }
 );
  }
  // To get the current week from the initiative status list 

  //to check whether current user has only read access
  public _checkifCurrentUserVisitor = async () => {
    let isVisitor: boolean = true;
    try {
      const perms = await webObj.getCurrentUserEffectivePermissions();
      if (webObj.hasPermissions(perms, PermissionKind.AddListItems) || webObj.hasPermissions(perms, PermissionKind.FullMask) || webObj.hasPermissions(perms, PermissionKind.EditListItems) || webObj.hasPermissions(perms, PermissionKind.DeleteListItems)) {
        isVisitor = false;
      }
    }
    catch (error) {
      console.log("Exception in DataService : \n Component Name-DataService.ts \n Function Name-_checkifCurrentUserVisitor \n Error-" + error.message);
      this._logError("DataService.tsx", "DataService", "_checkifCurrentUserVisitor()", error.message);
    }
    return isVisitor;
  }
  public _logError = (fileName: string, componentName: string, functionName: string, errMessage: string) => {
    return new Promise<any>(async (resolve, reject) => {
      try{
      await webObj.lists.getByTitle(`'${listName}'`).items.add({
      Title:"AAA"
      }).then(result => {
        resolve(result);
      }).catch((error) => {
        console.log(error);
        // this._logError("DataService.tsx", "DataService", "_logError()", error.message);
        console.log("Exception in DataService : \n Component Name-dataService \n Function Name- _logError \n Error-" + error.message);

        reject(error);
      });
    }catch(error)
    {
      console.log("Exception in DataService : \n Component Name-dataService \n Function Name- _logError \n Error-" + error.message);

        reject(error);
    }
  }
  );
  }

  public async _updateDisplayOrdersOfQuestions(listName: string, itemInfo: any,) {
    return new Promise<any>(async (resolve, reject) => {
      try {
        const list = webObj.lists.getById(lid);
        const [batchedSP, execute] = createBatch(list);
        list.using(batchedSP);
        itemInfo.forEach((item: any) => {
          const id = item.questionId
          list.items.getById(id).update({ DisplayOrder: item.displayOrder }).then((response) => resolve(response)).catch((error) => reject(error))
        });
        await execute().catch((e: any) => {
          reject(e);
        });
      }
      catch (error) {
        console.log(error)
        this._logError("DataService.tsx", "DataService", "_updateMultipleListItems()", error.message);
        console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _updateMultipleListItems \n Error-" + error.message);
        reject(error)
      }

    });
  }
//siteGroups calls

  //add user to DTO admin group
  public _addUserInDTOadminGroup(groupID: number, users: any): Promise<any> {
    // list.using(batchedSP);
    return new Promise<any>(async (resolve, reject) => { 
      try{


      // const user = await webObj.ensureUser(users.loginName)

      groupIds.map((gid:any)=>{
        personGroup.forEach((pg)=>{
          webObj.siteGroups.getById(gid).users.add(pg)
          .then((val) => {
            resolve(val);
          }).catch((error: { message: string; }) => {
            console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" + error.message);
            this._logError("DataService.tsx", "DataService", "_getDTOAdminGroupsUser()", error.message);
  
            reject(error);
          });
        })
      })
  
   
    }catch(error){
      console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" + error.message);
      this._logError("DataService.tsx", "DataService", "_getDTOAdminGroupsUser()", error.message);

      reject(error);
    }
    });
  }


  //to get detail of DTO admin
  public _getGroup(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try{
      webObj.siteGroups.getByName(groupName)()
        .then((val) => {
          resolve(val);
        }).catch((error: { message: string; }) => {
          console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroup \n Error-" + error.message);
          this._logError("DataService.tsx", "DataService", "_getDTOAdminGroup()", error.message);

          reject(error);
        });
    }catch(error){
      console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroup \n Error-" + error.message);
      this._logError("DataService.tsx", "DataService", "_getDTOAdminGroup()", error.message);

      reject(error);
    }
    });
  }

  //get all users of DTO admin group
  public getSIteGroups(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => { 
      try{
      webObj.siteGroups.getById(groupId).users()
        .then((val) => {
          resolve(val);
        }).catch((error: { message: string; }) => {
          console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" + error.message);
          this._logError("DataService.tsx", "DataService", "_getDTOAdminGroupsUser()", error.message);

          reject(error);
        });
    }catch(error){
      console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" + error.message);
      this._logError("DataService.tsx", "DataService", "_getDTOAdminGroupsUser()", error.message);

      reject(error);
    }});
  }

    //check if user is present in current DTO admin group
  public isUserInGroup(): Promise<any> {
      // list.using(batchedSP);
      return new Promise<any>(async (resolve, reject) => {
        try{
          groupIds.map((gid:any)=>{
            principleIds.forEach((pg)=>{
              webObj.siteGroups.getById(gid).users.getById(pg)().catch((error: { message: string; }) => {
                console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" + error.message);
                this._logError("DataService.tsx", "DataService", "_getDTOAdminGroupsUser()", error.message);
      
                reject(error);
              });
            })
          })
  
  
      }catch(error){
        console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" + error.message);
        this._logError("DataService.tsx", "DataService", "_getDTOAdminGroupsUser()", error.message);
  
        reject(error);
      }
      });
    }
  //get remove selected user 
  public _removeUserFromGroup(): Promise<any> {
    // list.using(batchedSP);
    return new Promise<any>(async (resolve, reject) => {
      try{
      webObj.siteGroups.getById(groupId).users.removeByLoginName(userLoginName)
        .then((val) => {
          resolve(val);
        }).catch((error: { message: string; }) => {
          console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" + error.message);
          this._logError("DataService.tsx", "DataService", "_getDTOAdminGroupsUser()", error.message);

          reject(error);
        });
    }catch(error){
      console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" + error.message);
          this._logError("DataService.tsx", "DataService", "_getDTOAdminGroupsUser()", error.message);

          reject(error);
    }});
  }

  public async _addOptionsToSiteColumn(fieldName: string, newChoices: string) {
    try {
      return new Promise(async (resolve, reject) => {
        // Field object
        let field = await webObj.fields.getByInternalNameOrTitle(fieldName);
        // Getting current choice values
        field.select('Choices')
          ().then((fieldData) => {
            // Append only unique choices
            let choices: string[] | any = fieldData.Choices; // For verbose & minimalmetadata
            let newUniqueChoices: any[] = [];
            if (choices.indexOf(newChoices) === -1) {
              newUniqueChoices.push(newChoices);
            }
            // Update choice field with new array of choices
            field
              .update({
                'Choices': choices.concat(newUniqueChoices)
              }, 'SP.FieldMultiChoice')
              .then(results => {
                resolve(true);
              }).catch((error) => {
                reject(error);
              });
          }).catch((error) => {
            this._logError("DataService.tsx", "DataService", "_addOptionsToSiteColumn()", error.message);
            console.log("Exception in DataService : \n Component Name- DataService \n Function Name-  _addOptionsToSiteColumn \n Error-" + error.message);

            reject(error);
          })
      });
    }
    catch (error) {
      console.log(error);
    }
  }

  public async _addAssignedByData(listName: string, dataToAdd: any[],objectiveId:any): Promise<any> {
    try {
      if (dataToAdd !== null && dataToAdd !== undefined) {
        if (dataToAdd.length > 0) {
          return new Promise<any>(async (resolve, reject) => {
            const list = webObj.lists.getById(lid);
            const [batchedSP, execute] = createBatch(list);
            list.using(batchedSP);
            dataToAdd.forEach((item: any) => {
              const addObject = {
                "OwnerLoginName": item.OwnerTitle.toString(),
                "Assigned_x0020_By": item.AssignedBy,
                "ObjectiveIdId": objectiveId,
                "IsDeleted": false
              };
              list.items.add(addObject).catch((error) => reject(error))
            });
            await execute().then((response) => resolve(true)).catch((e: any) => {
              this._logError("DataService.tsx", "DataService", "_addOrUpdateBulkData()", e.message);
              console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _addOrUpdateBulkData \n Error-" + e.message);
              reject(e);
            });
          });
        }
      }
    }
    catch(error) {
      console.log(error)
      this._logError("DataService.tsx", "DataService", "_addOrUpdateChecklist()", error.message);
      console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _addOrUpdateChecklist \n Error-" + error.message);
    }
  }

  //validate JSON string for owner and teammeber 
  isJsonString = (jsonString: any) => {
    try{
    if (typeof jsonString !== 'string') return false;
    try {
      const result = JSON.parse(jsonString);
      const type = Object.prototype.toString.call(result);
      return type === '[object Object]'
        || type === '[object Array]';
    } catch (err) {
      return false;
    }
  }catch(e)
  {
    this._logError("DataService.tsx", "DataService", "fetchDTOAdminListData()", e.message);
      console.log("Exception in DataService.tsx : \n Component Name- SynopsysPortfolioManagementIntakeTab \n Function Name-  fetchDTOAdminListData \n Error-" + e.message);
  }
  }

//update actuals data to the objective list


}