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
import * as React from "react";
// import styles from './Hacktheworld.module.scss';
import type { IHacktheworldProps } from "./IHacktheworldProps";
import "./Detaservice";
import { SiteUser, SPFI, spfi, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";
import "@pnp/sp/folders";
import "@pnp/sp/folders/web";
import "@pnp/sp/batching";
import "@pnp/sp/items/list";
import { Label, MessageBar, MessageBarType, PrimaryButton, Stack } from "@fluentui/react";
import DataService from "./Detaservice";
import DataService2 from "./DataService2";
import {CONFIG }  from'./config'
import { SHAREPOINT_CONSTANTS } from "./constants";
import ResearchApiPanel from "./research/ResearchApiPanel";
import { loadTestConfig } from "../config/loader";
import { checkLegacyAction } from "./legacy/legacyActionGuard";
const { 
  lid, itemId, folderId, groupId, groupName, userid, listName, fileName, SITE_URL, personGroup, kinds, roles, roles2, principleIds, HTTPMethods, groupIds,
  FOLDER_OPERATIONS, API_ENDPOINTS, VIEW_CONFIG, SHARING, MOVE_COPY, DATA_SERVICE
} = CONFIG; 
// import { random } from "@microsoft/sp-lodash-subset";
let sp: SPFI;
let webObj: IWeb;
let Token: string;
let dataService: DataService;
let dataService2: DataService2;

export default class Hacktheworld extends React.Component<IHacktheworldProps, { legacyNotice: string }> {
  ClientContext: any;
  public state = { legacyNotice: "" };
  constructor(props: IHacktheworldProps) {
    super(props);
  }

  private guardLegacyButton = (event: React.MouseEvent<HTMLElement>): void => {
    const target = event.target as HTMLElement;
    if (target.closest(".researchApiPanel") || target.closest(".ms-MessageBar")) return;
    const button = target.closest("button");
    if (!button) return;
    const label = (button.innerText || button.textContent || "Legacy API action").trim();
    const readiness = checkLegacyAction(label, loadTestConfig());
    if (readiness.missing.length > 0) {
      event.preventDefault();
      event.stopPropagation();
      this.setState({
        legacyNotice: "Cannot run ‘" + label + "’ yet. Update these settings in src/config/test-config.json, or use the Test config setup check above: " + readiness.missing.join("; ") + ".",
      });
      return;
    }
    if (readiness.destructive && !window.confirm("Run legacy action ‘" + label + "’? It may create, update, delete, or change sharing on tenant data.")) {
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    if (this.state.legacyNotice) this.setState({ legacyNotice: "" });
  };

  componentDidMount(): void {
    sp = spfi().using(SPFx(this.props.context));
    webObj = Web([sp.web, SITE_URL]);
    dataService = new DataService(this.props.context);
    dataService2 = new DataService2(this.props.context);
    this.token();
    Token = localStorage.getItem("reqd") ?? "*";
    console.log(Token);
    setInterval(() => {
      this.token();
      console.log("Interval triggered after 4 minutes");
    }, 240000); // 4 minutes in milliseconds
  }

  List_Item_Attachment = async () => {
    try{
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles().catch(err=>console.log(err))
    dataService.sendRequest("POST",`${SITE_URL}/_api/Web/Lists(guid'${lid}')/items(${itemId})/attachmentFiles`)?.catch(first=>console.log(first))
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.add(fileName, "dGVzdA==") .catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.add(fileName, "dGVzdA==") .catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).expand()().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).select()().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).getBlob().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).getBuffer().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).getBuffer().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).getJSON().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).getText().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).getText().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).setContent("xyz").catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).getText().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).getStream().catch((first) => console.log(first));
    dataService.sendRequest(
      "POST",
      `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/attachmentFiles(fileName='${fileName}')`
    );
    dataService.sendRequest(
      "POST",
      `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items('${itemId}')/attachmentFiles(fileName='${fileName}')/$value`
    );

    HTTPMethods.forEach((req) => {
      dataService.sendRequest(
        req,
        `${SITE_URL}/_api/web/GetList(@a1)/Items(@a2)/AttachmentFiles/addUsingPath(decodedUrl=@a3)?@a1=%27${encodeURIComponent(CONFIG.LIST_SERVER_RELATIVE_URL)}%27&@a2=${itemId}&@a3=%27${encodeURIComponent(fileName)}%27`,
        "Test"
      );
    });
    HTTPMethods.forEach((req) => {
      dataService.sendRequest(
        req,
        `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/attachmentFiles('${fileName}')/$value`,
        "Test"
      );
    });
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.add(fileName, "dGVzdA==").catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).delete().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).recycle().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.getByName(fileName).delete().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.add(fileName, "dGVzdA==").catch((first) => console.log(first));
    }
    catch(error)
    {
      console.log(error)
    }
  };
  list_Item_File = () => {
    // Get the list item by ID
    let item: any = webObj.lists.getById(lid).items.getById(itemId);
    webObj.lists.getById(lid).items.getById(itemId).file().catch((first) => console.log(first));
    webObj.lists.getById(lid).items.getById(itemId).file.versions().catch((first) => console.log(first));
    webObj.lists.getById(lid).items.getById(itemId).file.getText().catch((first) => console.log(first));
    webObj.lists.getById(lid).items.getById(itemId).file.getBlob().catch((first) => console.log(first));
    webObj.lists.getById(lid).items.getById(itemId).file.getJSON().catch((first) => console.log(first));
    webObj.lists.getById(lid).items.getById(itemId).attachmentFiles.add("Test", "TestContent").catch((first) => console.log(first));

    webObj.lists.getById(lid).items.getById(itemId).file.getItem().catch((first) => console.log(first)).then((fileItem) => console.log("File item metadata:", fileItem)).catch((error) =>
        console.log("Error getting file item metadata:", error)
      );
    webObj.lists.getById(lid).items.getById(itemId).file.delete().catch((first) => console.log(first)).then(() => console.log("File deleted successfully")).catch((error) => console.log("Error deleting file:", error));
    webObj.lists.getById(lid).items.getById(itemId).file.delete(item).then(() => console.log("File deleted successfully")).catch((error) => console.log("Error deleting file:", error));
  };

  list_Item_File2 = async (fileInput:any) => {
  const file =await fileInput.files[0];
  // const fileUrl =  `${file.name}`; // Relative URL for the file in the SharePoint list
  // const chunkSize = 10485760; // Chunk size (10 MB in this example)
  const fileContent = "Hello, SharePoint!";
  const blob = new Blob([fileContent], { type: "text/plain" });
  await  console.log(fileName)
   
  await  webObj.lists.getById(lid).items.getById(itemId).file().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).attachmentFiles().catch(err=>console.log(err))
  await  webObj.lists.getById(lid).items.getById(itemId).file.getBlob().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.copyByPath(`/Lists/${listName}/Attachments/${itemId}/`,true,true).catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.copyByPath(`/Lists/${listName}/Attachments/${itemId}/${fileName}`,true,true).catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.copyTo(`/Lists/${listName}/Attachments/${itemId}/`,true).catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.copyTo(`/Lists/${listName}/Attachments/${itemId}/`,true).catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.getBuffer().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.getItem().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.getJSON().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.getText().catch((first) => console.log(first));
  await  webObj.lists.getById(lid).items.getById(itemId).file.getBlob().catch((first) => console.log(first));
  await  webObj.lists.getById(lid).items.getById(itemId).file.getLimitedWebPartManager(1);//.catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.getLimitedWebPartManager(0);//.catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.getLockedByUser().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.getShareLink(5).catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.getSharingInformation().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.getStream().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.checkin().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.checkout().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.undoCheckout().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.approve().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.deleteSharingLinkByKind(5).catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.deny().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.exists().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.publish().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.setContent('ddd').catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.setContentChunked(file,{progress:data=>console.log("starting")}).catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.setContentChunked(file,{progress:data=>console.log("starting")}).catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.setContentChunked(blob,{progress:data=>console.log("starting")}).catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.unpublish().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.unshare().catch((first)=>console.log(first))
  await  webObj.lists.getById(lid).items.getById(itemId).file.unshareLink(5).catch((first)=>console.log(first))
  await webObj.lists.getById(lid).items.getById(itemId).file.versions().catch((first) => console.log(first));

  }
  share_list_folder = async () => {
    let itemId:any=folderId;

    try {
      var xhr = new XMLHttpRequest();
      xhr.open(
        "POST",
        `${SITE_URL}\/_api\/Web\/Lists(guid'${lid}')\/Items(${itemId})\/shareLink`,
        true
      );
      xhr.setRequestHeader("content-type", "application/json;charset=utf-8");
      xhr.setRequestHeader("accept", "application/json");
      xhr.setRequestHeader(
        "accept-language",
        "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
      );
      xhr.setRequestHeader(
        "X-RequestDigest",
        localStorage.getItem("reqd") ?? "*"
      );
      xhr.withCredentials = true;
      var body =
        '{"request":{"createLink":true,"emailData":null,"settings":{"expiration":null,"linkKind":5,"role":2}}}';
      var aBody = new Uint8Array(body.length);
      for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
      xhr.send(new Blob([aBody]));
     await roles.forEach(async (role) => {
        personGroup.forEach(async (pg) => {
          await  webObj.lists.getById(lid).items.getById(itemId).folder.shareWith(pg, role, true).catch((first) => console.log(first));
          await webObj.lists.getById(lid).items.getById(itemId).folder.shareWith(pg, role, false).catch((first) => console.log(first));
        });
     
      });
    await kinds.forEach(async (kid) => {
        webObj.lists .getById(lid)  .items.getById(itemId).folder.getShareLink(kid).catch((first) => console.log(first));
      });      
    await kinds.forEach(async (kid: any) => {
      await webObj.lists.getById(lid).items.getById(itemId).folder.deleteSharingLinkByKind(kid).catch((first) => console.log(first));
    });


    } catch (first) {
      console.log(first);
    }
  };
  GET_List_Item_Folder_Info = async () => {
    let itemId:any=folderId;
    try {
      //GET
      dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27%2FLists%2F${listName}%27&RootFolder=%2FLists%2F${listName}%2FTest&View=${CONFIG.VIEW_CONFIG.VIEW_ID}&TryNewExperienceSingle=TRUE`,
        JSON.stringify({
          parameters: {
            __metadata: { type: "SP.RenderListDataParameters" },
            RenderOptions: 1496583,
            AllowMultipleValueFilterForTaxonomyFields: true,
            AddRequiredFields: true,
            RequireFolderColoringFields: true,
          },
        })
      );
     let rs:any =  await webObj.lists.getById(lid).items.getById(itemId).folder.getItem().catch((first) => console.log(first));
     console.log(rs);
     
      await webObj.lists.getById(lid).items.getById(itemId).folder().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.folders().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.select("Title")().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.expand()().catch((first) => console.log(first));
      var body =
        `{"parameters":{"__metadata":{"type":"SP.RenderListDataParameters"},"RenderOptions":1496583,"ViewXml":"\x3cView Name=\\"{${VIEW_CONFIG.VIEW_ID}}\\" DefaultView=\\"TRUE\\" MobileView=\\"TRUE\\" MobileDefaultView=\\"TRUE\\" Type=\\"GRID\\" DisplayName=\\"All Items\\" Url=\\"/Lists/${listName}/AllItems.aspx\\" Level=\\"1\\" BaseViewID=\\"1\\" ContentTypeID=\\"0x\\" ImageUrl=\\"/_layouts/15/images/generic.png?rev=47\\"\x3e\x3cQuery/\x3e\x3cViewFields\x3e\x3cFieldRef Name=\\"DocIcon\\"/\x3e\x3cFieldRef Name=\\"LinkTitle\\"/\x3e\x3cFieldRef Name=\\"ID\\"/\x3e\x3cFieldRef Name=\\"Created\\"/\x3e\x3cFieldRef Name=\\"MyField\\"/\x3e\x3c/ViewFields\x3e\x3cRowLimit Paged=\\"TRUE\\"\x3e30\x3c/RowLimit\x3e\x3cJSLink\x3eclienttemplates.js\x3c/JSLink\x3e\x3cXslLink Default=\\"TRUE\\"\x3emain.xsl\x3c/XslLink\x3e\x3cToolbar Type=\\"Standard\\"/\x3e\x3c/View\x3e","AllowMultipleValueFilterForTaxonomyFields":true,"AddRequiredFields":true,"RequireFolderColoringFields":true}}`;
      dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/GetList(@v1)/RenderExtendedListFormData(itemId=${itemId},formId='editform',mode='2',options=30,cutoffVersion=0)?@v1=%27%2FLists%2F${listName}%27`
      )?.catch((first) => console.log(first));
      dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27%2FLists%2F${listName}%27&RootFolder=%2FLists%2F${listName}%2FTest&View=${CONFIG.VIEW_CONFIG.VIEW_ID}&TryNewExperienceSingle=TRUE`,
        body
      )?.catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.getContextInfo().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.getDefaultColumnValues().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.getParentInfos().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.getSharingInformation(
          {
            maxPrincipalsToReturn: 100,
            populateInheritedLinks: true,
            clientSupportedFeatures: "share",
          },
          ["permissionsInformation"],
          ["permissionsInformation", "anyoneLinkTrackUsers"]
        ).catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.listItemAllFields().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.parentFolder().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.storageMetrics().catch((first) => console.log(first));
      await dataService.sendRequest('GET',`${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/folder/Folders('${listName}')`)?.catch((first) => console.log(first))
      await dataService.sendRequest('GET',`${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/folder/Folders('${listName}')/Files`)?.catch((first) => console.log(first))
      await dataService.sendRequest('GET',`${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/folder/Files`)?.catch((first) => console.log(first))
      await dataService.sendRequest('POST',`${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/folder/Folders('${listName}')`)?.catch((first) => console.log(first))
      await dataService.sendRequest('PUT',`${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/folder/Folders('${listName}')`)?.catch((first) => console.log(first))
      await dataService.sendRequest('POST',`${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/folder/Folders('${listName}')/ParentFolder`)?.catch((first) => console.log(first))
      await dataService.sendRequest('POST',`${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/folder/ParentFolder`)?.catch((first) => console.log(first))
    } catch (first) {
      console.log(first);
    }
  };
  Update_Delete_List_Item_Folder_Info = async (): Promise<void> => {
    const itemId: any = folderId;
    
    try {
      // Helper function for HTTP requests with consistent error handling
      const makeHttpRequest = async (method: string, url: string, body?: string): Promise<void> => {
        try {
          await dataService.sendRequest(method, url, body);
        } catch (error) {
          console.log(`HTTP request failed for ${method} ${url}:`, error);
        }
      };

      // Helper function for PnP operations with consistent error handling
      const executePnPOperation = async (operation: () => Promise<any>, operationName: string): Promise<void> => {
        try {
          await operation();
        } catch (error) {
          console.log(`${operationName} failed:`, error);
        }
      };

      // Update operations
      const updateOperations = async (): Promise<void> => {
        // Validate update list item
        const body1 = `{"formValues":[{"FieldName":"Title","FieldValue":"${FOLDER_OPERATIONS.TEST_ITEM_TITLE}","HasException":false,"ErrorMessage":""}],"bNewDocumentUpdate":false,"RootFolder":"/Lists/${listName}/Test","View":"${VIEW_CONFIG.VIEW_ID}"}`;
        const url = `${API_ENDPOINTS.VALIDATE_UPDATE_FETCH_LIST_ITEM}?@a1=%27%2FLists%2F${listName}%27&@a2=${itemId}`;
        await makeHttpRequest("POST", url, body1);

        // Get list data as stream
        await makeHttpRequest(
          "POST",
          `${API_ENDPOINTS.GET_LIST_DATA_AS_STREAM}?listFullUrl=${encodeURIComponent("'" + CONFIG.PERSONAL_DOCUMENTS_URL + "'")}&TryNewExperienceSingle=TRUE`,
          JSON.stringify({
            parameters: {
              __metadata: { type: "SP.RenderListDataParameters" },
              RenderOptions: VIEW_CONFIG.RENDER_OPTIONS_2,
              ViewXml: VIEW_CONFIG.VIEW_XML,
              AddRequiredFields: true,
              RequireFolderColoringFields: true,
            },
          })
        );
      };

      // HTTP methods operations
   

      // PATCH operations
      const patchOperations = async (): Promise<void> => {
        const patchPayload = JSON.stringify({
          "__metadata": { "type": "SP.Folder" },
          "ProgID": FOLDER_OPERATIONS.FOLDER_PAYLOAD_PROG_ID,
          "ServerRelativeUrl": FOLDER_OPERATIONS.FOLDER_PAYLOAD_SERVER_URL,
          "Name": FOLDER_OPERATIONS.FOLDER_PAYLOAD_NAME
        });

        await makeHttpRequest(
          'PATCH',
          `${API_ENDPOINTS.FOLDER_OPERATIONS_BASE.replace('${lid}', lid).replace('${itemId}', itemId)}/Folders('${listName}')`,
          patchPayload
        );

        await makeHttpRequest(
          'PATCH',
          `${API_ENDPOINTS.FOLDER_OPERATIONS_BASE.replace('${lid}', lid).replace('${itemId}', itemId)})`,
          patchPayload
        );
      };

      // Properties operations
      const propertiesOperations = async (): Promise<void> => {
        const propertiesPayload = JSON.stringify({
          "__metadata": { "type": "SP.Folder" },
          "ProgID": FOLDER_OPERATIONS.FOLDER_PAYLOAD_PROG_ID,
          "ServerRelativeUrl": FOLDER_OPERATIONS.FOLDER_PAYLOAD_SERVER_URL,
          "Name": FOLDER_OPERATIONS.FOLDER_PAYLOAD_NAME
        });

        const propertiesUrls = [
          `${API_ENDPOINTS.FOLDER_OPERATIONS_BASE.replace('${lid}', lid).replace('${itemId}', itemId)}/Properties)`,
          `${API_ENDPOINTS.FOLDER_OPERATIONS_BASE.replace('${lid}', lid).replace('${itemId}', itemId)}/Folders('${listName}')/Properties)`
        ];

        for (const url of propertiesUrls) {
          await makeHttpRequest('GET', url, propertiesPayload);
          await makeHttpRequest('POST', url, propertiesPayload);
        }
      };

      // PnP operations
      const pnpOperations = async (): Promise<void> => {
        const folder = webObj.lists.getById(lid).items.getById(itemId).folder;

        // Folder management operations
        await executePnPOperation(
          () => folder.addSubFolderUsingPath(FOLDER_OPERATIONS.TEST_FOLDER_NAME),
          "Add sub folder using path"
        );

        await executePnPOperation(
          () => folder.moveByPath(`/Lists/${listName}/${listName}/${listName}`, true),
          "Move folder by path"
        );

        // Move file operations
        await makeHttpRequest(
          "POST",
          API_ENDPOINTS.MOVE_COPY_UTIL,
          JSON.stringify({
            destPath: {
              DecodedUrl: MOVE_COPY.DEST_PATH.replace('${listName}', listName),
              __metadata: { type: "SP.ResourcePath" },
            },
            options: {
              RetainEditorAndModifiedOnMove: MOVE_COPY.RETAIN_EDITOR_ON_MOVE,
              ShouldBypassSharedLocks: MOVE_COPY.BYPASS_SHARED_LOCKS,
              __metadata: { type: "SP.MoveCopyOptions" },
            },
            srcPath: {
              DecodedUrl: MOVE_COPY.SRC_PATH.replace('${listName}', listName),
              __metadata: { type: "SP.ResourcePath" },
            },
          })
        );

        // Sharing operations
        await executePnPOperation(
          () => folder.checkSharingPermissions(SHARING.SHARING_PERMISSIONS),
          "Check sharing permissions"
        );

        await executePnPOperation(
          () => folder.files(),
          "Get folder files"
        );

        await executePnPOperation(
          () => folder.copyByPath(`/Lists/${listName}/${listName}/${listName}`, true),
          "Copy folder by path"
        );

        await executePnPOperation(
          () => folder.setDefaultColumnValues([{ name: "Title", value: FOLDER_OPERATIONS.DEFAULT_COLUMN_VALUE }]),
          "Set default column values"
        );

        await executePnPOperation(
          () => folder.unshare(),
          "Unshare folder"
        );

        // Unshare link operations
        const shareId: any = await dataService.getShareLinkIds2().catch((error) => {
          console.log("Failed to get share link IDs:", error);
          return [];
        });

        for (const sid of shareId) {
          await executePnPOperation(
            () => folder.unshareLink(FOLDER_OPERATIONS.SHARE_LINK_KIND, sid),
            `Unshare link ${sid}`
          );
        }

        // Update and copy operations
        await executePnPOperation(
          () => folder.update({ Name: FOLDER_OPERATIONS.UPDATED_FOLDER_NAME }),
          "Update folder name"
        );

        await executePnPOperation(
          () => folder.copyByPath(`${SiteUser}/Test`, true),
          "Copy folder to user path"
        );

        await executePnPOperation(
          () => folder.clearDefaultColumnValues(),
          "Clear default column values"
        );
      };

      // Data service operations
      const dataServiceOperations = async (): Promise<void> => {
        await executePnPOperation(
          () => dataService._createFolderInList(listName, DATA_SERVICE.FOLDER_NAME),
          "Create folder in list"
        );

        await executePnPOperation(
          () => dataService._updateFolderOfListByTitle(listName, DATA_SERVICE.FOLDER_NAME),
          "Update folder of list by title"
        );
      };

      // List operations
      const listOperations = async (): Promise<void> => {
        const formValues = [{ FieldName: "Title", FieldValue: FOLDER_OPERATIONS.TEST_ITEM_TITLE_2 }];
        const folderPath = `/lists/${listName}/${listName}`;

        await executePnPOperation(
          () => webObj.lists.getById(lid).items.getById(itemId).list.addValidateUpdateItemUsingPath(
            formValues,
            folderPath,
            false,
            "Test"
          ),
          "Add validate update item using path (item level)"
        );

        await executePnPOperation(
          () => webObj.lists.getById(lid).items.getById(itemId).list.addValidateUpdateItemUsingPath(
            formValues,
            folderPath,
            false,
            "Test"
          ),
          "Add validate update item using path (item level) - duplicate"
        );

        await executePnPOperation(
          () => webObj.lists.getById(lid).addValidateUpdateItemUsingPath(
            [{ FieldName: "Title", FieldValue: FOLDER_OPERATIONS.TEST_ITEM_TITLE_3 }],
            `/Lists/${listName}/${listName}`
          ),
          "Add validate update item using path (list level)"
        );
      };
   const httpMethodsOperations = async (): Promise<void> => {
        const folderPayload = JSON.stringify({
          "ProgID": FOLDER_OPERATIONS.FOLDER_PAYLOAD_PROG_ID,
          "ServerRelativeUrl": FOLDER_OPERATIONS.FOLDER_PAYLOAD_SERVER_URL,
          "Name": FOLDER_OPERATIONS.FOLDER_PAYLOAD_NAME,
          "WelcomePage": FOLDER_OPERATIONS.FOLDER_PAYLOAD_WELCOME_PAGE
        });

        const folderPayloadSimple = JSON.stringify({
          "ProgID": FOLDER_OPERATIONS.FOLDER_PAYLOAD_PROG_ID,
          "ServerRelativeUrl": FOLDER_OPERATIONS.FOLDER_PAYLOAD_SERVER_URL,
          "Name": FOLDER_OPERATIONS.FOLDER_PAYLOAD_NAME
        });

        // Execute HTTP methods for folder operations
        const folderUrls = [
          `${API_ENDPOINTS.FOLDER_OPERATIONS_BASE.replace('${lid}', lid).replace('${itemId}', itemId)}/Folders('${listName}')`,
          `${API_ENDPOINTS.FOLDER_OPERATIONS_BASE.replace('${lid}', lid).replace('${itemId}', itemId)}')`
        ];

        for (const method of HTTPMethods) {
          for (const url of folderUrls) {
            await makeHttpRequest(method, url, folderPayload);
            await makeHttpRequest(method, url, folderPayloadSimple);
          }
        }

        // Add folder operations
        for (const method of HTTPMethods) {
          await makeHttpRequest(
            method,
            `${API_ENDPOINTS.FOLDER_OPERATIONS_BASE.replace('${lid}', lid).replace('${itemId}', itemId)}/folders/add('${FOLDER_OPERATIONS.TEST_FOLDER_PATH}')')`,
            JSON.stringify({})
          );
        }
      };
      // Execute all operations
      await Promise.all([
        updateOperations().catch(error => console.log("Update operations failed:", error)),
        patchOperations().catch(error => console.log("PATCH operations failed:", error)),
        propertiesOperations().catch(error => console.log("Properties operations failed:", error)),
        pnpOperations().catch(error => console.log("PnP operations failed:", error)),
        dataServiceOperations().catch(error => console.log("Data service operations failed:", error)),
        listOperations().catch(error => console.log("List operations failed:", error)),
                httpMethodsOperations().catch(error => console.log("HTTP methods operations failed:", error)),

      ]);

    } catch (error) {
      console.log("Update_Delete_List_Item_Folder_Info failed:", error);
    }
  };
  List_Item_Folder_Delete = async () => {
    let itemId:any=folderId;
    try {
      dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27%2FLists%2F${listName}%27&RootFolder=%2FLists%2F${listName}%2FSSS&View=${CONFIG.VIEW_CONFIG.VIEW_ID}&TryNewExperienceSingle=TRUE`,
        JSON.stringify({
          parameters: {
            __metadata: { type: "SP.RenderListDataParameters" },
            RenderOptions: 1496583,
            AllowMultipleValueFilterForTaxonomyFields: true,
            AddRequiredFields: true,
            RequireFolderColoringFields: true,
          },
        })
      );
      await webObj.lists.getById(lid).items.getById(itemId).folder.deleteWithParams({
          BypassSharedLock: true,
          DeleteIfEmpty: true,
          ETagMatch: "*",
        }).catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.recycle().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).folder.delete().catch((first) => console.log(first));
    } catch (first) {
      console.log(first);
    }
  };
  ListItemInformation = async () => {
    try {
      // const perms = await webObj.getCurrentUserEffectivePermissions();
      // for(let i=0;i<20;i++)
      // {
      //   webObj.hasPermissions(perms,i)
      // }
      var query = {
        ViewXml:
          "<View><Query><Where><Contains><FieldRef Name='Title'/><Value Type='Text'>111</Value></Contains></Where></Query></View>",
      };
      var query2 = {
        ViewXml: "<View><RowLimit>5000</RowLimit></View>",
      };
     await webObj.lists.getById(lid).getItemsByCAMLQuery(query).catch((first) => console.log(first));
     await webObj.lists.getById(lid).getItemsByCAMLQuery(query2).catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).hasPermissions({"Low":1,"High":1},4);
    await  webObj.lists.getById(lid).items.getItemByStringId(`${itemId}`)().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).getObjectSharingSettings().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId)().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).getLikedBy().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).effectiveBasePermissions().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).getCurrentUserEffectivePermissions().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).getLikedByInformation().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).fieldValuesForEdit().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).fieldValuesAsText().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).fieldValuesAsHTML().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).fieldValuesAsHTML.select("Title")().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).firstUniqueAncestorSecurableObject().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).getListItemChangesSinceToken({ ChangeToken: "" }).catch((first) => console.log(first));
    await  dataService.sendRequest("POST",`${SITE_URL}/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListFilterData?@a1=%27%2FLists%2F${listName}%27&FieldInternalName=%27Title%27`)
    await  dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/parentList`
      );
    await  dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/Web/Lists(guid'${lid}')/Author`
      );
    await  dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/Web/Lists(guid'${lid}')/DefaultView`
      );
  
    await  dataService.sendRequest2(
        "GET",
        `${SITE_URL}/_api/v2.0/sites/${CONFIG.GRAPH_SITE_ID}/lists/${lid}/items/${itemId}/Versions`
      );
     await dataService.sendRequest2(
        "GET",
        `${SITE_URL}/_api/v2.0/sites/${CONFIG.GRAPH_SITE_ID}/lists/${lid}/items/${itemId}/Analytics`
      );
    await  dataService.sendRequest2(
        "GET",
        `${SITE_URL}/_api/v2.0/sites/${CONFIG.GRAPH_SITE_ID}/lists/${lid}/items/${itemId}/activities?$expand=listItem($select=id,name,webUrl)&$top=25`
      );
    await  dataService.sendRequest2(
        "GET",
        `${SITE_URL}/_api/v2.1/sites/${CONFIG.GRAPH_SITE_ID}/lists/${lid}/items/${itemId}`      );
    await  dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/Web/Lists(guid'${lid}')/contentTypes`
      );
    await  dataService.List_Item_Info();
    await  this.ListItemsHARINGInformation();
   await   dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/displayName`
      );
     await dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/displayName`,
        "{}"
      );
    } catch (error) {
      console.log(error);
    }
  };
LibraryItemInformation = async () => {
    try {
   
      var query = {
        ViewXml:
          "<View><Query><Where><Contains><FieldRef Name='Title'/><Value Type='Text'>111</Value></Contains></Where></Query></View>",
      };
      var query2 = {
        ViewXml: "<View><RowLimit>5000</RowLimit></View>",
      };
    await  webObj.lists.getById(lid).getItemsByCAMLQuery(query).catch((first) => console.log(first));
    await  webObj.lists.getById(lid).getItemsByCAMLQuery(query2).catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).hasPermissions({"Low":1,"High":1},4);
    await  webObj.lists.getById(lid).items.getItemByStringId(`${itemId}`)().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).getObjectSharingSettings().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId)().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).getLikedBy().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).effectiveBasePermissions().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).getCurrentUserEffectivePermissions().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).getLikedByInformation().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).fieldValuesForEdit().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).fieldValuesAsText().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).fieldValuesAsHTML().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).fieldValuesAsHTML.select("Title")().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).items.getById(itemId).firstUniqueAncestorSecurableObject().catch((first) => console.log(first));
    await  webObj.lists.getById(lid).getListItemChangesSinceToken({ ChangeToken: "" }).catch((first) => console.log(first));
    await  dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/web/lists(@a1)/GetItemById(3)/?@a1=%27%7${lid}%7D%27`    );
    await  dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/parentList`
      );
    await  dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/Web/Lists(guid'${lid}')/Author`
      );
    await  dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/Web/Lists(guid'${lid}')/DefaultView`
      );
    await  dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/Web/Lists(guid'${lid}')/contentTypes`
      );
    await  dataService.Library_Item_Info();
    await  this.ListItemsHARINGInformation();
    await  dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/displayName`
      );
    await  dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/displayName`,
        "{}"
      );
    } catch (error) {
      console.log(error);
    }
  };
  ListItemVersions = async () => {
    // Fetch the versions of the item and extract VersionIds
    try {
      let versionIds: any = [];
      await webObj.lists.getById(lid).items.getById(itemId).versions().then((ids: any) => {
          ids.forEach((version: any) => versionIds.push(version.VersionId));
          console.log(versionIds);
        }).catch((first) => console.log(first));
      await versionIds.map(async (idd: any, index: any) => {
        dataService.sendRequest(
          "GET",
          `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/versions(${idd})`
        );
        index != 0
          ? webObj.lists.getById(lid).items.getById(itemId).versions.getById(idd).delete(idd).catch((first) => console.log(first))
          : "";
      });
      await webObj.lists.getById(lid).items.getById(itemId).versions.filter("ID eq " + versionIds[0])().catch((first: any) => console.log(first));
      await dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/versions(${
          versionIds[1] ? versionIds[1] : 512
        })`
      );

      HTTPMethods.forEach((req) => {
        dataService.sendRequest(
          req,
          `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/versions`
        );
      });
    } catch (err) {
      console.log("");
    }
  };
  ListItemAddUpdate = async () => {
    try {
      await webObj.lists.getById(lid).addValidateUpdateItemUsingPath(   [{ FieldName: "Title", FieldValue: "sss" }], `/lists/${listName}`,false,"Test").catch(first=>console.log(first))
      await dataService.sendRequest("POST",`${SITE_URL}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/Files/AddUsingPath(DecodedUrl=@a2,AutoCheckoutOnInvalidData=@a3,XorHash=@a4)?@a1=%27${encodeURIComponent(CONFIG.LIST_SERVER_RELATIVE_URL)}%27&@a2=%27${encodeURIComponent(fileName.replace(/(\.[^.]+)?$/, "12$1"))}%27&@a3=true&@a4=%27AAAAAAAAAAAAAAAAAAAAAAAAAAA%3D%27&$Select=ServerRelativeUrl,UniqueId,Name,VroomItemID,VroomDriveID,ServerRedirectedUrl&$Expand=ListItemAllFields`)
      await dataService.sendRequest("POST",`${SITE_URL}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/Files/AddUsingPath(DecodedUrl=@a2,AutoCheckoutOnInvalidData=@a3,XorHash=@a4)?@a1=%27${encodeURIComponent(CONFIG.LIST_SERVER_RELATIVE_URL)}%27&@a2=%27${encodeURIComponent(fileName.replace(/(\.[^.]+)?$/, "1233$1"))}%27&@a3=false&@a4=%27AAAAAAAAAAAAAAAAAAAAAAAAAAA%3D%27&$Select=ServerRelativeUrl,UniqueId,Name,VroomItemID,VroomDriveID,ServerRedirectedUrl&$Expand=ListItemAllFields`)
      await dataService.sendRequest("POST",`${SITE_URL}/_api/web/GetListUsingPath(DecodedUrl=@a1)/Items(@a2)?@a1=%27%2F${listName}%27&@a2=%27${itemId}%27`,JSON.stringify({"__metadata":{"type":"SP.Data.SSSSItem"},"FileLeafRef":"Document3333.docx"}))
      await dataService.sendRequest("MERGE",`${SITE_URL}/_api/web/lists(@a1)/GetItemById(3)/?@a1=%27%7B${lid}%7D%27`,JSON.stringify({"Title":"kkkkkkkkkkkkkkkkkk"}));
      await dataService.sendRequest2("MERGE",`${SITE_URL}/_api/v2.1/favorites/followedListItems/oneDrive.add`,        JSON.stringify({value: [{siteId: `{${CONFIG.SITE_COLLECTION_ID}}`,webId: `{${CONFIG.WEB_ID}}`,listId: `{${lid}}`,listItemUniqueId: `{${CONFIG.ITEM_UNIQUE_ID}}`,},],}));
      await webObj.lists.getById(lid).items.getById(itemId).update({ Title: "ss" }).catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).update({ Title: "ttt" }, "*").catch((first) => console.log(first));
       await webObj.lists.getById(lid).items.getById(itemId).validateUpdateListItem(
          [{ FieldName: "Title", FieldValue: "sss", ItemId: itemId }],
          false
        ).catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).list.addValidateUpdateItemUsingPath(
          [{ FieldName: "Title", FieldValue: "sss" }],
          `/lists/${listName}`,
          false,
          "Test"
        ).catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.add({ Title: "Title" }).catch((first) => console.log(first));
   
        HTTPMethods.forEach(async (req) => {
        await dataService.sendRequest(
          req,
          `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/AddValidateUpdateItemUsingPath()?formValues=[{"FieldName":"Title","FieldValue":"111111111111111111111111","ItemId":${itemId}}]`,
          JSON.stringify({
            ID: itemId,
          })
        );
        await dataService.sendRequest(
          req,
          `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/validateupdatelistitem()?formValues=[{"FieldName":"Title","FieldValue":"111111111111111111111111","ItemId":${itemId}}]`
        );
        await dataService.sendRequest(
          req,
          `${SITE_URL}/_api/web/GetListUsingPath(DecodedUrl=@a1)/BulkValidateUpdateListItems()?@a1='/Lists/${listName}'&formValues=[{"FieldName":"_ModerationStatus","FieldValue":"0"},{"FieldName":"_ModerationComments","FieldValue":"test"}]&itemIds=[${itemId}]`
        );
        await dataService.sendRequest(req,`${SITE_URL}/_api/web/lists(@a1)/GetItemById(3)/?@a1=%27%7B${lid}%7D%27`,JSON.stringify({"__metadata":{"type":`SP.Data.${listName.charAt(0).toLocaleUpperCase()}${listName.substring(1,listName.length)}ListItem`},"Title":"Attacker1"}));
    
        await dataService.sendRequest(
          req,
          `${SITE_URL}/_api/SP.MoveCopyUtil.MoveFileByPath()`,
          JSON.stringify({
            destPath: {
              DecodedUrl: `${SITE_URL}/Lists/${listName}/Test/1_.000`,
              __metadata: { type: "SP.ResourcePath" },
            },
            options: {
              RetainEditorAndModifiedOnMove: true,
              ShouldBypassSharedLocks: true,
              __metadata: { type: "SP.MoveCopyOptions" },
            },
            srcPath: {
              DecodedUrl: `${SITE_URL}/Lists/${listName}/1.000`,
              __metadata: { type: "SP.ResourcePath" },
            },
          })
        );
      });
      // await  webObj.lists.getById(lid).items.getById(itemId).setImageField("imageData","test.png ","UklGRi77BABXRUJQVlA4IHzZAwBQjQydASoABAAEPjEYiUOiIZER2hUkIAMEs7MOWmvvtOL3+3n7U95HyoHX/nDAi+K9kcAkUHO313XHEHx58Abxn+3a4NgDaAmpLljA3/n/T79WbeZD/K/sJ/iv/J/tPrf5R8m/LH4n84/4P/kf5f71umb3o+Q/8H0AfA55t+rf4f+6ftd/aP///4vu9+7f5LfRT9bf9P3Av4l/Lf8X/bP9L/nP79////D9c/7PfDX94fUV/WP7l/t/8P+8nzN/9L/lf5T4C/4H/Z/tJ/rvkD/of9t/6H51fHD7L37uf//3EP6F/pv/L65//q/1f/C////f+2P+l/6X/w/5v9//+Z9jH81/u3++/a7//fIB/5vaw/gH7/+wB2pPjm+dfu3+g/Mj8RPYX9A+2/8/+P/Nb5c/3X/+/8vei95/0f3m9Qf7J/Ev9v/L/43/j/w/7g/cb+J/afxF+Rf/x/pvYF/PP7Z/4v29").catch(first=>console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).like().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).unshare().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).unlike().catch((first) => console.log(first));
      this.validatelistItem();
      await dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/GetListUsingPath(DecodedUrl=@a1)/BulkValidateUpdateListItems()?@a1='/Lists/${listName}'`,
        JSON.stringify({
          itemIds: [itemId],
          formValues: [{ FieldName: "Title", FieldValue: "XYZ1" }],
          folderPath: "",
        })
      );
      await dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/GetListUsingPath(DecodedUrl=@a1)/BulkValidateUpdateListItems()?@a1='/Lists/${listName}'`,
        JSON.stringify({
          itemIds: [itemId],
          formValues: [
            { FieldName: "_ModerationComments", FieldValue: "qqqq" },
          ],
          folderPath: "",
        })
      );
      await dataService.sendRequest(
        "MERGE",
        `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})`,
        JSON.stringify({ Title: "AAAAAAAAAAA" })
      );
      await dataService.itemUpdateXMLCalls();
      await dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/lists('${lid}')/items(${itemId})/deleteObject`
      );
      await dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/deleteObject`
      );
      await webObj.lists.getById(lid).items.getById(itemId).delete("*").catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).recycle().catch((first) => console.log(first));
      await dataService.sendRequest("GET",`${SITE_URL}/_api/SP.Approvals.GetItemRequest?listId=%27${lid}%27&itemId=%27${itemId}%27&url=%27%27&version=%272%27`)?.catch(first=>console.log(first))
      await dataService.sendRequest("",`${SITE_URL}/_api/SP.Approvals.GetItemRequest?listId=%27${lid}%27&itemId=%27${itemId}%27&url=%27%27&version=%272%27`,{})?.catch(first=>console.log(first))
      await dataService.sendRequest("POST",`${SITE_URL}/_api/SP.Approvals.CreateItemRequest`,JSON.stringify({"creationInfo":{"listId":`${lid}`,"itemId":`${itemId}`,"url":"","version":"1","approvers":`${CONFIG.SHARING.EMAILS.ATTACKER},${CONFIG.SHARING.EMAILS.TMSOBB}`,"title":"Test","details":"","awaitAll":true,"itemUrlType":0,"markDocAsFinal":false}}))?.catch(first=>console.log(first))
      await dataService.sendRequest("POST",`${SITE_URL}/_api/SP.Approvals.CancelItemRequest`,JSON.stringify({"listId": `${lid}`,"itemId":`${itemId}`,"url":"","version":"2","approvalId":CONFIG.APPROVAL_ID}))?.catch(first=>console.log(first))
      await dataService.sendRequest("POST",`${SITE_URL}/_api/SP.Approvals.ApproveItemRequest`,JSON.stringify({"listId": `${lid}`,"itemId":`${itemId}`,"url":"","version":"2","approvalId":CONFIG.APPROVAL_ID}))?.catch(first=>console.log(first))
      await dataService.sendRequest("POST",`${SITE_URL}/_api/SP.Approvals.RejectItemRequest`,JSON.stringify({"listId": `${lid}`,"itemId":`${itemId}`,"url":"","version":"2","approvalId":CONFIG.APPROVAL_ID}))?.catch(first=>console.log(first))
       HTTPMethods.map((req: any) => {
        dataService.sendRequest(
          req,
          `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/recycle`
        );
        dataService.sendRequest(
          req,
          `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/deleteObject`
        );
      });
      // dataService._createFolderInList(listName, "Testxyz");
      // dataService._ensureFolder(`/Lists/${listName}`, listName, "Testxyz");
      // dataService._ensureFolder(`/Lists/${listName}`, listName, "Test");
      // dataService._updateFolderOfListByTitle(listName, "Test");
    } catch (first) {
      console.log(first);
    }
  };
  deleteList = async () => {
    await dataService.sendRequest("POST",`${SITE_URL}/_api/web/lists/getByTitle('${listName}')/DeleteObject/`)
    await dataService.sendRequest("POST",`${SITE_URL}/_api/web/lists/getByTitle('${listName}')/StartRecycle()`)
    await dataService.sendRequest("PUT",`${SITE_URL}/_api/web/lists/getByTitle('${listName}')/DeleteObject/`)
    await dataService.sendRequest("DEETE",`${SITE_URL}/_api/web/lists/getByTitle('${listName}')/DeleteObject/`)
    await webObj.lists.getById(lid).delete().catch((first) => console.log(first));
    await webObj.lists.getById(lid).recycle().catch((first) => console.log(first));



  };
  validatelistItem = () => {
    //approve list item
    try {
      var xhr = new XMLHttpRequest();
      xhr.open(
        "POST",
        `${SITE_URL}\/_api\/web\/GetListUsingPath(DecodedUrl=@a1)\/BulkValidateUpdateListItems()?@a1=%27%2FLists/${listName}%27`,
        true
      );
      xhr.setRequestHeader("content-type", "application/json;odata=verbose");
      xhr.setRequestHeader("accept", "application/json;odata=verbose");
      xhr.setRequestHeader(
        "accept-language",
        "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
      );
      xhr.setRequestHeader(
        "X-RequestDigest",
        localStorage.getItem("reqd") ?? "*"
      );
      xhr.withCredentials = true;
      var body = `{\"itemIds\":[${itemId}],\"formValues\":[{\"FieldName\":\"_ModerationStatus\",\"FieldValue\":\"0\"},{\"FieldName\":\"_ModerationComments\",\"FieldValue\":\"test\"}],\"folderPath\":\"\"}`;
      var aBody = new Uint8Array(body.length);
      for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
      xhr.send(new Blob([aBody]));
    } catch (first) {
      console.log(first);
    }
  };
  ListItemComments = async (): Promise<any> => {
    try {
      await webObj.lists.getById(lid).items.getById(itemId).comments.add("Attacker added").catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).comments.add("Test").then((res: any): Promise<any> => {
          return new Promise(async (resolve, reject) => {
            return resolve(res.id);
          });
        }).catch((first) => console.log(first));
      await dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/comments(1)/text`
      );
      await dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/comments(2)/text`
      );
      await dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/comments(2)/text`
      );
      await dataService.sendRequest(
        "GET",
        `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/comments(2)/author`
      );
      await dataService.sendRequest(
        "POST",
        `${SITE_URL}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/comments(2)/author`
      );
      await dataService.addCommentWithMention();
      let cIds: number[] = [];
      dataService.addComment();
      await webObj.lists.getById(lid).items.getById(itemId).comments().then((res: any) => {
          cIds = res.map((comment: any) => {
            return comment.id;
          });
        }).catch((first) => console.log(first));
      let commentId = cIds.length > 0 ? cIds[0] : 2;
      await dataService.patchComment(commentId);
      await webObj.lists.getById(lid).items.getById(itemId).comments.getById(commentId).replies.add("ttt").catch((first) => console.log(first));
      // await  webObj.lists.getById(lid).items.getById(itemId).comments.getById(commentId).like().catch(first=>console.log(first))
      await webObj.lists.getById(lid).items.getById(itemId).comments.getById(commentId).replies().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).comments.getById(commentId).replies.add("Hiii").catch((first) => console.log(first));
      // await webObj.lists.getById(lid).items.getById(itemId).comments.getById(commentId).unlike().catch(first=>console.log(first))
      await webObj.lists.getById(lid).items.getById(itemId).comments.getById(commentId).replies.top(itemId);
    } catch (first) {
      console.log(first);
    }
  };
  deleteListItemComments = async () => {
    let cIds: number[] = [];
    await webObj.lists.getById(lid).items.getById(itemId).comments().then((res: any) => {
        cIds = res.map((comment: any) => {
          return comment.id;
        });
      }).catch((first) => console.log(first));
    let commentId = cIds.length > 0 ? cIds[0] : 8;

    await webObj.lists.getById(lid).items.getById(itemId).comments.getById(commentId).delete().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).comments.getById(commentId - 1).delete().catch((first) => console.log(first));
    await webObj.lists.getById(lid).items.getById(itemId).comments.clear();

    dataService.deleteComment(commentId);
  };
  ListItemsHARINGInformation = async () => {
    try {
      console.log("ListItemsHARINGInformation")
      await webObj.lists.getById(lid).items.getById(itemId).getSharingInformation(
          {
            maxPrincipalsToReturn: 100,
            populateInheritedLinks: true,
            clientSupportedFeatures: "share",
          },
          ["permissionsInformation"],
          ["permissionsInformation", "anyoneLinkTrackUsers"]
        ).catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).getObjectSharingSettings().catch((first) => console.log(first));
      await webObj.lists.getById(lid).items.getById(itemId).checkSharingPermissions([{ email: `${CONFIG.SHARING.EMAILS.TMSOBB}` }]).catch((first) => console.log(first));
    } catch (error) {
      console.log(error);
    }
  };
  Share_List_Item_pnp_js = async () => {
    try {
      for (const kind of kinds) {
        try {
          await webObj.lists.getById(lid).items.getById(itemId).getShareLink(kind).catch((first) => console.log(first));
        } catch (error) {
          console.log(error);
        }
        try {
          await webObj.lists.getById(lid).items.getById(itemId).getShareLink(kind, new Date(SHAREPOINT_CONSTANTS.DATES.FUTURE_EXPIRATION)).catch((first) => console.log(first));
        } catch (error) {
          console.log(error);
        }
      }

      // const folderPath = `${listName}/${listName}`;
      // for (let i = 5; i >= 2; i--) {
      //   try {
      //     await webObj.getFolderByServerRelativePath(folderPath).getShareLink(i).catch((first) => console.log(first));
      //   } catch (error) {
      //     console.log(error);
      //   }
      // }
    } catch (error) {
      console.log(error);
    }
  };
  Share_List_Item_XHR = async () => {
    await dataService.Share_List_Item_XHRCAll1(); // SHARE ITEM DIRECT TO USER/GROUP , DO NOT CREATE LINK
    await dataService.Share_List_Item_XHRCAll2();
    await dataService.Share_List_Item_XHRCAll4();
  };
  Share_List_Item_with_Users = async () => {
    await dataService.Share_List_Item_XHRCAll2share_wthUsers();
    await dataService.Share_List_Item_XHRCAll2share_wthUsers2();
    await dataService.Share_List_Item_XHRCAll2share_wthUsers2_peoplechoosepermission_rm_true();
  };
  Share_List_Item_expiry_date = async () => {
    await dataService.Share_List_Item_XHRCAll2_dup_expdate();
    await dataService.Share_List_Item_XHRCAll2_dup_expdate2();
    await dataService.Share_List_Item_XHRCAll2_dup_expdate_withUser();
  };
  Share_List_Item_WithPassword = async () => {
    try {
      await dataService.Share_List_Item_XHRCAll_WithPassword();
      await dataService.Share_List_Item_XHRCAll_WithPassword2();
      await dataService.Share_List_Item_XHRCAll_WithPassword2_withUser();
    } catch (first) {
      console.log(first);
    }
  };
  Share_List_Item_role6 = async () => {
    //pnp js call
    await dataService.Share_List_Item_XHRCAll_rolekind6();
  };

  Delete_Sharing_Link = async () => {
    let shareId: any =
      (await dataService.getShareLinkIds2().catch((first) => console.log(first))) ?? [];
    // await webObj.lists.getById(lid).items.getById(itemId).unshare();
    console.log(shareId);

    try {
      kinds.map((kind) => {
        shareId.map(async (id) => {
          await webObj.lists.getById(lid).items.getById(itemId).unshareLink(kind, id).catch((first) => console.log(first));
        });
      });
    } catch (error) {
      console.log(error);
    }
    await dataService.deleteshareLinkXHR();
  };
  roleInheritance = () => {
    webObj.lists.getById(lid).items.getById(itemId).breakRoleInheritance(true, true);
    webObj.lists.getById(lid).items.getById(itemId).resetRoleInheritance();
    webObj.lists.getById(lid).breakRoleInheritance(true, true);
    webObj.lists.getById(lid).resetRoleInheritance();
  };
  updateroleAssignmentListItems = () => {

    try{
    principleIds.forEach(async (principleId) => {
      roles.forEach(async (role) => {
        await webObj.lists.getById(lid).items.getById(itemId).roleAssignments.add(principleId, role);
      });
          webObj.lists.getById(lid).items.getById(itemId).roleAssignments.getById(itemId).groups.add({
        AllowMembersEditMembership: true,
        AllowRequestToJoinLeave: false,
        AutoAcceptRequestToJoinLeave: true,
        Description: "This is a sample description for the group.",
        Id: principleId,
        IsHiddenInUI: false,
        OnlyAllowMembersViewMembership: true,
        OwnerTitle: SHAREPOINT_CONSTANTS.USERS.TMSOBB_NAME,
        PrincipalType: 8,
        RequestToJoinLeaveEmailSetting: CONFIG.SHARING.EMAILS.TMSOBB,
        Title: groupName,
      }).catch(first=>console.log(first));
    });
    personGroup.forEach(async (personG) => {0
          webObj.lists.getById(lid).items.getById(itemId).roleAssignments.getById(itemId).groups.add({
        AllowMembersEditMembership: true,
        AllowRequestToJoinLeave: false,
        AutoAcceptRequestToJoinLeave: true,
        Description: "This is a sample description for the group.",
      
        IsHiddenInUI: false,
        LoginName: personG,
        OnlyAllowMembersViewMembership: true,
        OwnerTitle: SHAREPOINT_CONSTANTS.USERS.TMSOBB_NAME,
        PrincipalType: 8,
        RequestToJoinLeaveEmailSetting: CONFIG.SHARING.EMAILS.TMSOBB,
        Title: groupName,
      }).catch(first=>console.log(first));
      roles2.forEach(async (role) => {
        var xhr = new XMLHttpRequest();
        xhr.open(
          "POST",
          `${SITE_URL}\/_api\/web\/Lists(@a1)\/GetItemById(@a2)\/UpdateDocumentSharingInfo?@a1=%27${lid}%27&@a2=${itemId}`,
          true
        );
        xhr.setRequestHeader("content-type", "application/json;odata=verbose");
        xhr.setRequestHeader("accept", "application/json;odata=verbose");
        xhr.setRequestHeader(
          "accept-language",
          "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
        );
        xhr.setRequestHeader(
          "X-RequestDigest",
          localStorage.getItem("reqd") ?? "*"
        );
        xhr.withCredentials = true;
        var body = `{\"userRoleAssignments\":[{\"UserId\":\"${personG}\",\"Role\":${role}}],\"validateExistingPermissions\":false,\"additiveMode\":false,\"sendServerManagedNotification\":false,\"customMessage\":null,\"includeAnonymousLinksInNotification\":false,\"propagateAcl\":true}`;
        var aBody = new Uint8Array(body.length);
        for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
        xhr.send(new Blob([aBody]));
      });
          // webObj.lists.getById(lid).items.getById(itemId).roleAssignments.getById(16).groups.getByName(`${personG}`)().catch(first=>console.log(first));;

    });
    personGroup.forEach(async (personG) => {
      roles2.forEach(async (role) => {
        var xhr = new XMLHttpRequest();
        xhr.open(
          "POST",
          `${SITE_URL}\/_api\/web\/Lists(@a1)\/GetItemById(@a2)\/UpdateDocumentSharingInfo?@a1=%27${lid}%27&@a2=${itemId}`,
          true
        );
        xhr.setRequestHeader("content-type", "application/json;odata=verbose");
        xhr.setRequestHeader("accept", "application/json;odata=verbose");
        xhr.setRequestHeader(
          "accept-language",
          "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
        );
        xhr.setRequestHeader(
          "X-RequestDigest",
          localStorage.getItem("reqd") ?? "*"
        );
        xhr.withCredentials = true;
        var body = `{\"userRoleAssignments\":[{\"UserId\":\"${personG}\",\"Role\":${role}}],\"validateExistingPermissions\":false,\"additiveMode\":false,\"sendServerManagedNotification\":false,\"customMessage\":null,\"includeAnonymousLinksInNotification\":true,\"propagateAcl\":true}`;
        var aBody = new Uint8Array(body.length);
        for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
        xhr.send(new Blob([aBody]));
      });
    });
    webObj.lists.getById(lid).items.getById(itemId).roleAssignments.getById(itemId).groups.add({
        AllowMembersEditMembership: true,
        AllowRequestToJoinLeave: false,
        AutoAcceptRequestToJoinLeave: true,
        Description: "This is a sample description for the group.",
        Id: groupId,
        IsHiddenInUI: false,
        LoginName: `i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`,
        OnlyAllowMembersViewMembership: true,
        OwnerTitle: SHAREPOINT_CONSTANTS.USERS.TMSOBB_NAME,
        PrincipalType: 8,
        RequestToJoinLeaveEmailSetting: CONFIG.SHARING.EMAILS.TMSOBB,
        Title: groupName,
      }).catch(first=>console.log(first));
  
    }
    catch(error)
    {
      console.log(error)
    }

  };
  updateroleAssignmentList= () => {

    try{
    principleIds.forEach(async (principleId) => {
      roles.forEach(async (role) => {
        await webObj.lists.getById(lid).roleAssignments.add(principleId, role);
      });
    });
    personGroup.forEach(async (personG) => {
      roles2.forEach(async (role) => {
        var xhr = new XMLHttpRequest();
        xhr.open(
          "POST",
          `${SITE_URL}\/_api\/web\/Lists(@a1)\/UpdateDocumentSharingInfo?@a1=%27${lid}%27&@a2=${itemId}`,
          true
        );
        xhr.setRequestHeader("content-type", "application/json;odata=verbose");
        xhr.setRequestHeader("accept", "application/json;odata=verbose");
        xhr.setRequestHeader(
          "accept-language",
          "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
        );
        xhr.setRequestHeader(
          "X-RequestDigest",
          localStorage.getItem("reqd") ?? "*"
        );
        xhr.withCredentials = true;
        var body = `{\"userRoleAssignments\":[{\"UserId\":\"${personG}\",\"Role\":${role}}],\"validateExistingPermissions\":false,\"additiveMode\":false,\"sendServerManagedNotification\":false,\"customMessage\":null,\"includeAnonymousLinksInNotification\":false,\"propagateAcl\":true}`;
        var aBody = new Uint8Array(body.length);
        for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
        xhr.send(new Blob([aBody]));
      });
    });
    personGroup.forEach(async (personG) => {
      roles2.forEach(async (role) => {
        var xhr = new XMLHttpRequest();
        xhr.open(
          "POST",
          `${SITE_URL}\/_api\/web\/Lists(@a1)\/UpdateDocumentSharingInfo?@a1=%27${lid}%27&@a2=${itemId}`,
          true
        );
        xhr.setRequestHeader("content-type", "application/json;odata=verbose");
        xhr.setRequestHeader("accept", "application/json;odata=verbose");
        xhr.setRequestHeader(
          "accept-language",
          "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
        );
        xhr.setRequestHeader(
          "X-RequestDigest",
          localStorage.getItem("reqd") ?? "*"
        );
        xhr.withCredentials = true;
        var body = `{\"userRoleAssignments\":[{\"UserId\":\"${personG}\",\"Role\":${role}}],\"validateExistingPermissions\":false,\"additiveMode\":false,\"sendServerManagedNotification\":false,\"customMessage\":null,\"includeAnonymousLinksInNotification\":true,\"propagateAcl\":true}`;
        var aBody = new Uint8Array(body.length);
        for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
        xhr.send(new Blob([aBody]));
      });
    });
    webObj.lists.getById(lid).roleAssignments.getById(itemId).groups.add({
        AllowMembersEditMembership: true,
        AllowRequestToJoinLeave: false,
        AutoAcceptRequestToJoinLeave: true,
        Description: "This is a sample description for the group.",
        Id: groupId,
        IsHiddenInUI: false,
        LoginName: `i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`,
        OnlyAllowMembersViewMembership: true,
        OwnerTitle: SHAREPOINT_CONSTANTS.USERS.TMSOBB_NAME,
        PrincipalType: 8,
        RequestToJoinLeaveEmailSetting: CONFIG.SHARING.EMAILS.TMSOBB,
        Title: groupName,
      }).catch(first=>console.log(first));
    webObj.lists.getById(lid).roleAssignments.getById(itemId).groups.getByName("Community Site Members")().catch(first=>console.log(first));;
  
    }
    catch(error)
    {
      console.log(error)
    }

  };
  removeListItemRoleAsignment = () => {
    principleIds.forEach(async (principleId) => {
      roles.forEach(async (role) => {
        await webObj.lists.getById(lid).items.getById(itemId).roleAssignments.remove(principleId, role);
      });
    });
    principleIds.forEach(async (principleId) => {
      roles.forEach(async (role) => {
        await webObj.lists.getById(lid).roleAssignments.remove(principleId, role);
      });
    });
  };
  roleAssignmentListItemsInfo = () => {
    try{
      principleIds.map(id=>{
  dataService.sendRequest(
      "GET",
      `${SITE_URL}/_api/Web/Lists(guid'${lid}')/Items(${itemId})/RoleAssignments/GetByPrincipalId(${id})/RoleDefinitionBindings`
    );
      
    dataService.sendRequest(
      "GET",
      `${SITE_URL}/_api/Web/Lists(guid'${lid}')/RoleAssignments/GetByPrincipalId(${id})/RoleDefinitionBindings`
    );
      
    webObj.lists.getById(lid).items.getById(itemId).roleAssignments().catch(first=>console.log(first))
    webObj.lists.getById(lid).items.getById(itemId).roleAssignments.getById(id).bindings().catch(first=>console.log(first));;

    webObj.lists.getById(lid).items.getById(itemId).roleAssignments.getById(12).groups.getByName("Test")().catch(first=>console.log(first));
    webObj.lists.getById(lid).roleAssignments.getById(id).bindings().catch(first=>console.log(first));;

    webObj.lists.getById(lid).roleAssignments.getById(id).groups.getByName("Test")().catch(first=>console.log(first));
    webObj.lists.getById(lid).roleAssignments.getById(id).groups.getByName("Test")().catch(first=>console.log(first));
    webObj.lists.getById(lid).items.getById(itemId).roleAssignments.getById(id).delete().catch(first=>console.log(first));;

      })

  }
  catch(error)
  {
    console.log(error)
  }

  };
  ListroleInheritance = () => {
    webObj.lists.getById(lid).resetRoleInheritance();
    webObj.lists.getById(lid).breakRoleInheritance(true, true);
  };
  token = async () => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", `${SITE_URL}/_api/contextinfo`, true);
    xhr.setRequestHeader("accept", "application/json;odata=verbose");
    xhr.setRequestHeader("content-type", "application/json;odata=verbose");
    xhr.setRequestHeader(
      "accept-language",
      "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
    );
    xhr.withCredentials = true;

    xhr.onload = function () {
      if (xhr.status >= 200 && xhr.status < 300) {
        // Log the response if the request was successful
        console.log("Response:", xhr.responseText);
        const jsonResponse = JSON.parse(xhr.responseText);
        const formDigestValue =
          jsonResponse.d.GetContextWebInformation.FormDigestValue;
        localStorage.setItem("reqd", formDigestValue);
        console.log("Parsed Response:", formDigestValue);
      } else {
        console.error(
          "Request failed with status:",
          xhr.status,
          xhr.statusText
        );
      }
    };
    xhr.onerror = function () {
      console.error("Request error:", xhr.status, xhr.statusText);
    };

    // Sending an empty body
    xhr.send();
  };
  getAllUserPrincipalIds = async () => {
    try {
      // Get all users from the site
      const users = await webObj.siteUsers();

      // Map to get an array of users with their principal IDs
      const userPrincipalIds = users.map((user) => ({
        LoginName: user.LoginName,
        PrincipalId: user.Id,
        Email: user.Email,
      }));

      console.log("Principal IDs of all users:", userPrincipalIds);

      return userPrincipalIds;
    } catch (error) {
      console.error("Error retrieving user principal IDs:", error);
    }
  };

  share_library = async () => {
    //
    webObj.getFolderByServerRelativePath("/Test/Test1").getShareLink(5).catch((first) => console.log(first));
    webObj.getFolderByServerRelativePath("/Test/Test1").shareWith(`i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`).catch((first) => console.log(first));
    webObj.getFolderByServerRelativePath("/Test/Test1").shareWith(`i:0#.f|membership|${CONFIG.SHARING.EMAILS.ATTACKER}`).catch((first) => console.log(first));
    webObj.shareObject(        `${SITE_URL}/Test/Test1`,        `i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`,        1,        { subject: "Item shared1", body: "hhh" },
        0,        true,        true,        true      ).catch((first) => console.log(first)).catch((first: any) => console.log(first));
    webObj.shareObject(        `${SITE_URL}/Test/Test1`,        `i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`,        1,        { subject: "Item shared1", body: "hhh" },
        0,        true,        true,        true).catch((first) => console.log(first)).catch((first: any) => console.log(first));
    webObj.shareObjectRaw({        url: `${SITE_URL}/Test/Test1`,         peoplePickerInput: JSON.stringify([          { Key: `i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}` }, // User/recipient information
        ]),
        roleValue: "role:1073741826", // Corrected role value for "Edit" permission
        groupId: 0, // Group ID (0 means no group is specified)
        propagateAcl: false, // Whether to propagate the access control list
        sendEmail: true, // Whether to send an email to recipients
        includeAnonymousLinkInEmail: false, // Whether to include an anonymous link in the email
        emailSubject: "Sharing Invitation", // Subject of the email
        emailBody: "You have been granted access to the resource.", // Body of the email
        useSimplifiedRoles: true, // Use simplified roles such as 'view' or 'edit'
      }).then((response) => {
        console.log("Sharing link updated successfully:", response);
      }).catch((error) => {
        console.error("Error updating sharing link:", error);
      });
  };
  getShareId = async () => {
    dataService.getShareLinkIds();
  };
 Test  = () => {
    try {
      // groupIds.forEach((gid: any) => {
      //   dataService.sendRequest(
      //     "GET",      //   );
      // });

      webObj.roleDefinitions().catch(first=>console.log(first))

    } catch (first: any) {
      console.log(first);
    }
  };

  getSiteGroups = async () => {
    console.log(itemId);
    try {
      groupIds.forEach((gid: any) => {
        dataService.sendRequest(
          "GET",
          `${SITE_URL}/_api/web/SiteGroups/GetById(id=${gid})?$expand=Users`
        );
      });
      dataService2._getGroup().catch(first=>console.log(first))
      dataService2.getSIteGroups().catch(first=>console.log(first))
      dataService2.isUserInGroup().catch(first=>console.log(first))
      dataService2._removeUserFromGroup().catch(first=>console.log(first))
      webObj.createDefaultAssociatedGroups("TestSite",`i:0#.f|membership|${CONFIG.SHARING.EMAILS.ATTACKER}`,true,true,`i:0#.f|membership|${CONFIG.SHARING.EMAILS.ATTACKER}`).catch(first=>console.log(first))
      webObj.siteGroups().catch(first=>console.log(first))
      webObj.siteGroups.add({"Title":"SSSSKKKK"}).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId)().catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate")().catch(first=>console.log(first))
      webObj.siteGroups.removeById(groupId).catch(first=>console.log(first))
      webObj.siteGroups.removeByLoginName("AdminPrivate").catch(first=>console.log(first))
      webObj.siteGroups.orderBy("Title")().catch(first=>console.log(first))


      webObj.siteGroups.getById(groupId).setUserAsOwner(11).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).update({Title:"TestUserUpdate1"}).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users().catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.add(`i:0#.f|membership|${CONFIG.SHARING.EMAILS.ATTACKER}`).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`)().catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).delete().catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).update({Title:"TestUserUpdate2"}).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.ATTACKER}`).groups.add({"Title":"SSSSKKKK"}).catch(first=>console.log(first))

      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.getById(groupId)().catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.ATTACKER}`).groups.getById(groupId).setUserAsOwner(11).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.getById(groupId).update({"Title":"SSSSKKKK"}).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.getById(groupId).users().catch(first=>console.log(first))

      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.getByName("AdminPrivate")().catch(first=>console.log(first))

      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.removeById(groupId).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.removeByLoginName(`i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`).catch(first=>console.log(first))

      webObj.siteGroups.getById(groupId).users.getById(userid)().catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getByLoginName(`i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`)().catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.removeById(11).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.removeByLoginName(`i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`).catch(first=>console.log(first))
      
      webObj.siteGroups.getById(groupId).users.getById(userid).delete().catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getById(userid).groups.add({"Title":"SSSSKKKK"}).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getById(userid).update({Title:"Test"}).catch(first=>console.log(first))
      webObj.siteGroups.getById(groupId).users.getById(userid).groups.getByName("AdminPrivate")().catch(first=>console.log(first))

      webObj.siteGroups.getByName("AdminPrivate").setUserAsOwner(11).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").update({Title:"TestUpdate3"}).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users().catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.add(`i:0#.f|membership|${CONFIG.SHARING.EMAILS.ATTACKER}`).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`)().catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).delete().catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).update({Title:"TestUpdate4"}).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.ATTACKER}`).groups.add({"Title":"SSSSKKKK"}).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.ATTACKER}`).groups.getById(groupId)().catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.ATTACKER}`).groups.getById(groupId).setUserAsOwner(11).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.getById(groupId).update({"Title":"SSSSKKKK"}).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.getById(groupId).users().catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.getByName("AdminPrivate")().catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.removeById(groupId).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByEmail(`${CONFIG.SHARING.EMAILS.TMSOBB}`).groups.removeByLoginName(`i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getById(userid)().catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getByLoginName(`i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`)().catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.removeById(11).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.removeByLoginName(`i:0#.f|membership|${CONFIG.SHARING.EMAILS.TMSOBB}`).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getById(userid).delete().catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getById(userid).groups.add({"Title":"SSSSKKKK"}).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getById(userid).update({Title:"Test"}).catch(first=>console.log(first))
      webObj.siteGroups.getByName("AdminPrivate").users.getById(userid).groups.getByName("AdminPrivate")().catch(first=>console.log(first))


 
    } catch (first) {
      console.log(first);
    }
  };

  public render(): React.ReactElement<IHacktheworldProps> {
    return (
      <>
        {
          <Stack gap={15} onClickCapture={this.guardLegacyButton}>
            <ResearchApiPanel context={this.props.context} />
            {this.state.legacyNotice && <MessageBar messageBarType={MessageBarType.warning} onDismiss={() => this.setState({ legacyNotice: "" })}>{this.state.legacyNotice}</MessageBar>}
            {/* <PrimaryButton onClick={() => this.Test()}>TestTestTest1 </PrimaryButton> */}
            <PrimaryButton onClick={() => dataService.getPrincipalIds()}>getPrincipalIds </PrimaryButton>

            <Stack
              wrap={true}
              style={{ textWrap: "wrap" }}
              horizontal={true}
              gap={15}
            >
              <Label style={{ width: "100%" }}>list Item Calls</Label>
              <PrimaryButton
                style={{ textWrap: "wrap" }}
                onClick={() => this.ListItemInformation()}
              >
                Get Item Information
              </PrimaryButton>
              <PrimaryButton
                style={{ textWrap: "wrap" }}
                onClick={() => this.LibraryItemInformation()}
              >
                Get Library Item Information
              </PrimaryButton>
              <PrimaryButton
                style={{ textWrap: "wrap" }}
                onClick={() => this.ListItemVersions()}
              >
                Get ListItem Version Information
              </PrimaryButton>
              <PrimaryButton onClick={() => this.ListItemAddUpdate()}>
                Add,Update,Delete List item
              </PrimaryButton>
              <PrimaryButton onClick={() => this.list_Item_File(  )}>
                List_Item_File
              </PrimaryButton>
              <PrimaryButton onClick={() => this.list_Item_File2(document.getElementById(
                      "fileUploadInput"
                    ) as HTMLInputElement)}>
                list_Item_File2
              </PrimaryButton>
              <PrimaryButton onClick={() => this.List_Item_Attachment()}>
                List_Item_Attachment
              </PrimaryButton>
              <PrimaryButton onClick={() => this.ListItemComments()}>
                Comments
              </PrimaryButton>
              <PrimaryButton onClick={() => this.deleteListItemComments()}>
                delete Comments
              </PrimaryButton>
              <PrimaryButton onClick={() => this.roleInheritance()}>
                Item-roleInheritance{" "}
              </PrimaryButton>
              <PrimaryButton onClick={() => this.ListItemsHARINGInformation()}>
                List Items SHARING Information
              </PrimaryButton>
              <PrimaryButton onClick={() => this.getShareId()}>
                Share_List_Item ids
              </PrimaryButton>
              <Label style={{ width: "100%" }}>Share list Item </Label>
              <PrimaryButton onClick={() => dataService.share_Object()}>
                Share List Obj Endp{" "}
              </PrimaryButton>
              <PrimaryButton onClick={() => this.Share_List_Item_pnp_js()}>
                Share_List_Item pnp js{" "}
              </PrimaryButton>
              <PrimaryButton onClick={() => this.Share_List_Item_XHR()}>
                Share_List_Item XHR
              </PrimaryButton>
              <PrimaryButton
                onClick={() => this.Share_List_Item_WithPassword()}
              >
                Share_List_Item Password
              </PrimaryButton>
              <PrimaryButton onClick={() => this.Share_List_Item_expiry_date()}>
                Share_List_Item wiith expiry date
              </PrimaryButton>
              <PrimaryButton onClick={() => this.Share_List_Item_with_Users()}>
                Share_List_Item wiith Users
              </PrimaryButton>
              <PrimaryButton onClick={() => this.Share_List_Item_role6()}>
                Share_List_Item_Role_6{" "}
              </PrimaryButton>
              <PrimaryButton
                onClick={() => dataService.Update_Share_List_Item_expdate()}
              >
                Update_Share_List_Item_expdate{" "}
              </PrimaryButton>
              <PrimaryButton
                onClick={() => dataService.Update_Share_List_Item_password()}
              >
                Update_Sharing_List_Item_pass{" "}
              </PrimaryButton>
              <PrimaryButton
                onClick={() => dataService.updateSharedLinkWithOutPassword()}
              >
                Update_Sharing_List_Item_Share_Link_NoPassword{" "}
              </PrimaryButton>
              <PrimaryButton onClick={() => this.Delete_Sharing_Link()}>
                Delete_Sharing_Link
              </PrimaryButton>
              <Label style={{ width: "100%" }}>Assign role to list Item </Label>
              <PrimaryButton onClick={() => this.roleAssignmentListItemsInfo()}>
                List Items Role Assignment Info
              </PrimaryButton>
           
              <PrimaryButton
                onClick={() => this.updateroleAssignmentListItems()}
              >
                Update List Items Role Assignment{" "}
              </PrimaryButton>
              <PrimaryButton onClick={() => this.removeListItemRoleAsignment()}>
                List Items Role remove{" "}
              </PrimaryButton>
              {/* <PrimaryButton onClick={this.getAllUserPrincipalIds}>getAllUserPrincipalIds </PrimaryButton> */}

              <Label style={{ width: "100%" }}>List Folder Calls</Label>
            
              <PrimaryButton onClick={() => this.GET_List_Item_Folder_Info()}>
                GET_List_Item_Folder_Info{" "}
              </PrimaryButton>
              <PrimaryButton
                onClick={() => this.Update_Delete_List_Item_Folder_Info()}
              >
                Update_Delete_List_Item_Folder_Info{" "}
              </PrimaryButton>
              <PrimaryButton onClick={() => this.share_list_folder()}>
                share_list_item_folder
              </PrimaryButton>
              <PrimaryButton onClick={() => this.List_Item_Folder_Delete()}>
                List_Item_Folder_Delete{" "}
              </PrimaryButton>
            </Stack>

            <Stack
              wrap={true}
              style={{ textWrap: "wrap" }}
              horizontal={true}
              gap={15}
              padding={15}
            >
              <Label style={{ width: "100%" }}>list Calls</Label>
              <PrimaryButton onClick={() => dataService2.listInfo()}>
                {" "}
                List info
              </PrimaryButton>
              <Label style={{ width: "100%" }}>Share list </Label>
              <PrimaryButton onClick={() => dataService.Share_List()}>
                Share_List
              </PrimaryButton>
              <PrimaryButton onClick={() => dataService.list_role()}>
                Assign role to list
              </PrimaryButton>
              <PrimaryButton onClick={() => dataService.Share_List_WithUsers()}>
                Share_List_WithUsers
              </PrimaryButton>
                 <PrimaryButton
                onClick={() => this.updateroleAssignmentList()}
              >
                Update List Role Assignment{" "}
              </PrimaryButton>
              <PrimaryButton onClick={() => dataService.deleteShareListLinks()}>
                Delete ShareListLinks
              </PrimaryButton>
              <PrimaryButton onClick={() => this.deleteList()}>
                DeleteList
              </PrimaryButton>
            </Stack>

            <Stack style={{ margin: 10 }}></Stack>

            <Stack
              wrap={true}
              style={{ textWrap: "wrap" }}
              horizontal={true}
              gap={15}
            >
              <input type="file" id="fileUploadInput" />
              <PrimaryButton
                onClick={() => dataService2.ListColumnAddUpdateDelete()}
              >
                ListColumn AddUpdateDelete
              </PrimaryButton>
              <PrimaryButton onClick={() => dataService2.ListCruddefaultView()}>
                ListCrud defaultView
              </PrimaryButton>
              <PrimaryButton
                onClick={() => dataService2.ListCrudviewsandgetviews()}
              >
                ListCrud viewsandgetviews
              </PrimaryButton>
              <PrimaryButton onClick={() => dataService2.ListCrudForm()}>
                ListCrudForm
              </PrimaryButton>
              <PrimaryButton onClick={() => dataService2.ListCrudContentType()}>
                {" "}
                ListCrud ContentType
              </PrimaryButton>
              <PrimaryButton
                onClick={() =>
                  dataService2.ListCrudrootFolder(
                    document.getElementById(
                      "fileUploadInput"
                    ) as HTMLInputElement
                  )
                }
              >
                ListCrudrootFolder
              </PrimaryButton>
              <PrimaryButton
                onClick={() =>
                  dataService2.ListCrudrootFolderFiles(
                    document.getElementById(
                      "fileUploadInput"
                    ) as HTMLInputElement
                  )
                }
              >
                ListCrudrootFolder Files
              </PrimaryButton>
              <PrimaryButton
                onClick={() =>
                  dataService2.ListCrudrootFolderSubFolderFiles(
                    document.getElementById(
                      "fileUploadInput"
                    ) as HTMLInputElement
                  )
                }
              >
                ListCrudrootFolderSubFolder Files
              </PrimaryButton>
              <PrimaryButton
                onClick={() =>
                  dataService2.ListCrudrootFolderSubFolder(
                    document.getElementById(
                      "fileUploadInput"
                    ) as HTMLInputElement
                  )
                }
              >
                ListCrud rootFolder SubFolder
              </PrimaryButton>
              <PrimaryButton
                onClick={() => dataService2.ListCrudrootFolderShare()}
              >
                ListCrud rootFolder Share
              </PrimaryButton>
              <PrimaryButton
                onClick={() => dataService2.ListLibraryFileFileversions()}
              >
                ListLibraryFileFileversions
              </PrimaryButton>
            </Stack>
        

              <Stack>
              <Label style={{ width: "100%" }}>Site Calls</Label>
                  <PrimaryButton onClick={this.getSiteGroups}>Site</PrimaryButton>
              </Stack>

          </Stack>
         
        }
      </>
    );
  }
}
