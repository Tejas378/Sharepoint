// Constants for SharePoint operations
export const SHAREPOINT_CONSTANTS = {
  // Base URLs
  BASE_URL: "https://tmsobb2.sharepoint.com",
  API_BASE_URL: "https://tmsobb2.sharepoint.com/_api",
  CLIENT_SVC_URL: "https://tmsobb2.sharepoint.com/_vti_bin/client.svc/ProcessQuery",
  
  // User emails and identities
  USERS: {
    ATTACKER_EMAIL: "attacker@tmsobb2.onmicrosoft.com",
    ATTACKER_LOGIN_NAME: "i:0#.f|membership|attacker@tmsobb2.onmicrosoft.com",
    TMSOBB_EMAIL: "tmsobb@tmsobb2.onmicrosoft.com",
    TMSOBB_LOGIN_NAME: "i:0#.f|membership|tmsobb@tmsobb2.onmicrosoft.com",
    TMSOBB_NAME: "TMSOBB",
    TMSOBB_LAST_NAME: "Testt",
    TMSOBB_WORK_PHONE: "9075197067",
    TMSOBB_PICTURE_URL: "https://tmsobb-my.sharepoint.com:443/User%20Photos/Profile%20Pictures/a188d0ef-798b-40f7-b98d-0b300a6d7da3_MThumb.jpg",
    TMSOBB_PICTURE_TIMESTAMP: "63870559229"
  },

  // Passwords
  PASSWORDS: {
    TEST_PASSWORD: "Test@1234",
    TEJAS_PASSWORD: "Tejas@1234",
    OKKK_PASSWORD: "Okkk@1234"
  },

  // Entity IDs and GUIDs
  ENTITY_IDS: {
    OBJECT_ID: "e4d5e997-9b18-4527-887e-5484d8d19ba3",
    USER_GUID: "6a7e2b80-b18d-4bc1-a42c-4fd57c51dcbd",
    CONTENT_TYPE_ID: "0x010A00A37404FF361E30429BA657CF50DB1989",
    VIEW_ID: "3CBA6B32-3C46-428B-B00F-BC6F22AE1FD9",
    TEMPLATE_FEATURE_ID: "00bfea71-de22-43b2-a848-c05709900100",
    DEFAULT_WORKFLOW_ID: "00000000-0000-0000-0000-000000000000"
  },

  // SharePoint specific values
  SHAREPOINT: {
    SITE_TITLE: "Communication site",
    SITE_TEMPLATE: "SITEPAGEPUBLISHING",
    WELCOME_PAGE: "SitePages/Home.aspx",
    TEST_WELCOME_PAGE: "SitePages/test.aspx",
    MASTER_URL: "/_catalogs/masterpage/seattle.master",
    IMAGE_URL: "/_layouts/15/images/itgen.png?rev=47",
    LANGUAGE: 1033,
    UI_VERSION: 15
  },

  // Email subjects and messages
  EMAIL: {
    SHARING_SUBJECT: "tmsobb MSOBB is inviting you to collaborate on three",
    SHARING_SUBJECT_2: "Item shared2",
    SHARING_BODY: "hhh",
    SHARING_INVITATION_SUBJECT: "Sharing Invitation",
    SHARING_INVITATION_BODY: "You have been granted access to the resource.",
    TEST_SUBJECT: "Test Subject",
    TEST_BODY: "Test Body"
  },

  // Comment text
  COMMENTS: {
    ANONYMOUS_COMMENT: "Written comment anonymously even if it is disabled",
    MENTION_COMMENT: "Hii @mention{0} ",
    ATTACKER_COMMENT: "Attacker addeddd",
    AAAA_COMMENT: "aaaa"
  },

  // Date formats and timestamps
  DATES: {
    FUTURE_EXPIRATION: "20500926T000000+0530",
    PAST_EXPIRATION: "20511026T000000+0530",
    CREATED_DATE: "2026-12-24T04:47:23Z",
    MODIFIED_DATE: "2026-12-24T12:02:18Z",
    LAST_MODIFIED: "2026-12-28T07:34:11Z",
    LAST_USER_MODIFIED: "2026-12-27T20:01:52Z"
  },

  // Change tokens
  CHANGE_TOKENS: {
    LIST_TOKEN: "1;3;4af2f3f4-28ee-4aae-82a6-bbd98d0d8d0b;638706161728670000;998946844",
    SITE_TOKEN: "1;2;27341d3a-24dc-48fc-ae4e-0a439cf859fd;638709731707500000;999929707",
    UPDATED_SITE_TOKEN: "1;2;27341d3a-24dc-48fc-ae4e-0a439cf859fd;638750541870370000;1010560322"
  },

  // List and item properties
  LIST_PROPERTIES: {
    TITLE: "Test",
    DESCRIPTION: "",
    DIRECTION: "none",
    DRAFT_VERSION_VISIBILITY: 0,
    MAJOR_VERSION_LIMIT: 50,
    ITEM_COUNT: 8,
    BASE_TEMPLATE: 100,
    BASE_TYPE: 0
  },

  // Web properties
  WEB_PROPERTIES: {
    ALLOW_RSS_FEEDS: true,
    ENABLE_MINIMAL_DOWNLOAD: false,
    MEGA_MENU_ENABLED: true,
    NAV_AUDIENCE_TARGETING_ENABLED: false,
    OBJECT_CACHE_ENABLED: false,
    OVERWRITE_TRANSLATIONS_ON_CHANGE: false,
    QUICK_LAUNCH_ENABLED: true,
    RECYCLE_BIN_ENABLED: true,
    SEARCH_SCOPE: 0,
    SYNDICATION_ENABLED: true,
    TENANT_ADMIN_MEMBERS_CAN_SHARE: 0,
    TREE_VIEW_ENABLED: false,
    UI_VERSION_CONFIGURATION_ENABLED: false
  },

  // Default values
  DEFAULTS: {
    PRINCIPAL_TYPE: 1,
    SPS_USER_TYPE: 0,
    SPS_PICTURE_PLACEHOLDER_STATE: 1,
    O_DATA_UI_VERSION_STRING: "1.0",
    ATTACHMENTS: false
  }
};

// Helper function to create people picker input
export const createPeoplePickerInput = (email: string) => {
  return `[{"Key":"${email}"}]`;
};

// Helper function to create people picker input with entity data
export const createPeoplePickerInputWithEntity = (email: string, objectId: string) => {
  return `[{"Key":"${email}","IsResolved":false,"EntityData":{"ObjectId":"${objectId}"},"PeopleType":"Person","PeopleSubtype":"OrganizationUser"}]`;
};

// Helper function to create mention data
export const createMentionData = (email: string, name: string) => {
  return {
    text: SHAREPOINT_CONSTANTS.COMMENTS.MENTION_COMMENT,
    mentions: [{
      email: email,
      name: name,
      loginName: `i:0#.f|membership|${email}`
    }]
  };
};

// Helper function to create user data object
export const createUserData = (email: string, name: string, lastName: string, workPhone: string) => {
  return {
    Id: 11,
    ServerRedirectedEmbedUri: null,
    ServerRedirectedEmbedUrl: "",
    ContentTypeId: SHAREPOINT_CONSTANTS.ENTITY_IDS.CONTENT_TYPE_ID,
    Title: name,
    OData__ColorTag: null,
    ComplianceAssetId: null,
    Name: `i:0#.f|membership|${email}`,
    EMail: email,
    OtherMail: null,
    UserExpiration: null,
    UserLastDeletionTime: null,
    MobilePhone: null,
    Notes: null,
    SipAddress: email,
    IsSiteAdmin: true,
    Deleted: false,
    UserInfoHidden: false,
    Picture: {
      Description: SHAREPOINT_CONSTANTS.USERS.TMSOBB_PICTURE_URL,
      Url: SHAREPOINT_CONSTANTS.USERS.TMSOBB_PICTURE_URL
    },
    Department: null,
    JobTitle: null,
    FirstName: name,
    LastName: lastName,
    WorkPhone: workPhone,
    UserName: email,
    WebSite: null,
    SPSResponsibility: null,
    Office: null,
    SPSPictureTimestamp: SHAREPOINT_CONSTANTS.USERS.TMSOBB_PICTURE_TIMESTAMP,
    SPSPicturePlaceholderState: SHAREPOINT_CONSTANTS.DEFAULTS.SPS_PICTURE_PLACEHOLDER_STATE,
    SPSPictureExchangeSyncState: null,
    SPSUserType: SHAREPOINT_CONSTANTS.DEFAULTS.SPS_USER_TYPE,
    ID: 11,
    Modified: SHAREPOINT_CONSTANTS.DATES.MODIFIED_DATE,
    Created: "2024-12-23T12:55:38Z",
    AuthorId: 11,
    EditorId: 11,
    OData__UIVersionString: SHAREPOINT_CONSTANTS.DEFAULTS.O_DATA_UI_VERSION_STRING,
    Attachments: SHAREPOINT_CONSTANTS.DEFAULTS.ATTACHMENTS,
    GUID: SHAREPOINT_CONSTANTS.ENTITY_IDS.USER_GUID
  };
};
