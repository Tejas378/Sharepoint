// Constants for SharePoint operations
//
// Every identifier in this file comes from src/config/test-config.json through
// loadTestConfig(). Do not add literal GUIDs, tenant URLs or user identities
// here - edit the JSON and run `npm run config:validate` instead.
//
// Password-like values are not stored here at all: they are resolved at runtime
// from src/config/secrets.ts and only when a password-protected sharing test is
// actually executed.
import { loadTestConfig, TestConfig } from "../config/loader";
import { getSecret } from "../../../config/secrets";

interface LegacyConstants {
  users: {
    primaryName: string;
    primaryLastName: string;
    primaryWorkPhone: string;
    primaryPictureUrl: string;
    primaryPictureTimestamp: string;
  };
  entityIds: {
    objectId: string;
    userGuid: string;
    viewId: string;
    templateFeatureId: string;
    defaultWorkflowId: string;
  };
  sharepoint: {
    siteTitle: string;
    siteTemplate: string;
    welcomePage: string;
    testWelcomePage: string;
    masterUrl: string;
    imageUrl: string;
    language: number;
    uiVersion: number;
  };
  email: {
    sharingSubject: string;
    sharingSubject2: string;
    sharingBody: string;
    sharingInvitationSubject: string;
    sharingInvitationBody: string;
    testSubject: string;
    testBody: string;
  };
  comments: {
    anonymousComment: string;
    mentionComment: string;
    attackerComment: string;
    aaaaComment: string;
  };
  dates: {
    futureExpiration: string;
    pastExpiration: string;
    createdDate: string;
    modifiedDate: string;
    lastModified: string;
    lastUserModified: string;
  };
  changeTokens: {
    listToken: string;
    siteToken: string;
    updatedSiteToken: string;
  };
  listProperties: {
    title: string;
    description: string;
    direction: string;
    draftVersionVisibility: number;
    majorVersionLimit: number;
    itemCount: number;
    baseTemplate: number;
    baseType: number;
  };
  webProperties: {
    allowRssFeeds: boolean;
    enableMinimalDownload: boolean;
    megaMenuEnabled: boolean;
    navAudienceTargetingEnabled: boolean;
    objectCacheEnabled: boolean;
    overwriteTranslationsOnChange: boolean;
    quickLaunchEnabled: boolean;
    recycleBinEnabled: boolean;
    searchScope: number;
    syndicationEnabled: boolean;
    tenantAdminMembersCanShare: number;
    treeViewEnabled: boolean;
    uiVersionConfigurationEnabled: boolean;
  };
  defaults: {
    principalType: number;
    spsUserType: number;
    spsPicturePlaceholderState: number;
    oDataUiVersionString: string;
    attachments: boolean;
  };
}

const TEST_CONFIG: TestConfig = loadTestConfig();

function readLegacyConstants(config: TestConfig): LegacyConstants {
  const legacy = config.testData.constants as LegacyConstants | undefined;
  if (!legacy) {
    throw new Error(
      "test-config.json is missing testData.constants. Add the section described in " +
        "src/config/README.md; constants.ts has no hardcoded fallbacks."
    );
  }
  return legacy;
}

const LEGACY = readLegacyConstants(TEST_CONFIG);

const SITE_URL = TEST_CONFIG.environment.siteUrl;

function password(name: "TEST_PASSWORD" | "TEJAS_PASSWORD" | "OKKK_PASSWORD"): string {
  return getSecret(name);
}

export const SHAREPOINT_CONSTANTS = {
  // Base URLs
  BASE_URL: SITE_URL,
  API_BASE_URL: SITE_URL + "/_api",
  CLIENT_SVC_URL: SITE_URL + "/_vti_bin/client.svc/ProcessQuery",

  // User emails and identities
  USERS: {
    ATTACKER_EMAIL: TEST_CONFIG.principals.users.secondary.upn,
    ATTACKER_LOGIN_NAME: TEST_CONFIG.principals.users.secondary.loginName,
    TMSOBB_EMAIL: TEST_CONFIG.principals.users.primary.upn,
    TMSOBB_LOGIN_NAME: TEST_CONFIG.principals.users.primary.loginName,
    TMSOBB_NAME: LEGACY.users.primaryName,
    TMSOBB_LAST_NAME: LEGACY.users.primaryLastName,
    TMSOBB_WORK_PHONE: LEGACY.users.primaryWorkPhone,
    TMSOBB_PICTURE_URL: LEGACY.users.primaryPictureUrl,
    TMSOBB_PICTURE_TIMESTAMP: LEGACY.users.primaryPictureTimestamp
  },

  // Passwords are runtime-only. Accessing one without supplying it throws a
  // message that explains how to set it; there is no committed default.
  PASSWORDS: {
    get TEST_PASSWORD(): string {
      return password("TEST_PASSWORD");
    },
    get TEJAS_PASSWORD(): string {
      return password("TEJAS_PASSWORD");
    },
    get OKKK_PASSWORD(): string {
      return password("OKKK_PASSWORD");
    }
  },

  // Entity IDs and GUIDs
  ENTITY_IDS: {
    OBJECT_ID: LEGACY.entityIds.objectId,
    USER_GUID: LEGACY.entityIds.userGuid,
    CONTENT_TYPE_ID: TEST_CONFIG.resources.contentTypes.customContentTypeId,
    VIEW_ID: LEGACY.entityIds.viewId,
    TEMPLATE_FEATURE_ID: LEGACY.entityIds.templateFeatureId,
    DEFAULT_WORKFLOW_ID: LEGACY.entityIds.defaultWorkflowId
  },

  // SharePoint specific values
  SHAREPOINT: {
    SITE_TITLE: LEGACY.sharepoint.siteTitle,
    SITE_TEMPLATE: LEGACY.sharepoint.siteTemplate,
    WELCOME_PAGE: LEGACY.sharepoint.welcomePage,
    TEST_WELCOME_PAGE: LEGACY.sharepoint.testWelcomePage,
    MASTER_URL: LEGACY.sharepoint.masterUrl,
    IMAGE_URL: LEGACY.sharepoint.imageUrl,
    LANGUAGE: LEGACY.sharepoint.language,
    UI_VERSION: LEGACY.sharepoint.uiVersion
  },

  // Email subjects and messages
  EMAIL: {
    SHARING_SUBJECT: LEGACY.email.sharingSubject,
    SHARING_SUBJECT_2: LEGACY.email.sharingSubject2,
    SHARING_BODY: LEGACY.email.sharingBody,
    SHARING_INVITATION_SUBJECT: LEGACY.email.sharingInvitationSubject,
    SHARING_INVITATION_BODY: LEGACY.email.sharingInvitationBody,
    TEST_SUBJECT: LEGACY.email.testSubject,
    TEST_BODY: LEGACY.email.testBody
  },

  // Comment text
  COMMENTS: {
    ANONYMOUS_COMMENT: LEGACY.comments.anonymousComment,
    MENTION_COMMENT: LEGACY.comments.mentionComment,
    ATTACKER_COMMENT: LEGACY.comments.attackerComment,
    AAAA_COMMENT: LEGACY.comments.aaaaComment
  },

  // Date formats and timestamps
  DATES: {
    FUTURE_EXPIRATION: LEGACY.dates.futureExpiration,
    PAST_EXPIRATION: LEGACY.dates.pastExpiration,
    CREATED_DATE: LEGACY.dates.createdDate,
    MODIFIED_DATE: LEGACY.dates.modifiedDate,
    LAST_MODIFIED: LEGACY.dates.lastModified,
    LAST_USER_MODIFIED: LEGACY.dates.lastUserModified
  },

  // Change tokens
  CHANGE_TOKENS: {
    LIST_TOKEN: LEGACY.changeTokens.listToken,
    SITE_TOKEN: LEGACY.changeTokens.siteToken,
    UPDATED_SITE_TOKEN: LEGACY.changeTokens.updatedSiteToken
  },

  // List and item properties
  LIST_PROPERTIES: {
    TITLE: LEGACY.listProperties.title,
    DESCRIPTION: LEGACY.listProperties.description,
    DIRECTION: LEGACY.listProperties.direction,
    DRAFT_VERSION_VISIBILITY: LEGACY.listProperties.draftVersionVisibility,
    MAJOR_VERSION_LIMIT: LEGACY.listProperties.majorVersionLimit,
    ITEM_COUNT: LEGACY.listProperties.itemCount,
    BASE_TEMPLATE: LEGACY.listProperties.baseTemplate,
    BASE_TYPE: LEGACY.listProperties.baseType
  },

  // Web properties
  WEB_PROPERTIES: {
    ALLOW_RSS_FEEDS: LEGACY.webProperties.allowRssFeeds,
    ENABLE_MINIMAL_DOWNLOAD: LEGACY.webProperties.enableMinimalDownload,
    MEGA_MENU_ENABLED: LEGACY.webProperties.megaMenuEnabled,
    NAV_AUDIENCE_TARGETING_ENABLED: LEGACY.webProperties.navAudienceTargetingEnabled,
    OBJECT_CACHE_ENABLED: LEGACY.webProperties.objectCacheEnabled,
    OVERWRITE_TRANSLATIONS_ON_CHANGE: LEGACY.webProperties.overwriteTranslationsOnChange,
    QUICK_LAUNCH_ENABLED: LEGACY.webProperties.quickLaunchEnabled,
    RECYCLE_BIN_ENABLED: LEGACY.webProperties.recycleBinEnabled,
    SEARCH_SCOPE: LEGACY.webProperties.searchScope,
    SYNDICATION_ENABLED: LEGACY.webProperties.syndicationEnabled,
    TENANT_ADMIN_MEMBERS_CAN_SHARE: LEGACY.webProperties.tenantAdminMembersCanShare,
    TREE_VIEW_ENABLED: LEGACY.webProperties.treeViewEnabled,
    UI_VERSION_CONFIGURATION_ENABLED: LEGACY.webProperties.uiVersionConfigurationEnabled
  },

  // Default values
  DEFAULTS: {
    PRINCIPAL_TYPE: LEGACY.defaults.principalType,
    SPS_USER_TYPE: LEGACY.defaults.spsUserType,
    SPS_PICTURE_PLACEHOLDER_STATE: LEGACY.defaults.spsPicturePlaceholderState,
    O_DATA_UI_VERSION_STRING: LEGACY.defaults.oDataUiVersionString,
    ATTACHMENTS: LEGACY.defaults.attachments
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
