/**
 * Configuration for Outlook MCP Server
 */
const path = require('path');
const os = require('os');

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';
const DEFAULT_READ_ONLY_SCOPES = ['offline_access', 'User.Read', 'Mail.Read'];
const DEFAULT_FULL_SCOPES = [
  'offline_access',
  'User.Read',
  'Mail.Read',
  'Mail.ReadWrite',
  'Mail.Send',
  'Calendars.Read',
  'Calendars.ReadWrite',
  'Files.Read',
  'Files.ReadWrite'
];

function parseBoolean(value, defaultValue) {
  if (value === undefined) {
    return defaultValue;
  }

  return String(value).toLowerCase() === 'true';
}

function parseScopes(rawValue, fallbackScopes) {
  const scopeString = typeof rawValue === 'string' ? rawValue.trim() : '';
  if (!scopeString) {
    return fallbackScopes;
  }

  return scopeString.split(/\s+/).filter(Boolean);
}

const READ_ONLY_MODE = parseBoolean(process.env.OUTLOOK_READ_ONLY_MODE, true);
const AUTH_MODE = process.env.OUTLOOK_AUTH_MODE || 'device_code';
const ENABLE_UNSAFE_TOOLS = parseBoolean(process.env.OUTLOOK_ENABLE_UNSAFE_TOOLS, false);
const AUTH_HOST = process.env.OUTLOOK_AUTH_HOST || '127.0.0.1';
const AUTH_PORT = Number(process.env.OUTLOOK_AUTH_PORT || 3333);
const defaultScopes = READ_ONLY_MODE ? DEFAULT_READ_ONLY_SCOPES : DEFAULT_FULL_SCOPES;
const TENANT_ID = process.env.OUTLOOK_TENANT_ID || process.env.MS_TENANT_ID || 'common';
const AUTHORITY_HOST = (process.env.OUTLOOK_AUTHORITY_HOST || process.env.MS_AUTHORITY_HOST || 'https://login.microsoftonline.com').replace(/\/+$/, '');
const configuredScopes = parseScopes(
  process.env.OUTLOOK_SCOPES || process.env.MS_SCOPES,
  defaultScopes
);

module.exports = {
  // Server information
  SERVER_NAME: "m365-assistant",
  SERVER_VERSION: "2.0.0",
  
  // Test mode setting
  USE_TEST_MODE: process.env.USE_TEST_MODE === 'true',
  READ_ONLY_MODE,
  AUTH_MODE,
  ENABLE_UNSAFE_TOOLS,
  
  // Authentication configuration
  AUTH_CONFIG: {
    clientId: process.env.OUTLOOK_CLIENT_ID || process.env.MS_CLIENT_ID || '',
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET || process.env.MS_CLIENT_SECRET || '',
    redirectUri: process.env.OUTLOOK_REDIRECT_URI || process.env.MS_REDIRECT_URI || `http://${AUTH_HOST}:${AUTH_PORT}/auth/callback`,
    scopes: configuredScopes,
    tokenStorePath: path.join(homeDir, '.outlook-mcp-tokens.json'),
    authServerUrl: process.env.OUTLOOK_AUTH_SERVER_URL || `http://${AUTH_HOST}:${AUTH_PORT}`,
    authServerHost: AUTH_HOST,
    authServerPort: AUTH_PORT,
    authEndpoint: process.env.OUTLOOK_AUTH_ENDPOINT || `${AUTHORITY_HOST}/${TENANT_ID}/oauth2/v2.0/authorize`,
    deviceCodeEndpoint: process.env.OUTLOOK_DEVICE_CODE_ENDPOINT || `${AUTHORITY_HOST}/${TENANT_ID}/oauth2/v2.0/devicecode`,
    tokenEndpoint: process.env.OUTLOOK_TOKEN_ENDPOINT || `${AUTHORITY_HOST}/${TENANT_ID}/oauth2/v2.0/token`
  },
  
  // Microsoft Graph API
  GRAPH_API_ENDPOINT: 'https://graph.microsoft.com/v1.0/',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,start,end,location,bodyPreview,isAllDay,recurrence,attendees',

  // Email constants
  EMAIL_SELECT_FIELDS: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead',
  EMAIL_DETAIL_FIELDS: 'id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead,internetMessageHeaders',
  
  // Calendar constants
  CALENDAR_SELECT_FIELDS: 'id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,isCancelled',
  
  // Pagination
  DEFAULT_PAGE_SIZE: 25,
  MAX_RESULT_COUNT: 50,

  // Timezone
  DEFAULT_TIMEZONE: "Central European Standard Time",

  // OneDrive constants
  ONEDRIVE_SELECT_FIELDS: 'id,name,size,lastModifiedDateTime,webUrl,folder,file,parentReference',
  ONEDRIVE_UPLOAD_THRESHOLD: 4 * 1024 * 1024, // 4MB - files larger than this need chunked upload

  // Power Automate / Flow constants
  FLOW_API_ENDPOINT: 'https://api.flow.microsoft.com',
  FLOW_SCOPE: 'https://service.flow.microsoft.com/.default',
};
