/**
 * Configuration for Outlook MCP Server
 */
const path = require('path');
const os = require('os');

// Ensure we have a home directory path even if process.env.HOME is undefined
const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir() || '/tmp';

// Public base URL where the OAuth callback is reachable. In a Docker/Traefik
// deployment this is e.g. https://outlook-mcp.example.com. For local stdio
// usage it falls back to the legacy localhost:3333 server.
const publicBaseUrl = (process.env.OAUTH_PUBLIC_BASE_URL || 'http://localhost:3333').replace(/\/+$/, '');

module.exports = {
  // Server information
  SERVER_NAME: "m365-assistant",
  SERVER_VERSION: "2.0.0",

  // Test mode setting
  USE_TEST_MODE: process.env.USE_TEST_MODE === 'true',

  // Authentication configuration
  AUTH_CONFIG: {
    clientId: process.env.MS_CLIENT_ID || process.env.OUTLOOK_CLIENT_ID || '',
    clientSecret: process.env.MS_CLIENT_SECRET || process.env.OUTLOOK_CLIENT_SECRET || '',
    tenantId: process.env.MS_TENANT_ID || 'common',
    authorityHost: (process.env.MS_AUTHORITY_HOST || 'https://login.microsoftonline.com').replace(/\/+$/, ''),
    redirectUri: process.env.MS_REDIRECT_URI || `${publicBaseUrl}/auth/callback`,
    scopes: ['offline_access', 'Mail.Read', 'Mail.ReadWrite', 'Mail.Send', 'User.Read', 'Calendars.Read', 'Calendars.ReadWrite', 'Contacts.Read', 'Files.Read', 'Files.ReadWrite'],
    tokenStorePath: process.env.TOKEN_STORE_PATH || path.join(homeDir, '.outlook-mcp-tokens.json'),
    authServerUrl: publicBaseUrl
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
