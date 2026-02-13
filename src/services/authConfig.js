/**
 * MSAL Configuration for Microsoft 365 Copilot Authentication
 * 
 * Required Permissions (Delegated):
 * - User.Read
 * - Sites.Read.All
 * - Mail.Read
 * - People.Read.All
 * - OnlineMeetingTranscript.Read.All
 * - Chat.Read
 * - ChannelMessage.Read.All
 * - ExternalItem.Read.All
 */

export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_CLIENT_ID || '',
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_TENANT_ID || 'common'}`,
    redirectUri: import.meta.env.VITE_REDIRECT_URI || 'http://localhost:3000',
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
};

// Required scopes for M365 Copilot Chat API
export const loginRequest = {
  scopes: [
    'User.Read',
    'Sites.Read.All',
    'Mail.Read',
    'People.Read.All',
    'OnlineMeetingTranscript.Read.All',
    'Chat.Read',
    'ChannelMessage.Read.All',
    'ExternalItem.Read.All'
  ],
};

// Microsoft Graph endpoints
export const graphConfig = {
  graphMeEndpoint: 'https://graph.microsoft.com/v1.0/me',

  // ✅ Correct endpoint for creating and managing Copilot conversations
  copilotConversationsEndpoint: 'https://graph.microsoft.com/beta/copilot/conversations'
};