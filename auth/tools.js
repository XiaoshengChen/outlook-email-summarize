/**
 * Authentication-related tools for the Outlook MCP server
 */
const config = require('../config');
const tokenManager = require('./token-manager');
const { pollDeviceCodeToken, requestDeviceCode } = require('./device-code');

let pendingDeviceAuth = null;

function formatDeviceCodeMessage(deviceAuth) {
  return [
    'Device code authentication started.',
    `1. Open ${deviceAuth.verification_uri}`,
    `2. Enter code: ${deviceAuth.user_code}`,
    '3. After approving access, run the authenticate tool again to complete sign-in.'
  ].join('\n');
}

function buildLoopbackAuthUrl() {
  return `${config.AUTH_CONFIG.authServerUrl}/auth?client_id=${config.AUTH_CONFIG.clientId}`;
}

/**
 * About tool handler
 * @returns {object} - MCP response
 */
async function handleAbout() {
  return {
    content: [{
      type: "text",
      text: `M365 Assistant MCP Server v${config.SERVER_VERSION}\n\nProvides access to Microsoft 365 services through Microsoft Graph API:\n- Outlook (email, calendar, folders, rules)\n- OneDrive (files, folders, sharing)\n- Power Automate (flows, environments, runs)\n\nModular architecture for improved maintainability.`
    }]
  };
}

/**
 * Authentication tool handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleAuthenticate(args) {
  const force = args && args.force === true;

  if (force) {
    pendingDeviceAuth = null;
    tokenManager.clearTokenCache();
  }
  
  // For test mode, create a test token
  if (config.USE_TEST_MODE) {
    // Create a test token with a 1-hour expiry
    tokenManager.createTestTokens();
    
    return {
      content: [{
        type: "text",
        text: 'Successfully authenticated with Microsoft Graph API (test mode)'
      }]
    };
  }

  const existingToken = tokenManager.getAccessToken();
  if (existingToken && !force) {
    return {
      content: [{
        type: 'text',
        text: 'Already authenticated and ready.'
      }]
    };
  }

  if (config.AUTH_MODE === 'auth_code_loopback') {
    if (!config.AUTH_CONFIG.clientId) {
      return {
        content: [{
          type: 'text',
          text: 'Authentication is not configured. Set OUTLOOK_CLIENT_ID and OUTLOOK_CLIENT_SECRET first.'
        }]
      };
    }

    const authUrl = buildLoopbackAuthUrl();

    return {
      content: [{
        type: 'text',
        text: `Loopback authentication required. Visit ${authUrl} in a browser running on this machine, complete Microsoft sign-in, then come back and use your email tools.`
      }]
    };
  }

  if (!config.AUTH_CONFIG.clientId) {
    return {
      content: [{
        type: 'text',
        text: 'Authentication is not configured. Set OUTLOOK_CLIENT_ID first.'
      }]
    };
  }

  const now = Date.now();
  if (!pendingDeviceAuth || pendingDeviceAuth.expiresAt <= now) {
    const deviceAuth = await requestDeviceCode({
      clientId: config.AUTH_CONFIG.clientId,
      scopes: config.AUTH_CONFIG.scopes,
      deviceCodeEndpoint: config.AUTH_CONFIG.deviceCodeEndpoint
    });

    pendingDeviceAuth = {
      ...deviceAuth,
      expiresAt: now + (deviceAuth.expires_in * 1000)
    };

    return {
      content: [{
        type: 'text',
        text: formatDeviceCodeMessage(deviceAuth)
      }]
    };
  }

  const pollResult = await pollDeviceCodeToken({
    clientId: config.AUTH_CONFIG.clientId,
    deviceCode: pendingDeviceAuth.device_code,
    tokenEndpoint: config.AUTH_CONFIG.tokenEndpoint
  });

  if (pollResult.status === 'pending' || pollResult.status === 'slow_down') {
    return {
      content: [{
        type: 'text',
        text: `${pollResult.message || 'Authorization is still pending.'}\n\nApprove the device code, then run authenticate again.`
      }]
    };
  }

  if (pollResult.status === 'failed') {
    pendingDeviceAuth = null;
    return {
      content: [{
        type: 'text',
        text: `Authentication failed: ${pollResult.message}\n\nRun authenticate again to start a fresh device code flow.`
      }]
    };
  }

  const saved = tokenManager.saveTokenCache(pollResult.tokens);
  pendingDeviceAuth = null;
  if (!saved) {
    throw new Error('Authentication succeeded, but saving tokens failed.');
  }

  return {
    content: [{
      type: 'text',
      text: 'Authentication successful. Outlook email read tools are ready.'
    }]
  };
}

/**
 * Check authentication status tool handler
 * @returns {object} - MCP response
 */
async function handleCheckAuthStatus() {
  const tokens = tokenManager.loadTokenCache();

  if (!tokens || !tokens.access_token) {
    if (pendingDeviceAuth && pendingDeviceAuth.expiresAt > Date.now()) {
      return {
        content: [{
          type: 'text',
          text: `Authentication pending. Open ${pendingDeviceAuth.verification_uri}, enter code ${pendingDeviceAuth.user_code}, approve access, then run authenticate again.`
        }]
      };
    }

    return {
      content: [{ type: "text", text: "Not authenticated" }]
    };
  }

  return {
    content: [{ type: "text", text: `Authenticated and ready. Token expires at ${new Date(tokens.expires_at).toISOString()}.` }]
  };
}

// Tool definitions
const authTools = [
  {
    name: "about",
    description: "Returns information about this M365 Assistant server",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleAbout
  },
  {
    name: "authenticate",
    description: "Authenticate with Microsoft Graph API to access Outlook data",
    inputSchema: {
      type: "object",
      properties: {
        force: {
          type: "boolean",
          description: "Force re-authentication even if already authenticated"
        }
      },
      required: []
    },
    handler: handleAuthenticate
  },
  {
    name: "check-auth-status",
    description: "Check the current authentication status with Microsoft Graph API",
    inputSchema: {
      type: "object",
      properties: {},
      required: []
    },
    handler: handleCheckAuthStatus
  }
];

module.exports = {
  authTools,
  handleAbout,
  handleAuthenticate,
  handleCheckAuthStatus
};
