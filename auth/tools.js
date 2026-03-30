/**
 * Authentication-related tools for the Outlook MCP server
 */
const fs = require('fs');
const path = require('path');
const config = require('../config');
const tokenManager = require('./token-manager');
const { pollDeviceCodeToken, requestDeviceCode } = require('./device-code');

const PENDING_DEVICE_AUTH_PATH = path.join(
  path.dirname(config.AUTH_CONFIG.tokenStorePath),
  '.outlook-mcp-pending-device-auth.json'
);

function writeJsonFile(filePath, payload) {
  fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), { mode: 0o600 });

  try {
    fs.chmodSync(filePath, 0o600);
  } catch (error) {
    // Windows may ignore chmod semantics. That is fine.
  }
}

function clearPendingDeviceAuth() {
  pendingDeviceAuth = null;

  try {
    if (fs.existsSync(PENDING_DEVICE_AUTH_PATH)) {
      fs.unlinkSync(PENDING_DEVICE_AUTH_PATH);
    }
  } catch (error) {
    console.error('Error clearing pending device auth:', error.message);
  }
}

function loadPendingDeviceAuth() {
  try {
    if (!fs.existsSync(PENDING_DEVICE_AUTH_PATH)) {
      return null;
    }

    const payload = JSON.parse(fs.readFileSync(PENDING_DEVICE_AUTH_PATH, 'utf8'));
    if (!payload || !payload.device_code || !payload.user_code || !payload.verification_uri || !payload.expiresAt) {
      clearPendingDeviceAuth();
      return null;
    }

    if (payload.expiresAt <= Date.now()) {
      clearPendingDeviceAuth();
      return null;
    }

    return payload;
  } catch (error) {
    console.error('Error loading pending device auth:', error.message);
    clearPendingDeviceAuth();
    return null;
  }
}

function savePendingDeviceAuth(payload) {
  try {
    writeJsonFile(PENDING_DEVICE_AUTH_PATH, payload);
    pendingDeviceAuth = payload;
    return true;
  } catch (error) {
    console.error('Error saving pending device auth:', error.message);
    return false;
  }
}

function getPendingDeviceAuth() {
  if (pendingDeviceAuth && pendingDeviceAuth.expiresAt > Date.now()) {
    return pendingDeviceAuth;
  }

  pendingDeviceAuth = loadPendingDeviceAuth();
  return pendingDeviceAuth;
}

let pendingDeviceAuth = loadPendingDeviceAuth();

function formatDeviceCodeMessage(deviceAuth) {
  return [
    'Device code authentication started.',
    `1. Open ${deviceAuth.verification_uri}`,
    `2. Enter code: ${deviceAuth.user_code}`,
    '3. After approving access, run authenticate again or check-auth-status to complete sign-in.'
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
    clearPendingDeviceAuth();
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
  const activePendingAuth = getPendingDeviceAuth();
  if (!activePendingAuth || activePendingAuth.expiresAt <= now) {
    const deviceAuth = await requestDeviceCode({
      clientId: config.AUTH_CONFIG.clientId,
      scopes: config.AUTH_CONFIG.scopes,
      deviceCodeEndpoint: config.AUTH_CONFIG.deviceCodeEndpoint
    });

    const nextPendingAuth = {
      ...deviceAuth,
      expiresAt: now + (deviceAuth.expires_in * 1000)
    };
    savePendingDeviceAuth(nextPendingAuth);

    return {
      content: [{
        type: 'text',
        text: formatDeviceCodeMessage(deviceAuth)
      }]
    };
  }

  const pollResult = await pollDeviceCodeToken({
    clientId: config.AUTH_CONFIG.clientId,
    deviceCode: activePendingAuth.device_code,
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
    clearPendingDeviceAuth();
    return {
      content: [{
        type: 'text',
        text: `Authentication failed: ${pollResult.message}\n\nRun authenticate again to start a fresh device code flow.`
      }]
    };
  }

  const saved = tokenManager.saveTokenCache(pollResult.tokens);
  clearPendingDeviceAuth();
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
    const activePendingAuth = getPendingDeviceAuth();
    if (activePendingAuth) {
      const pollResult = await pollDeviceCodeToken({
        clientId: config.AUTH_CONFIG.clientId,
        deviceCode: activePendingAuth.device_code,
        tokenEndpoint: config.AUTH_CONFIG.tokenEndpoint
      });

      if (pollResult.status === 'authorized') {
        const saved = tokenManager.saveTokenCache(pollResult.tokens);
        clearPendingDeviceAuth();

        if (!saved) {
          throw new Error('Authentication succeeded, but saving tokens failed.');
        }

        const refreshedTokens = tokenManager.loadTokenCache();
        return {
          content: [{
            type: 'text',
            text: `Authenticated and ready. Token expires at ${new Date(refreshedTokens.expires_at).toISOString()}.`
          }]
        };
      }

      if (pollResult.status === 'failed') {
        clearPendingDeviceAuth();
        return {
          content: [{
            type: 'text',
            text: `Not authenticated. Device code authorization ${pollResult.error || 'failed'} and must be restarted.`
          }]
        };
      }

      return {
        content: [{
          type: 'text',
          text: `Authentication pending. Open ${activePendingAuth.verification_uri}, enter code ${activePendingAuth.user_code}, approve access, then run authenticate again or check-auth-status.`
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
