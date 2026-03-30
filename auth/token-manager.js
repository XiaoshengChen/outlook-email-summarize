/**
 * Token management for Microsoft Graph API authentication
 */
const fs = require('fs');
const https = require('https');
const querystring = require('querystring');
const config = require('../config');

let cachedTokens = null;
let refreshPromise = null;

function normalizeTokens(tokens) {
  if (!tokens || !tokens.access_token) {
    return null;
  }

  const expiresAt = tokens.expires_at || (
    typeof tokens.expires_in === 'number'
      ? Date.now() + (tokens.expires_in * 1000)
      : 0
  );

  return {
    ...tokens,
    expires_at: expiresAt
  };
}

function writeJsonFile(filePath, payload) {
  fs.writeFileSync(filePath, JSON.stringify(payload, null, 2), { mode: 0o600 });

  try {
    fs.chmodSync(filePath, 0o600);
  } catch (error) {
    // Windows may ignore chmod semantics. That is fine.
  }
}

/**
 * Loads authentication tokens from the token file
 * @returns {object|null} - The loaded tokens or null if not available
 */
function readStoredTokens() {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    if (!fs.existsSync(tokenPath)) {
      return null;
    }

    const tokenData = fs.readFileSync(tokenPath, 'utf8');
    return normalizeTokens(JSON.parse(tokenData));
  } catch (error) {
    console.error('Error loading token cache:', error.message);
    return null;
  }
}

/**
 * Loads authentication tokens from the token file
 * @returns {object|null} - The loaded tokens or null if not available
 */
function loadTokenCache() {
  const tokens = readStoredTokens();
  if (!tokens) {
    return null;
  }

  if (Date.now() >= tokens.expires_at) {
    return null;
  }

  cachedTokens = tokens;
  return tokens;
}

/**
 * Saves authentication tokens to the token file
 * @param {object} tokens - The tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveTokenCache(tokens) {
  try {
    const normalizedTokens = normalizeTokens(tokens);
    if (!normalizedTokens) {
      return false;
    }

    writeJsonFile(config.AUTH_CONFIG.tokenStorePath, normalizedTokens);
    cachedTokens = normalizedTokens;
    return true;
  } catch (error) {
    console.error('Error saving token cache:', error.message);
    return false;
  }
}

function clearTokenCache() {
  cachedTokens = null;
  refreshPromise = null;

  try {
    if (fs.existsSync(config.AUTH_CONFIG.tokenStorePath)) {
      fs.unlinkSync(config.AUTH_CONFIG.tokenStorePath);
    }
  } catch (error) {
    console.error('Error clearing token cache:', error.message);
  }
}

/**
 * Gets the current Graph API access token, refreshing it if needed
 * @returns {string|null} - The access token or null if not available
 */
async function getAccessToken() {
  if (cachedTokens && cachedTokens.access_token && Date.now() < cachedTokens.expires_at) {
    return cachedTokens.access_token;
  }

  const storedTokens = readStoredTokens();
  if (!storedTokens) {
    return null;
  }

  cachedTokens = storedTokens;
  if (Date.now() < storedTokens.expires_at) {
    return storedTokens.access_token;
  }

  if (!storedTokens.refresh_token) {
    return null;
  }

  try {
    const refreshedTokens = await refreshAccessToken(storedTokens);
    return refreshedTokens ? refreshedTokens.access_token : null;
  } catch (error) {
    console.error('Error refreshing access token:', error.message);
    return null;
  }
}

/**
 * Gets the current Flow API access token
 * @returns {string|null} - The Flow access token or null if not available
 */
function getFlowAccessToken() {
  const tokens = loadTokenCache();
  if (!tokens) return null;

  if (tokens.flow_access_token && tokens.flow_expires_at && Date.now() < tokens.flow_expires_at) {
    return tokens.flow_access_token;
  }

  return null;
}

function refreshAccessToken(existingTokens) {
  if (!existingTokens || !existingTokens.refresh_token) {
    return Promise.resolve(null);
  }

  if (refreshPromise) {
    return refreshPromise;
  }

  const payload = {
    client_id: config.AUTH_CONFIG.clientId,
    grant_type: 'refresh_token',
    refresh_token: existingTokens.refresh_token,
    scope: config.AUTH_CONFIG.scopes.join(' ')
  };

  if (config.AUTH_CONFIG.clientSecret) {
    payload.client_secret = config.AUTH_CONFIG.clientSecret;
  }

  const body = querystring.stringify(payload);

  refreshPromise = new Promise((resolve, reject) => {
    const request = https.request(config.AUTH_CONFIG.tokenEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    }, (response) => {
      let data = '';

      response.on('data', (chunk) => {
        data += chunk;
      });

      response.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (response.statusCode < 200 || response.statusCode >= 300) {
            reject(new Error(parsed.error_description || 'Token refresh failed.'));
            return;
          }

          const mergedTokens = {
            ...existingTokens,
            ...parsed,
            refresh_token: parsed.refresh_token || existingTokens.refresh_token,
            expires_at: parsed.expires_at || (
              typeof parsed.expires_in === 'number'
                ? Date.now() + (parsed.expires_in * 1000)
                : existingTokens.expires_at
            )
          };

          const saved = saveTokenCache(mergedTokens);
          if (!saved) {
            reject(new Error('Token refresh succeeded, but saving the refreshed token cache failed.'));
            return;
          }

          resolve(cachedTokens);
        } catch (error) {
          reject(new Error(`Failed to parse Microsoft refresh response: ${error.message}`));
        }
      });
    });

    request.on('error', reject);
    request.write(body);
    request.end();
  }).finally(() => {
    refreshPromise = null;
  });

  return refreshPromise;
}

/**
 * Saves Flow API tokens alongside existing Graph tokens
 * @param {object} flowTokens - The Flow tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveFlowTokens(flowTokens) {
  try {
    const existingTokens = loadTokenCache() || {};
    const mergedTokens = {
      ...existingTokens,
      flow_access_token: flowTokens.access_token,
      flow_refresh_token: flowTokens.refresh_token,
      flow_expires_at: flowTokens.expires_at || (Date.now() + (flowTokens.expires_in || 3600) * 1000)
    };

    writeJsonFile(config.AUTH_CONFIG.tokenStorePath, mergedTokens);
    cachedTokens = mergedTokens;
    return true;
  } catch (error) {
    console.error('Error saving Flow tokens:', error.message);
    return false;
  }
}

/**
 * Creates a test access token for use in test mode
 * @returns {object} - The test tokens
 */
function createTestTokens() {
  const testTokens = {
    access_token: `test_access_token_${Date.now()}`,
    refresh_token: `test_refresh_token_${Date.now()}`,
    expires_at: Date.now() + (3600 * 1000)
  };

  saveTokenCache(testTokens);
  return testTokens;
}

module.exports = {
  clearTokenCache,
  createTestTokens,
  getAccessToken,
  getFlowAccessToken,
  loadTokenCache,
  normalizeTokens,
  refreshAccessToken,
  saveFlowTokens,
  saveTokenCache
};
