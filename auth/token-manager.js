/**
 * Token management for Microsoft Graph API authentication
 */
const fs = require('fs');
const config = require('../config');

let cachedTokens = null;

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
function loadTokenCache() {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    if (!fs.existsSync(tokenPath)) {
      return null;
    }

    const tokenData = fs.readFileSync(tokenPath, 'utf8');
    const tokens = normalizeTokens(JSON.parse(tokenData));
    if (!tokens) {
      return null;
    }

    if (Date.now() >= tokens.expires_at) {
      return null;
    }

    cachedTokens = tokens;
    return tokens;
  } catch (error) {
    console.error('Error loading token cache:', error.message);
    return null;
  }
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

  try {
    if (fs.existsSync(config.AUTH_CONFIG.tokenStorePath)) {
      fs.unlinkSync(config.AUTH_CONFIG.tokenStorePath);
    }
  } catch (error) {
    console.error('Error clearing token cache:', error.message);
  }
}

/**
 * Gets the current Graph API access token, loading from cache if necessary
 * @returns {string|null} - The access token or null if not available
 */
function getAccessToken() {
  if (cachedTokens && cachedTokens.access_token && Date.now() < cachedTokens.expires_at) {
    return cachedTokens.access_token;
  }

  const tokens = loadTokenCache();
  return tokens ? tokens.access_token : null;
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
  saveFlowTokens,
  saveTokenCache
};
