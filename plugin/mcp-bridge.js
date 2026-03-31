const path = require('path');
const { Client } = require('@modelcontextprotocol/sdk/client/index.js');
const { StdioClientTransport } = require('@modelcontextprotocol/sdk/client/stdio.js');

const SERVER_ENTRY = path.join(__dirname, '..', 'index.js');
const SERVER_CWD = path.join(__dirname, '..');
const BRIDGE_CLIENT_INFO = {
  name: 'outlook-email-summarize-plugin',
  version: '2.0.0'
};

const bridgeSessions = new Map();

function normalizePluginConfig(rawConfig = {}) {
  return {
    clientId: rawConfig.clientId || '',
    tenantId: rawConfig.tenantId || 'common',
    authMode: rawConfig.authMode || 'device_code',
    readOnlyMode: rawConfig.readOnlyMode !== false
  };
}

function buildServerEnv(rawConfig = {}) {
  const config = normalizePluginConfig(rawConfig);
  const env = {
    ...process.env,
    OUTLOOK_CLIENT_ID: config.clientId,
    OUTLOOK_TENANT_ID: config.tenantId,
    OUTLOOK_AUTH_MODE: config.authMode,
    OUTLOOK_READ_ONLY_MODE: String(config.readOnlyMode)
  };

  console.error(
    `[outlook-email-summarize bridge] buildServerEnv hasClientId=${Boolean(config.clientId)} clientIdLength=${config.clientId.length} tenantId=${config.tenantId} authMode=${config.authMode} readOnlyMode=${config.readOnlyMode}`
  );

  return env;
}

function buildConfigKey(rawConfig = {}) {
  return JSON.stringify(normalizePluginConfig(rawConfig));
}

async function createBridgeSession(rawConfig = {}) {
  const env = buildServerEnv(rawConfig);
  console.error(
    `[outlook-email-summarize bridge] createBridgeSession envSummary=${JSON.stringify({
      hasClientId: Boolean(env.OUTLOOK_CLIENT_ID),
      clientIdLength: String(env.OUTLOOK_CLIENT_ID || '').length,
      tenantId: env.OUTLOOK_TENANT_ID || '',
      authMode: env.OUTLOOK_AUTH_MODE || '',
      readOnlyMode: env.OUTLOOK_READ_ONLY_MODE || ''
    })}`
  );
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [SERVER_ENTRY],
    cwd: SERVER_CWD,
    env,
    stderr: 'pipe'
  });
  const client = new Client(BRIDGE_CLIENT_INFO);

  if (transport.stderr) {
    transport.stderr.on('data', (chunk) => {
      const message = String(chunk).trim();
      if (message) {
        console.error(`[outlook-email-summarize bridge] ${message}`);
      }
    });
  }

  await client.connect(transport);

  return {
    client,
    transport
  };
}

async function getBridgeSession(rawConfig = {}) {
  const configKey = buildConfigKey(rawConfig);
  const existing = bridgeSessions.get(configKey);

  if (existing && existing.ready) {
    return existing.ready;
  }

  const ready = createBridgeSession(rawConfig)
    .then((session) => {
      bridgeSessions.set(configKey, { ready: Promise.resolve(session) });
      return session;
    })
    .catch((error) => {
      bridgeSessions.delete(configKey);
      throw error;
    });

  bridgeSessions.set(configKey, { ready });
  return ready;
}

function shouldResetSession(error) {
  const message = String(error && error.message ? error.message : error).toLowerCase();
  return (
    message.includes('closed') ||
    message.includes('not connected') ||
    message.includes('econnreset') ||
    message.includes('broken pipe')
  );
}

async function disposeBridgeSession(session) {
  if (!session) {
    return;
  }

  try {
    await session.client.close();
  } catch {
    // Ignore child shutdown noise.
  }

  try {
    await session.transport.close();
  } catch {
    // Ignore child shutdown noise.
  }
}

async function callServerTool(rawConfig = {}, toolName, args = {}) {
  const configKey = buildConfigKey(rawConfig);
  let session = await getBridgeSession(rawConfig);

  try {
    return await session.client.callTool({
      name: toolName,
      arguments: args
    });
  } catch (error) {
    if (!shouldResetSession(error)) {
      throw error;
    }

    bridgeSessions.delete(configKey);
    await disposeBridgeSession(session);
    session = await getBridgeSession(rawConfig);

    return session.client.callTool({
      name: toolName,
      arguments: args
    });
  }
}

async function stopAllSessions() {
  const sessions = Array.from(bridgeSessions.values());
  bridgeSessions.clear();

  await Promise.all(
    sessions.map(async (entry) => {
      const session = await entry.ready.catch(() => null);
      await disposeBridgeSession(session);
    })
  );
}

module.exports = {
  buildServerEnv,
  callServerTool,
  normalizePluginConfig,
  stopAllSessions
};
