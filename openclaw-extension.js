const fs = require('fs');
const os = require('os');
const path = require('path');
const { authTools } = require('./auth');
const { emailTools } = require('./email');
const { READ_ONLY_SAFE_TOOLS } = require('./utils/feature-gates');
const { callServerTool, stopAllSessions } = require('./plugin/mcp-bridge');

const SAFE_PLUGIN_TOOLS = [...authTools, ...emailTools].filter((tool) =>
  READ_ONLY_SAFE_TOOLS.has(tool.name)
);

const configSchema = {
  type: 'object',
  additionalProperties: false,
  properties: {
    clientId: {
      type: 'string',
      minLength: 1
    },
    tenantId: {
      type: 'string',
      default: 'common'
    },
    authMode: {
      type: 'string',
      enum: ['device_code', 'auth_code_loopback'],
      default: 'device_code'
    },
    readOnlyMode: {
      type: 'boolean',
      default: true
    }
  },
  required: ['clientId']
};

function summarizeConfigForLogs(config = {}) {
  const clientId = typeof config.clientId === 'string' ? config.clientId : '';

  return {
    hasClientId: clientId.length > 0,
    clientIdLength: clientId.length,
    tenantId: config.tenantId || '',
    authMode: config.authMode || '',
    readOnlyMode: config.readOnlyMode
  };
}

function hasUsableClientId(config) {
  return Boolean(
    config &&
    typeof config.clientId === 'string' &&
    config.clientId.trim().length > 0
  );
}

function readConfigFileFallback() {
  const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir();
  const configPath = process.env.OPENCLAW_CONFIG_PATH || path.join(homeDir, '.openclaw', 'openclaw.json');

  try {
    if (!configPath || !fs.existsSync(configPath)) {
      return {};
    }

    const raw = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    const fallbackConfig = raw?.plugins?.entries?.['outlook-email-summarize']?.config || {};
    console.error(
      `[outlook-email-summarize plugin] readConfigFileFallback path=${configPath} summary=${JSON.stringify(summarizeConfigForLogs(fallbackConfig))}`
    );
    return fallbackConfig;
  } catch {
    return {};
  }
}

function resolveToolConfig(...contexts) {
  const candidates = [
    ['ctx.config', (ctx) => ctx.config],
    ['ctx.pluginConfig', (ctx) => ctx.pluginConfig],
    ['ctx.entry.config', (ctx) => ctx.entry?.config],
    ['ctx.plugin.config', (ctx) => ctx.plugin?.config],
    ['ctx.extension.config', (ctx) => ctx.extension?.config]
  ];

  for (const ctx of contexts) {
    if (!ctx || typeof ctx !== 'object') {
      continue;
    }

    for (const [source, getter] of candidates) {
      const config = getter(ctx);
      if (hasUsableClientId(config)) {
        console.error(
          `[outlook-email-summarize plugin] resolveToolConfig source=${source} summary=${JSON.stringify(summarizeConfigForLogs(config))}`
        );
        return config;
      }
    }
  }

  const fallbackConfig = readConfigFileFallback();
  console.error(
    `[outlook-email-summarize plugin] resolveToolConfig source=file-fallback summary=${JSON.stringify(summarizeConfigForLogs(fallbackConfig))}`
  );
  return fallbackConfig;
}

function createToolFactory(tool) {
  return (ctx = {}) => ({
    name: tool.name,
    description: tool.description,
    parameters: tool.inputSchema,
    async execute(_id, params = {}, runtimeCtx = {}) {
      return callServerTool(resolveToolConfig(runtimeCtx, ctx), tool.name, params);
    }
  });
}

function register(api) {
  for (const tool of SAFE_PLUGIN_TOOLS) {
    api.registerTool(createToolFactory(tool), {
      names: [tool.name]
    });
  }

  api.registerService({
    id: 'outlook-email-summarize.bridge',
    start: () => undefined,
    stop: () => stopAllSessions()
  });
}

const pluginEntry = {
  id: 'outlook-email-summarize',
  name: 'Outlook Email Summarize',
  description: 'Safe Outlook inbox summarizer for OpenClaw-China using a bridged MCP server.',
  configSchema,
  register
};

module.exports = pluginEntry;
module.exports.default = pluginEntry;
module.exports.hasUsableClientId = hasUsableClientId;
module.exports.readConfigFileFallback = readConfigFileFallback;
module.exports.resolveToolConfig = resolveToolConfig;
