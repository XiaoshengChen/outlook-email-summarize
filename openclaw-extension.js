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

function readConfigFileFallback() {
  const homeDir = process.env.HOME || process.env.USERPROFILE || os.homedir();
  const configPath = process.env.OPENCLAW_CONFIG_PATH || path.join(homeDir, '.openclaw', 'openclaw.json');

  try {
    if (!configPath || !fs.existsSync(configPath)) {
      return {};
    }

    const raw = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    return raw?.plugins?.entries?.['outlook-email-summarize']?.config || {};
  } catch {
    return {};
  }
}

function resolveToolConfig(...contexts) {
  for (const ctx of contexts) {
    if (!ctx || typeof ctx !== 'object') {
      continue;
    }

    const config = (
      ctx.config ||
      ctx.pluginConfig ||
      ctx.entry?.config ||
      ctx.plugin?.config ||
      ctx.extension?.config
    );

    if (config && Object.keys(config).length > 0) {
      return config;
    }
  }

  return (
    readConfigFileFallback()
  );
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
module.exports.readConfigFileFallback = readConfigFileFallback;
module.exports.resolveToolConfig = resolveToolConfig;
