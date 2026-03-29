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

function createToolFactory(tool) {
  return (ctx = {}) => ({
    name: tool.name,
    description: tool.description,
    parameters: tool.inputSchema,
    async execute(_id, params = {}) {
      return callServerTool(ctx.config || {}, tool.name, params);
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
