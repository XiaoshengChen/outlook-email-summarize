const config = require('../config');

const READ_ONLY_SAFE_TOOLS = new Set([
  'about',
  'authenticate',
  'check-auth-status',
  'list-emails',
  'search-emails',
  'read-email'
]);

function isToolEnabled(toolName) {
  if (config.ENABLE_UNSAFE_TOOLS || !config.READ_ONLY_MODE) {
    return true;
  }

  return READ_ONLY_SAFE_TOOLS.has(toolName);
}

function filterEnabledTools(tools) {
  return tools.filter((tool) => isToolEnabled(tool.name));
}

module.exports = {
  READ_ONLY_SAFE_TOOLS,
  isToolEnabled,
  filterEnabledTools
};
