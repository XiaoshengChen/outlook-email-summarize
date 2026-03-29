const { buildServerEnv, normalizePluginConfig } = require('../../plugin/mcp-bridge');

describe('mcp bridge config mapping', () => {
  test('normalizes native plugin config to safe defaults', () => {
    expect(normalizePluginConfig({ clientId: 'abc' })).toEqual({
      clientId: 'abc',
      tenantId: 'common',
      authMode: 'device_code',
      readOnlyMode: true
    });
  });

  test('builds child-process env for the standalone MCP server', () => {
    const env = buildServerEnv({
      clientId: 'abc',
      tenantId: 'organizations',
      authMode: 'auth_code_loopback',
      readOnlyMode: false
    });

    expect(env).toEqual(
      expect.objectContaining({
        OUTLOOK_CLIENT_ID: 'abc',
        OUTLOOK_TENANT_ID: 'organizations',
        OUTLOOK_AUTH_MODE: 'auth_code_loopback',
        OUTLOOK_READ_ONLY_MODE: 'false'
      })
    );
  });
});
