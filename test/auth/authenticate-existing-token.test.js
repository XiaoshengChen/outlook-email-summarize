describe('authenticate tool with async token manager', () => {
  beforeEach(() => {
    jest.resetModules();
  });

  test('returns already authenticated when token manager resolves an access token', async () => {
    jest.doMock('../../config', () => ({
      USE_TEST_MODE: false,
      AUTH_MODE: 'device_code',
      AUTH_CONFIG: {
        clientId: 'client-123',
        scopes: ['offline_access', 'User.Read', 'Mail.Read'],
        tokenStorePath: '/tmp/.outlook-mcp-tokens.json',
        deviceCodeEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode',
        tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
      }
    }));

    jest.doMock('../../auth/token-manager', () => ({
      getAccessToken: jest.fn().mockResolvedValue('access-token-123'),
      clearTokenCache: jest.fn(),
      saveTokenCache: jest.fn(),
      loadTokenCache: jest.fn().mockReturnValue({
        access_token: 'access-token-123',
        expires_at: Date.now() + 3600 * 1000
      })
    }));

    jest.doMock('../../auth/device-code', () => ({
      requestDeviceCode: jest.fn(),
      pollDeviceCodeToken: jest.fn()
    }));

    const { handleAuthenticate } = require('../../auth/tools');
    const result = await handleAuthenticate({});

    expect(result.content[0].text).toBe('Already authenticated and ready.');
  });
});
