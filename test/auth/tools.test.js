const fs = require('fs');
const os = require('os');
const path = require('path');

describe('auth tools device code persistence', () => {
  let tempDir;

  beforeEach(() => {
    jest.resetModules();
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-auth-tools-'));
  });

  afterEach(() => {
    fs.rmSync(tempDir, { recursive: true, force: true });
  });

  function loadToolsWithMocks({ requestDeviceCode, pollDeviceCodeToken, loadTokenCache, saveTokenCache, getAccessToken, clearTokenCache }) {
    jest.doMock('../../config', () => ({
      USE_TEST_MODE: false,
      AUTH_MODE: 'device_code',
      AUTH_CONFIG: {
        clientId: 'client-123',
        scopes: ['offline_access', 'User.Read', 'Mail.Read'],
        tokenStorePath: path.join(tempDir, '.outlook-mcp-tokens.json'),
        deviceCodeEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode',
        tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
      }
    }));

    jest.doMock('../../auth/device-code', () => ({
      requestDeviceCode,
      pollDeviceCodeToken
    }));

    jest.doMock('../../auth/token-manager', () => ({
      loadTokenCache,
      saveTokenCache,
      getAccessToken,
      clearTokenCache
    }));

    return require('../../auth/tools');
  }

  test('authenticate resumes a persisted pending device flow after module reload', async () => {
    const requestDeviceCode = jest.fn().mockResolvedValue({
      device_code: 'device-code-1',
      user_code: 'ABCD-EFGH',
      verification_uri: 'https://microsoft.com/devicelogin',
      expires_in: 900,
      interval: 5
    });
    const pollDeviceCodeToken = jest.fn().mockResolvedValue({
      status: 'authorized',
      tokens: {
        access_token: 'token-123',
        expires_in: 3600
      }
    });
    const saveTokenCache = jest.fn().mockReturnValue(true);
    const loadTokenCache = jest.fn().mockReturnValue(null);
    const getAccessToken = jest.fn().mockReturnValue(null);
    const clearTokenCache = jest.fn();

    let tools = loadToolsWithMocks({
      requestDeviceCode,
      pollDeviceCodeToken,
      loadTokenCache,
      saveTokenCache,
      getAccessToken,
      clearTokenCache
    });

    const firstResult = await tools.handleAuthenticate({});
    expect(firstResult.content[0].text).toContain('Enter code: ABCD-EFGH');
    expect(requestDeviceCode).toHaveBeenCalledTimes(1);
    expect(pollDeviceCodeToken).not.toHaveBeenCalled();

    jest.resetModules();

    tools = loadToolsWithMocks({
      requestDeviceCode,
      pollDeviceCodeToken,
      loadTokenCache,
      saveTokenCache,
      getAccessToken,
      clearTokenCache
    });

    const secondResult = await tools.handleAuthenticate({});
    expect(secondResult.content[0].text).toContain('Authentication successful');
    expect(requestDeviceCode).toHaveBeenCalledTimes(1);
    expect(pollDeviceCodeToken).toHaveBeenCalledWith({
      clientId: 'client-123',
      deviceCode: 'device-code-1',
      tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    });
    expect(saveTokenCache).toHaveBeenCalledWith({
      access_token: 'token-123',
      expires_in: 3600
    });
  });

  test('check-auth-status finalizes a completed device code flow without a second authenticate call', async () => {
    const requestDeviceCode = jest.fn().mockResolvedValue({
      device_code: 'device-code-2',
      user_code: 'WXYZ-1234',
      verification_uri: 'https://microsoft.com/devicelogin',
      expires_in: 900,
      interval: 5
    });
    const pollDeviceCodeToken = jest.fn().mockResolvedValue({
      status: 'authorized',
      tokens: {
        access_token: 'token-456',
        expires_at: Date.now() + 3600 * 1000
      }
    });

    const tokenState = { current: null };
    const saveTokenCache = jest.fn().mockImplementation((tokens) => {
      tokenState.current = tokens;
      return true;
    });
    const loadTokenCache = jest.fn().mockImplementation(() => tokenState.current);
    const getAccessToken = jest.fn().mockImplementation(() => tokenState.current && tokenState.current.access_token);
    const clearTokenCache = jest.fn();

    const tools = loadToolsWithMocks({
      requestDeviceCode,
      pollDeviceCodeToken,
      loadTokenCache,
      saveTokenCache,
      getAccessToken,
      clearTokenCache
    });

    await tools.handleAuthenticate({});
    const statusResult = await tools.handleCheckAuthStatus();

    expect(statusResult.content[0].text).toContain('Authenticated and ready.');
    expect(pollDeviceCodeToken).toHaveBeenCalledWith({
      clientId: 'client-123',
      deviceCode: 'device-code-2',
      tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    });
    expect(saveTokenCache).toHaveBeenCalled();
  });
});
