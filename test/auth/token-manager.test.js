jest.mock('https');

const fs = require('fs');
const os = require('os');
const path = require('path');

describe('token-manager refresh flow', () => {
  const originalEnv = process.env;
  let tempDir;
  let tokenPath;
  let https;
  let mockRequest;

  beforeEach(() => {
    jest.resetModules();
    jest.clearAllMocks();

    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-token-manager-'));
    tokenPath = path.join(tempDir, '.outlook-mcp-tokens.json');

    process.env = {
      ...originalEnv,
      HOME: tempDir,
      OUTLOOK_CLIENT_ID: 'client-123',
      OUTLOOK_TENANT_ID: 'common'
    };
    delete process.env.OUTLOOK_CLIENT_SECRET;
    delete process.env.OUTLOOK_TOKEN_ENDPOINT;

    https = require('https');
    mockRequest = {
      on: jest.fn((event, cb) => {
        if (event === 'error') {
          mockRequest.errorHandler = cb;
        }
        return mockRequest;
      }),
      write: jest.fn(),
      end: jest.fn()
    };

    https.request.mockImplementation((url, options, callback) => {
      mockRequest.callback = callback;
      return mockRequest;
    });
  });

  afterEach(() => {
    process.env = originalEnv;
    fs.rmSync(tempDir, { recursive: true, force: true });
  });

  test('refreshes an expired access token when refresh token is available', async () => {
    fs.writeFileSync(tokenPath, JSON.stringify({
      access_token: 'expired-token',
      refresh_token: 'refresh-token-123',
      expires_at: Date.now() - 60_000
    }));

    const tokenManager = require('../../auth/token-manager');
    const tokenPromise = tokenManager.getAccessToken();

    const refreshResponse = {
      statusCode: 200,
      on: (event, cb) => {
        if (event === 'data') {
          cb(Buffer.from(JSON.stringify({
            access_token: 'fresh-token',
            refresh_token: 'fresh-refresh-token',
            expires_in: 3600
          })));
        }
        if (event === 'end') {
          cb();
        }
      }
    };

    mockRequest.callback(refreshResponse);

    await expect(tokenPromise).resolves.toBe('fresh-token');

    const persisted = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
    expect(persisted.access_token).toBe('fresh-token');
    expect(persisted.refresh_token).toBe('fresh-refresh-token');
    expect(persisted.expires_at).toBeGreaterThan(Date.now());
    expect(mockRequest.write.mock.calls[0][0]).toContain('grant_type=refresh_token');
    expect(mockRequest.write.mock.calls[0][0]).toContain('client_id=client-123');
    expect(mockRequest.write.mock.calls[0][0]).toContain('refresh_token=refresh-token-123');
    expect(mockRequest.write.mock.calls[0][0]).not.toContain('client_secret=');
  });
});
