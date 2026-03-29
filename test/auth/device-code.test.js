jest.mock('https');

describe('device code auth flow', () => {
  let https;
  let requestDeviceCode;
  let pollDeviceCodeToken;
  let mockRequest;

  beforeEach(() => {
    jest.resetModules();
    jest.clearAllMocks();
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

    ({ requestDeviceCode, pollDeviceCodeToken } = require('../../auth/device-code'));
  });

  test('requests a device code with the configured scopes', async () => {
    const promise = requestDeviceCode({
      clientId: 'test-client-id',
      scopes: ['offline_access', 'User.Read', 'Mail.Read'],
      deviceCodeEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/devicecode'
    });

    const mockResponse = {
      statusCode: 200,
      on: (event, cb) => {
        if (event === 'data') {
          cb(Buffer.from(JSON.stringify({
            device_code: 'device-code',
            user_code: 'ABCD-EFGH',
            verification_uri: 'https://microsoft.com/devicelogin',
            expires_in: 900,
            interval: 5
          })));
        }
        if (event === 'end') {
          cb();
        }
      }
    };

    mockRequest.callback(mockResponse);

    await expect(promise).resolves.toEqual(expect.objectContaining({
      device_code: 'device-code',
      user_code: 'ABCD-EFGH'
    }));
    expect(mockRequest.write.mock.calls[0][0]).toContain('scope=offline_access%20User.Read%20Mail.Read');
  });

  test('returns authorization pending status while the user has not approved yet', async () => {
    const promise = pollDeviceCodeToken({
      clientId: 'test-client-id',
      deviceCode: 'device-code',
      tokenEndpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    });

    const pendingResponse = {
      statusCode: 400,
      on: (event, cb) => {
        if (event === 'data') {
          cb(Buffer.from(JSON.stringify({
            error: 'authorization_pending',
            error_description: 'Waiting for user approval'
          })));
        }
        if (event === 'end') {
          cb();
        }
      }
    };

    mockRequest.callback(pendingResponse);

    await expect(promise).resolves.toEqual(expect.objectContaining({
      status: 'pending'
    }));
  });
});
