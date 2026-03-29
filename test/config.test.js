describe('config defaults', () => {
  const originalEnv = process.env;

  beforeEach(() => {
    jest.resetModules();
    process.env = { ...originalEnv };
    delete process.env.OUTLOOK_SCOPES;
    delete process.env.OUTLOOK_READ_ONLY_MODE;
    delete process.env.OUTLOOK_AUTH_MODE;
    delete process.env.OUTLOOK_CLIENT_ID;
    delete process.env.OUTLOOK_CLIENT_SECRET;
  });

  afterAll(() => {
    process.env = originalEnv;
  });

  test('defaults to read-only device-code auth with minimal scopes', () => {
    const config = require('../config');

    expect(config.READ_ONLY_MODE).toBe(true);
    expect(config.AUTH_MODE).toBe('device_code');
    expect(config.AUTH_CONFIG.scopes).toEqual(['offline_access', 'User.Read', 'Mail.Read']);
    expect(config.AUTH_CONFIG.authServerUrl).toBe('http://127.0.0.1:3333');
    expect(config.AUTH_CONFIG.redirectUri).toBe('http://127.0.0.1:3333/auth/callback');
  });

  test('allows opting out of read-only mode and overriding scopes', () => {
    process.env.OUTLOOK_READ_ONLY_MODE = 'false';
    process.env.OUTLOOK_AUTH_MODE = 'auth_code_loopback';
    process.env.OUTLOOK_SCOPES = 'offline_access User.Read Mail.Read Mail.Send';

    const config = require('../config');

    expect(config.READ_ONLY_MODE).toBe(false);
    expect(config.AUTH_MODE).toBe('auth_code_loopback');
    expect(config.AUTH_CONFIG.scopes).toEqual([
      'offline_access',
      'User.Read',
      'Mail.Read',
      'Mail.Send'
    ]);
  });
});
