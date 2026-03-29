describe('feature gates', () => {
  const originalEnv = process.env;

  beforeEach(() => {
    jest.resetModules();
    process.env = { ...originalEnv };
    delete process.env.OUTLOOK_READ_ONLY_MODE;
    delete process.env.OUTLOOK_ENABLE_UNSAFE_TOOLS;
  });

  afterAll(() => {
    process.env = originalEnv;
  });

  test('exposes only auth and email read tools by default', () => {
    const { filterEnabledTools } = require('../../utils/feature-gates');

    const tools = [
      { name: 'about' },
      { name: 'authenticate' },
      { name: 'check-auth-status' },
      { name: 'list-emails' },
      { name: 'search-emails' },
      { name: 'read-email' },
      { name: 'send-email' },
      { name: 'create-folder' },
      { name: 'list-events' },
      { name: 'onedrive-list' }
    ];

    expect(filterEnabledTools(tools).map((tool) => tool.name)).toEqual([
      'about',
      'authenticate',
      'check-auth-status',
      'list-emails',
      'search-emails',
      'read-email'
    ]);
  });

  test('allows the full toolset when unsafe tools are explicitly enabled', () => {
    process.env.OUTLOOK_ENABLE_UNSAFE_TOOLS = 'true';

    const { filterEnabledTools } = require('../../utils/feature-gates');
    const tools = [{ name: 'read-email' }, { name: 'send-email' }, { name: 'list-events' }];

    expect(filterEnabledTools(tools).map((tool) => tool.name)).toEqual([
      'read-email',
      'send-email',
      'list-events'
    ]);
  });
});
