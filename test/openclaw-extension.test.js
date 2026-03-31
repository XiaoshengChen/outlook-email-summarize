jest.mock('../plugin/mcp-bridge', () => ({
  callServerTool: jest.fn().mockResolvedValue({
    content: [{ type: 'text', text: 'ok' }]
  }),
  stopAllSessions: jest.fn().mockResolvedValue(undefined)
}));

const { callServerTool, stopAllSessions } = require('../plugin/mcp-bridge');
const plugin = require('../openclaw-extension');

describe('openclaw native extension wrapper', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  test('exports a native plugin entry with a register function', () => {
    expect(plugin.id).toBe('outlook-email-summarize');
    expect(plugin.name).toBe('Outlook Email Summarize');
    expect(typeof plugin.register).toBe('function');
    expect(plugin.configSchema).toEqual(
      expect.objectContaining({
        type: 'object',
        properties: expect.any(Object)
      })
    );
  });

  test('register wires safe outlook tools and cleanup service', async () => {
    const api = {
      registerTool: jest.fn(),
      registerService: jest.fn()
    };

    plugin.register(api);

    expect(api.registerTool).toHaveBeenCalledTimes(6);
    expect(api.registerTool.mock.calls.map((call) => call[1].names[0])).toEqual([
      'about',
      'authenticate',
      'check-auth-status',
      'list-emails',
      'search-emails',
      'read-email'
    ]);

    const authenticateFactory = api.registerTool.mock.calls[1][0];
    const authenticateTool = authenticateFactory({
      config: { clientId: 'client-123', tenantId: 'common' }
    });

    await authenticateTool.execute('tool-call-1', { force: true });

    expect(callServerTool).toHaveBeenCalledWith(
      { clientId: 'client-123', tenantId: 'common' },
      'authenticate',
      { force: true }
    );

    const service = api.registerService.mock.calls[0][0];
    await service.stop();

    expect(stopAllSessions).toHaveBeenCalledTimes(1);
  });

  test('tool execution falls back to plugin-level config when ctx.config is missing', async () => {
    const api = {
      registerTool: jest.fn(),
      registerService: jest.fn()
    };

    plugin.register(api);

    const searchFactory = api.registerTool.mock.calls[4][0];
    const searchTool = searchFactory({
      plugin: {
        config: {
          clientId: 'client-from-plugin',
          tenantId: 'common',
          authMode: 'device_code',
          readOnlyMode: true
        }
      }
    });

    await searchTool.execute('tool-call-2', { from: 'Matt Levine', count: 1 });

    expect(callServerTool).toHaveBeenCalledWith(
      {
        clientId: 'client-from-plugin',
        tenantId: 'common',
        authMode: 'device_code',
        readOnlyMode: true
      },
      'search-emails',
      { from: 'Matt Levine', count: 1 }
    );
  });
});
