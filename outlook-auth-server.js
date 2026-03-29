#!/usr/bin/env node
const http = require('http');
const url = require('url');
const querystring = require('querystring');
const crypto = require('crypto');

require('dotenv').config();

const config = require('./config');
const TokenStorage = require('./auth/token-storage');

const tokenStorage = new TokenStorage(config.AUTH_CONFIG);
const issuedStates = new Set();

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function renderPage(title, body, color = '#333') {
  return `
    <html>
      <head>
        <title>${escapeHtml(title)}</title>
        <style>
          body { font-family: Arial, sans-serif; max-width: 720px; margin: 48px auto; padding: 0 16px; line-height: 1.5; }
          h1 { color: ${color}; }
          .card { background: #f7f7f7; border: 1px solid #ddd; border-radius: 8px; padding: 16px; }
          code { background: #efefef; padding: 2px 4px; border-radius: 4px; }
        </style>
      </head>
      <body>
        <h1>${escapeHtml(title)}</h1>
        <div class="card">${body}</div>
      </body>
    </html>
  `;
}

function sendHtml(res, statusCode, title, body, color) {
  res.writeHead(statusCode, { 'Content-Type': 'text/html; charset=utf-8' });
  res.end(renderPage(title, body, color));
}

function buildAuthUrl() {
  const state = crypto.randomBytes(16).toString('hex');
  issuedStates.add(state);

  return `${config.AUTH_CONFIG.authEndpoint || 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize'}?${querystring.stringify({
    client_id: config.AUTH_CONFIG.clientId,
    response_type: 'code',
    redirect_uri: config.AUTH_CONFIG.redirectUri,
    scope: config.AUTH_CONFIG.scopes.join(' '),
    response_mode: 'query',
    state
  })}`;
}

const server = http.createServer(async (req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;

  if (pathname === '/') {
    return sendHtml(
      res,
      200,
      'Outlook Loopback Auth',
      '<p>This server only exists for local loopback OAuth debugging.</p><p>For headless deployments, use the <code>authenticate</code> tool and the default device code flow instead.</p>',
      '#0078d4'
    );
  }

  if (pathname === '/auth') {
    if (!config.AUTH_CONFIG.clientId || !config.AUTH_CONFIG.clientSecret) {
      return sendHtml(
        res,
        500,
        'Configuration Error',
        '<p>Set <code>OUTLOOK_CLIENT_ID</code> and <code>OUTLOOK_CLIENT_SECRET</code> before using loopback auth.</p>',
        '#d9534f'
      );
    }

    res.writeHead(302, { Location: buildAuthUrl() });
    res.end();
    return;
  }

  if (pathname === '/auth/callback') {
    const { code, error, error_description: errorDescription, state } = parsedUrl.query;

    if (error) {
      return sendHtml(
        res,
        400,
        'Authorization Failed',
        `<p><strong>Error:</strong> ${escapeHtml(error)}</p><p><strong>Description:</strong> ${escapeHtml(errorDescription || 'No description provided')}</p>`,
        '#d9534f'
      );
    }

    if (!state) {
      return sendHtml(
        res,
        400,
        'Authorization Failed',
        '<p>Missing OAuth state parameter. Start authentication again.</p>',
        '#d9534f'
      );
    }

    if (!issuedStates.has(state)) {
      return sendHtml(
        res,
        400,
        'Authorization Failed',
        '<p>OAuth state mismatch. Start authentication again.</p>',
        '#d9534f'
      );
    }

    issuedStates.delete(state);

    if (!code) {
      return sendHtml(
        res,
        400,
        'Authorization Failed',
        '<p>No authorization code was provided.</p>',
        '#d9534f'
      );
    }

    try {
      await tokenStorage.exchangeCodeForTokens(code);
      return sendHtml(
        res,
        200,
        'Authentication Successful',
        '<p>Microsoft authentication is complete. You can close this window and return to your MCP client.</p>',
        '#2e8b57'
      );
    } catch (authError) {
      return sendHtml(
        res,
        500,
        'Token Exchange Failed',
        `<p>${escapeHtml(authError.message)}</p>`,
        '#d9534f'
      );
    }
  }

  res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end('Not Found');
});

server.listen(config.AUTH_CONFIG.authServerPort, config.AUTH_CONFIG.authServerHost, () => {
  console.log(`Loopback auth server running at http://${config.AUTH_CONFIG.authServerHost}:${config.AUTH_CONFIG.authServerPort}`);
  console.log(`Callback URL: ${config.AUTH_CONFIG.redirectUri}`);
  console.log('This mode is for local debugging only. Headless servers should use device code auth.');
});

process.on('SIGINT', () => process.exit(0));
process.on('SIGTERM', () => process.exit(0));
