# Outlook Email Summarize MCP

An MCP server fork focused on one job: securely reading Outlook email for inbox triage and summarization.

The upstream project tried to be a Microsoft 365 swiss-army knife. That is how you end up giving a summary bot write access to mail, calendar, files, and automation flows. This fork flips the default:

- `device_code` auth is the default, so headless servers actually work
- `Mail.Read + User.Read + offline_access` are the default scopes
- only safe email-read tools are exposed by default
- loopback OAuth is still available, but only for local debugging on `127.0.0.1`

## Default Tool Surface

These tools stay enabled in the default safe mode:

- `about`
- `authenticate`
- `check-auth-status`
- `list-emails`
- `search-emails`
- `read-email`

Everything else is hidden unless you explicitly opt into unsafe mode.

## Quick Start

1. Install dependencies

```bash
npm install
```

2. Copy `.env.example` to `.env` and fill in your Microsoft app credentials.

3. Start the MCP server in your client.

```bash
npm start
```

4. Run the `authenticate` tool.

5. Open the Microsoft device login URL, enter the code, approve access, then run `authenticate` again once.

6. Use `list-emails`, `search-emails`, or `read-email`.

## Microsoft App Registration

Use Microsoft Entra ID and create a public/delegated app for Microsoft Graph.

Recommended delegated permissions for this fork:

- `offline_access`
- `User.Read`
- `Mail.Read`

Do not add `Mail.ReadWrite`, `Mail.Send`, `Files.ReadWrite`, or calendar write scopes unless you consciously want a bigger blast radius.

If you need personal Outlook accounts such as `outlook.com` or `hotmail.com`, choose an app registration that supports both organizational and personal Microsoft accounts.

## Configuration

### Minimal `.env`

```bash
OUTLOOK_CLIENT_ID=your-client-id
OUTLOOK_TENANT_ID=common
OUTLOOK_READ_ONLY_MODE=true
OUTLOOK_AUTH_MODE=device_code
USE_TEST_MODE=false
```

`OUTLOOK_CLIENT_SECRET` is optional for the default device code flow. It is only required if you choose loopback auth with the local callback server.

### Important Environment Variables

- `OUTLOOK_CLIENT_ID`: Microsoft application client ID
- `OUTLOOK_CLIENT_SECRET`: only needed for loopback auth code flow
- `OUTLOOK_TENANT_ID`: tenant GUID or `common`
- `OUTLOOK_AUTH_MODE`: `device_code` or `auth_code_loopback`
- `OUTLOOK_READ_ONLY_MODE`: defaults to `true`
- `OUTLOOK_ENABLE_UNSAFE_TOOLS`: set to `true` only if you intentionally want the broader tool surface back
- `OUTLOOK_SCOPES`: optional explicit scope override
- `USE_TEST_MODE`: fake tokens and mock data for local testing

Legacy `MS_*` env vars are still accepted for compatibility.

## OpenClaw Example

```json
{
  "mcpServers": {
    "outlook-email-summarize": {
      "command": "node",
      "args": ["D:/5.github/outlook-email-summarize/index.js"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your-client-id",
        "OUTLOOK_TENANT_ID": "common",
        "OUTLOOK_AUTH_MODE": "device_code",
        "OUTLOOK_READ_ONLY_MODE": "true"
      }
    }
  }
}
```

## Optional Loopback Auth

This fork still ships a local callback server for debugging on a machine with a browser.

```bash
npm run auth-server
```

Loopback mode now:

- binds only to `127.0.0.1`
- validates OAuth `state`
- uses the same minimal default scopes

This is not the recommended path for cloud deployment. For Alibaba Cloud, just use device code auth.

## Unsafe Mode

If you want the original broad tool surface back, set:

```bash
OUTLOOK_ENABLE_UNSAFE_TOOLS=true
OUTLOOK_READ_ONLY_MODE=false
```

Then also provide the broader scopes yourself with `OUTLOOK_SCOPES`. If you do that, the extra risk is on you.

## Testing

Run the focused hardening tests:

```bash
npm test -- --runInBand test/config.test.js test/utils/feature-gates.test.js test/auth/device-code.test.js test/auth/oauth-server.test.js
```

Run the full test suite:

```bash
npm test -- --runInBand
```

## Notes

- Tokens are stored in `~/.outlook-mcp-tokens.json`
- sensitive token debug logging was removed
- the old environment dump helper was deleted on purpose

## License

MIT. See [LICENSE](./LICENSE).
