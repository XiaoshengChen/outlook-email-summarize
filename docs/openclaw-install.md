# OpenClaw Install And Smoke Test

## 1. Install the bundle

On the machine running OpenClaw:

```bash
openclaw plugins install /absolute/path/to/outlook-email-summarize
openclaw plugins list
openclaw plugins inspect outlook-email-summarize
```

You want to see the plugin detected as a bundle with MCP capability.

## 2. Add Microsoft credentials

Put this in `~/.openclaw/.env` on the OpenClaw host:

```bash
OUTLOOK_CLIENT_ID=your-client-id
OUTLOOK_TENANT_ID=common
OUTLOOK_AUTH_MODE=device_code
OUTLOOK_READ_ONLY_MODE=true
OUTLOOK_ENABLE_UNSAFE_TOOLS=false
```

Only add `OUTLOOK_CLIENT_SECRET` if you intentionally switch to loopback auth.

## 3. Restart OpenClaw

```bash
openclaw gateway restart
```

## 4. First authentication

Send this in your chat with ClawBot:

```text
authenticate
```

OpenClaw should reply with:

- a Microsoft device login URL
- a short user code

Open the URL, enter the code, approve access, then send:

```text
authenticate
```

again once.

## 5. Smoke test

Send:

```text
总结我 Outlook 收件箱最新 10 封邮件，按 发件人 / 主题 / 时间 / 一句话摘要 / 是否需要回复 输出
```

If auth is working, the agent should use `list-emails` and optionally `read-email` for the important messages.

## 6. If it fails

- `openclaw logs`
- `openclaw plugins inspect outlook-email-summarize`
- check `~/.openclaw/.env`
- confirm your Microsoft app has delegated permissions:
  - `offline_access`
  - `User.Read`
  - `Mail.Read`
