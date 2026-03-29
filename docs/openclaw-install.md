# OpenClaw-China Native Plugin Install

## 1. Install the repo as a native plugin

On the machine running OpenClaw-China:

```bash
openclaw plugins install -l /absolute/path/to/outlook-email-summarize
openclaw plugins list
openclaw plugins info outlook-email-summarize
```

This repo now exposes a real native plugin entry through `openclaw.plugin.json` and `package.json.openclaw.extensions`.

Do not inject it through root-level `mcpServers`. That was the wrong install shape and it can break gateway startup.

## 2. Configure the plugin entry

Add plugin config under `plugins.entries.outlook-email-summarize.config` in `~/.openclaw/openclaw.json`:

```json
{
  "plugins": {
    "entries": {
      "outlook-email-summarize": {
        "enabled": true,
        "config": {
          "clientId": "your-microsoft-client-id",
          "tenantId": "common",
          "authMode": "device_code",
          "readOnlyMode": true
        }
      }
    }
  }
}
```

`clientId` is required. `tenantId` should stay `common` for personal Outlook accounts.

## 3. Restart OpenClaw

```bash
openclaw gateway restart
openclaw gateway status --deep
```

## 4. First authentication

Send this in your chat with ClawBot:

```text
authenticate
```

It should return:

- a Microsoft device login URL
- a short user code

Open the URL, enter the code, approve access, then send `authenticate` again once.

## 5. Smoke test

Send:

```text
总结我 Outlook 收件箱最新 10 封邮件，按 发件人 / 主题 / 时间 / 一句话摘要 / 是否需要回复 输出
```

If auth is working, the agent should use `list-emails` and `read-email`.

## 6. If it fails

- `openclaw gateway status --deep`
- `journalctl --user -u openclaw-gateway.service -n 200 --no-pager`
- `openclaw plugins info outlook-email-summarize`
- confirm your Microsoft app has delegated permissions:
  - `offline_access`
  - `User.Read`
  - `Mail.Read`
