---
name: outlook-email-summarize
description: 连接个人 Outlook 邮箱，读取过去24小时 Focus 类别未读邮件，输出可追溯摘要并提取3个最有洞察力的事实/观点。
version: 1.0.0
author: XiaoshengChen
metadata:
  hermes:
    tags: [Email, Outlook, Microsoft Graph API, Summarization, Insights]
    homepage: https://github.com/XiaoshengChen/outlook-email-summarize
prerequisites:
  commands: [python3]
---

# Outlook Email Summarize

连接个人 Outlook 邮箱，读取 Focus 类别未读邮件，输出可追溯的摘要和洞察。

## 触发条件

当用户说"看邮件"、"邮件摘要"、"outlook邮件"、"今日邮件"等时触发。

## 认证（首次使用）

脚本使用 Microsoft Graph API + device_code OAuth 认证，与 `~/.outlook-mcp-tokens.json` 共享 token 文件。

**首次认证步骤：**

```bash
# 步骤1：启动认证
python3 scripts/outlook_mail.py --auth

# 步骤2：用户在浏览器中完成授权后，轮询确认
python3 scripts/outlook_mail.py --poll
```

**检查认证状态：**

```bash
python3 scripts/outlook_mail.py --check
```

Token 自动 refresh，过期时会自动使用 refresh_token 获取新 token。

## 工作流

### 步骤1：拉取邮件列表

```bash
python3 scripts/outlook_mail.py --fetch --json
```

输出 JSON 包含：
- `emails[]`: 每封邮件的 id, subject, from_name, from_address, received_time, body_preview, importance, has_attachments, inference_classification
- `total_unread`: 所有未读数量
- `focused_count`: Focus 类别未读数量

### 步骤2：筛选与读取

根据邮件列表，**只对可能有洞察的邮件读取正文**（最多5封）：

```bash
python3 scripts/outlook_mail.py --read "<email_id>"
```

**筛选优先级：**
1. 非自动通知类邮件（人工撰写的内容优先）
2. importance 为 "high" 的邮件
3. 来自人（而非 newsletter机器人）的邮件
4. body_preview 有实质内容的邮件

**跳过规则：**
- 纯通知/提醒类邮件（如"你的订单已发货"）
- 重复或无实质信息的邮件
- 短邮件如果 body_preview 已足够，不必再读取正文

### 步骤3：AI 撰写摘要与洞察

基于拉取的邮件内容，AI 撰写输出，严格遵循以下格式：

---

**总摘要**（300-500字）

概述过去24小时邮件的整体情况，包括关键主题、趋势、需要关注的事项。

---

**3个最有洞察力的事实/观点**

格式严格如下：

```
发件人1
洞察1，只用1-2句话
必要时附关键原文

发件人2
洞察2，只用1-2句话
必要时附关键原文

发件人3
洞察3，只用1-2句话
必要时附关键原文
```

---

## 洞察提取规则

"洞察"应优先体现：
- **新的有趣事实** — 不是常识的信息
- **反共识的观点** — 与主流认知不同的看法
- **犀利的洞察角度** — 从独特视角解读现象
- **有利于投资的信号** — 市场趋势、行业变化、政策暗示

**不要只是复述邮件内容**。要从内容中提炼出超越表层的信息价值。

## 诚信规则

- 不确定、不了解的信息**必须明确说明**
- **不要编造**任何不在邮件中的信息
- 如果邮件数量不足以提取3个洞察，如实说明"本期仅能提取N个洞察"
- 原文引用时确保准确，标注出处（发件人+主题）

## 参数

| 参数 | 默认值 | 说明 |
|------|--------|------|
| --hours | 24 | 回溯时间（小时） |
| --fetch | - | 拉取邮件列表 |
| --read | - | 读取指定邮件正文（需 email_id） |
| --auth | - | 启动 device_code 认证 |
| --poll | - | 确认 device_code 授权 |
| --check | - | 检查认证状态 |

## 自动化：每日邮件摘要 Cron

已配置 cron job，每天早上 8:30 自动运行并推送摘要到微信：

- **Cron Job ID**: `969a0f97ec0c`
- **Schedule**: `30 8 * * *`（每天 8:30 Asia/Shanghai）
- **Deliver**: 微信（weixin）
- **无需手动 approve**：`python3` 已加入 `command_allowlist`（在 `~/.hermes/config.yaml`）
- **Token 过期自动刷新**：脚本内置 refresh_token 机制

管理命令：
```bash
hermes cron list              # 查看所有 cron jobs
hermes cron poll 969a0f97ec0c  # 查看上次运行状态
hermes cron pause 969a0f97ec0c # 暂停
hermes cron resume 969a0f97ec0c # 恢复
hermes cron remove 969a0f97ec0c # 删除
```

## 故障处理

| 问题 | 原因 | 解决 |
|------|------|------|
| `auth_required` | token 过期或不存在 | `--auth` 重新认证 |
| 无 Focus 邮件 | Outlook 未启用 Focus/Other 分离 | 脚本自动回退到所有未读邮件 |
| API 401 | token 失效 | `--auth` 重新认证 |
| Cron 未推送 | token 过期或 API 错误 | 手动 `--check` 检查状态，必要时 `--auth` 重新认证 |