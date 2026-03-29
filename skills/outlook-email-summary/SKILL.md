---
name: outlook-email-summary
description: Use the Outlook inbox MCP tools in a token-efficient way to summarize recent email and identify what needs action.
---

# Outlook Email Summary

Use this when the user asks to summarize Outlook mail, triage recent messages, or identify which emails need a reply.

## Workflow

1. Call `list-emails` first.
2. Use `search-emails` only if the user asks for a filtered slice.
3. Call `read-email` only for the messages that look important, ambiguous, or action-heavy.
4. Summarize in this format:

- sender
- subject
- received time
- one-line summary
- action needed: `reply` | `read later` | `ignore`

## Rules

- Prefer reading at most 3-5 full emails unless the user explicitly asks for more detail.
- If the inbox is dominated by newsletters or automated mail, say so plainly.
- Surface deadlines, asks, approvals, and anything that looks like a hidden landmine.
- Do not invent unread counts or action items that are not in the messages.
