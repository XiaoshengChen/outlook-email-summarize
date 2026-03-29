# Email Summary Security Hardening Design

## Goal

Preserve Outlook inbox reading and summarization while making the default deployment safe for a headless OpenClaw server.

## Problem

The upstream project is optimized for a broad Microsoft 365 toolbox, not a narrowly scoped email summarizer. That creates four bad defaults:

- OAuth scopes are broader than the product needs.
- Sensitive auth details can leak into logs.
- Loopback auth is not hardened for cloud and headless deployments.
- Write-capable tools stay available even when the operator only wants inbox summary.

## Chosen Direction

Implement a read-first, server-safe mode and make it the default.

- Default auth mode becomes `device_code`, which fits headless cloud servers.
- Default scope set becomes `offline_access User.Read Mail.Read`.
- Mutating tools are disabled by default through a central feature gate.
- Loopback auth remains available for local debugging, but binds to `127.0.0.1` and validates OAuth `state`.
- Sensitive debug output is removed, and token persistence is tightened.

This keeps the one thing that matters, email summary, while shrinking the blast radius.

## Scope

### In

- Authentication flow hardening.
- Default read-only configuration.
- Tool gating for high-risk capabilities.
- Token storage and logging cleanup.
- Docs and license cleanup.
- Tests covering the new defaults and guardrails.

### Out

- Rebuilding the whole MCP surface.
- Adding enterprise-only workflows.
- Expanding functionality beyond read/search/summarize email.

## Architecture

### Auth Model

Two auth modes will exist:

- `device_code`: default, for Alibaba Cloud and other headless deployments.
- `auth_code_loopback`: optional, for local development and debugging.

Both modes share centralized config for scopes and token persistence. Device code flow becomes the primary path documented for OpenClaw.

### Read-Only Mode

Introduce a central `READ_ONLY_MODE` flag that defaults to `true`.

When enabled:

- Email read/list/search remains enabled.
- Mutating Outlook actions are blocked.
- Calendar, OneDrive, Power Automate, rules, and folder-mutation features are blocked unless explicitly re-enabled.

This is not cosmetic. The server should fail closed.

### Tool Gating

Add a small gating utility so mutating modules do not each invent their own ad hoc checks.

The gating layer should:

- Expose named capabilities such as `mail_write`, `calendar_write`, `storage_write`, `automation_write`, and `rules_write`.
- Throw a consistent operator-facing error when a disabled capability is called.
- Make the default mode obvious in logs without leaking secrets.

### Token Storage and Logging

Keep file-based token storage for now because it is simple and good enough for a single-user OpenClaw server. Harden it by:

- Removing token content logging.
- Removing environment dump utilities.
- Writing token files with restricted permissions when possible.
- Making auth-related logs descriptive but non-sensitive.

### Loopback OAuth Hardening

The loopback callback server should:

- Bind only to `127.0.0.1`.
- Generate cryptographically strong `state`.
- Validate returned `state` before exchanging the code.

That closes the dumbest hole in the current implementation.

## Approaches Considered

### A. Reuse modular auth pieces and add device code flow

Use the cleaner `auth/` modules as the center of gravity, then adapt the standalone auth entrypoint around them.

Why this wins:

- Lower duplication.
- Easier to test.
- Lets us keep local loopback auth without making it the default.

### B. Patch only the root standalone auth server

Rejected. It keeps too much auth logic concentrated in the least disciplined file and makes the codebase harder to reason about later.

### C. Delete everything except email read tools

Rejected for now. It is cleaner in theory, but it is a larger product decision than the user asked for and risks breaking future reuse.

## Testing Strategy

Add or update tests for:

- Default read-only scope selection.
- OAuth state validation in loopback flow.
- Device code auth initiation and polling behavior.
- Disabled write tools in read-only mode.
- Sensitive logs not containing token material.

Verification should include at least targeted Jest coverage for auth, config, and the gated tool paths.

## Risks and Mitigations

### Risk: hidden coupling between write scopes and existing read paths

Mitigation: keep read-email tests and inbox listing tests green while introducing gating.

### Risk: breaking local developer auth

Mitigation: preserve loopback auth as an explicit non-default mode and document how to enable it.

### Risk: partial gating leaves side doors open

Mitigation: centralize capability checks and cover the highest-risk mutating tools with tests.

## Expected Outcome

After this change, the repo should behave like an Outlook email summarizer first and a dangerous Microsoft swiss-army knife only when an operator consciously opts into that risk.
