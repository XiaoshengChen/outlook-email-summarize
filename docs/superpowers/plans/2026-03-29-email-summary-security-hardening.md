# Email Summary Security Hardening Implementation Plan

## Goal

Ship a safe default for `OpenClaw + WeChat ClawBot + Outlook inbox summary` without losing the ability to read and summarize recent emails.

## Architecture

- Node.js MCP server with centralized config and capability gates.
- Default auth path: Microsoft device code flow.
- Optional local debug auth path: loopback auth code flow on `127.0.0.1`.
- Default permissions: `offline_access User.Read Mail.Read`.

## Tech Stack

- Node.js
- Jest
- Microsoft Graph OAuth endpoints
- Existing MCP tool architecture in this repo

## Task 1: Lock expected behavior with tests

Add failing tests that define the new contract:

- read-only mode is enabled by default
- default scopes are minimal and read-only
- loopback callback rejects invalid `state`
- device code flow can be started and polled
- mutating tools fail closed in read-only mode

## Task 2: Centralize read-only config and capability gating

Create a small feature-gate utility and wire config through it.

Files:

- `config.js`
- `utils/feature-gates.js`
- modules that expose write actions

## Task 3: Harden token storage and kill dangerous debug paths

Remove sensitive logging and environment dumping.

Files:

- `auth/token-manager.js`
- `auth/token-storage.js`
- delete `debug-env.js`

## Task 4: Add device code auth flow

Implement device code start and poll helpers, then expose them through the auth tooling.

Files:

- `auth/device-code.js`
- `auth/tools.js`
- auth entrypoints as needed

## Task 5: Harden loopback auth

Make loopback auth explicit and safer.

Files:

- `auth/oauth-server.js`
- `outlook-auth-server.js`

## Task 6: Disable risky capabilities by default

Keep inbox reading alive, but block writes unless the operator opts in.

Candidate modules:

- email send and draft
- email state mutation
- folder mutation
- rules
- calendar write operations
- OneDrive write operations
- Power Automate write operations

## Task 7: Update docs and licensing

Explain the new default deployment path and make the license explicit.

Files:

- `README.md`
- `LICENSE`

## Task 8: Verify

Run targeted Jest suites first, then the full relevant test command if the targeted passes are clean.

Success means:

- inbox list/read tests still pass
- new auth and gating tests pass
- no sensitive debug output remains in the touched paths
