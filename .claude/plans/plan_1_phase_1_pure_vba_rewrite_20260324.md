# Plan 1 — Phase 1: Pure VBA Auth Scanner Rewrite

**Date:** 2026-03-24
**Status:** Proposed
**Objective:** Replace PS1 dependency with pure VBA — single .bas module that scans auth state from within Excel.

## Pre-conditions

- Enterprise blocks PS1 execution
- Solution must be self-contained in XLSM (no external scripts)
- VBA can access: registry, file system, environment, COM objects, workbook internals

## What VBA Can Detect (no PowerShell)

| Category | Method | Data |
|----------|--------|------|
| Windows Identity | `Environ()`, `WScript.Network` | Username, domain, computer |
| Office Identity | Registry `HKCU\...\Office\16.0\Common\Identity` | Signed-in O365 accounts, tenant, UPN |
| Token Stores | `FileSystemObject` + `Dir()` | OneAuth, TokenBroker file count/age/size |
| Token Expiry | Read JSON files via `ADODB.Stream` + regex extraction | `expires_on`, `exp` fields from token files |
| Workbook Connections | `ThisWorkbook.Connections` | Connection strings, auth types, server targets |
| Credential Manager | `cmd.exe /c cmdkey /list > temp` then read file | Office/SharePoint/SQL cached creds |
| Network Reachability | `MSXML2.XMLHTTP` | Can we reach login.microsoftonline.com, SharePoint |

## What VBA Cannot Do

- Parse JSON natively (need regex-based field extraction)
- Access 64-bit `ScriptControl` (no `eval()`)
- Directly query WAM/token broker APIs
- Delete Credential Manager entries without shelling to `cmdkey /delete` (still cmd.exe, not PS1)

## Architecture

Single module `ExcelCredManager.bas` with sections:

1. **Constants & Config** — sheet name, token store paths, stale thresholds
2. **Identity Collectors** — Windows user, Office accounts from registry
3. **Token Store Scanner** — FileSystemObject walks OneAuth/TokenBroker dirs, reads JSON for expiry
4. **Connection Inspector** — Enumerates workbook data connections, extracts auth method from connection strings
5. **Credential Manager Reader** — Shells `cmd /c cmdkey /list > temp`, parses output
6. **Report Builder** — Writes all findings to a formatted "AuthReport" sheet with sections and color coding
7. **Button Macros** — `ScanAuth()`, `RefreshConnections()`

### Key design decisions

- **No CSV intermediary** — write directly to worksheet, no temp file round-trip
- **Sectioned report** not flat table — group by category (Identity, Tokens, Connections, Credentials) since data shapes differ
- **JSON extraction via regex** — `VBScript.RegExp` to pull specific fields, no full parser needed
- **cmd.exe for cmdkey only** — the only shell-out, and it's cmd not PS1
- **Remove Clear/Delete functionality** — scanning only; deletion of tokens from VBA would still need shell commands and is risky without PS1's error handling

## Steps

- [ ] 1. Write new `ExcelCredManager.bas` with all sections
- [ ] 2. Remove PS1 dependency (delete `ExcelCredManager.ps1` or move to archive)
- [ ] 3. Test considerations doc (manual test matrix since VBA has no unit test framework)

## Success Criteria

- No PowerShell invocation anywhere
- Single XLSM file contains all logic
- Report sheet shows: identity, token status, connections, cached credentials
- Works in enterprise-locked environment

## Risks

- Registry paths vary across Office versions (15.0, 16.0) — need to check both
- Token JSON structure varies by Office build — regex extraction may miss fields
- `cmdkey /list` localization issue persists (carried over from PS1) — document as known limitation
- 64-bit vs 32-bit Office affects some COM objects

## Exit Criteria

- User confirms scan runs in their enterprise environment
- Report shows meaningful auth data
