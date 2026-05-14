# Email AI Agent — Master Project Reference

> This document is the single source of truth for all Claude sessions on this project.
> Read this at the start of every session. Update the "Current State" section at the end of every session.
> Never assume anything not written here. Ask if uncertain.

---

## 1. Who this is for

A sole operator running three businesses/personas:

- Financial planning business (primary client-facing)
- Tax business (separate client base)
- Personal life

Goal: AI agent running on a local Mac mini that autonomously manages email triage, spam deletion, filing, draft generation, and calendar — with human approval for anything outgoing. Accessible via browser at home and via Telegram when travelling.

---

## 2. Hardware and environment

| Item | Detail |
|---|---|
| Machine | MacBook Pro 16 M4 Max (development). Mac mini Apple Silicon on order — will migrate when it arrives |
| macOS | 26.4.1 |
| RAM allocated to this project | 32 GB maximum |
| Package managers installed | Homebrew 5.1.9, npm 11.12.1, pip 26.1 |
| Python | 3.14.4 installed via Homebrew — use `python3` and `pip3` |
| Node | 25.9.0 installed via Homebrew |
| Ollama | 0.23.0 — runs as Mac App Store menu bar app, NOT a brew service |
| Shell | zsh |
| Browser | Microsoft Edge (primary). Safari used as fallback for Microsoft portal sessions only |
| Network | Tailscale installed on all devices. No public internet exposure. All services bind to Tailscale IP only, never 0.0.0.0 |
| Tailscale | Two installs: Mac App Store (menu bar) and Homebrew CLI (`/usr/local/bin/tailscale`). System extension (`io.tailscale.ipn.macsys.network-extension`) is the actual daemon — both GUI and CLI talk to it. CLI has full control: `tailscale up`, `tailscale down`, `tailscale status`, `tailscale ip -4`. Menu bar app is optional UI only. Never use `ifconfig` for IP — use `tailscale ip -4`. |
| Remote access | Tailscale mesh — MacBook Pro, iPhone, and other devices all on same Tailscale network |

---

## 3. Email and calendar accounts

| Account | Email | Platform | API | Status |
|---|---|---|---|---|
| Financial planning | `<FINANCIAL_EMAIL>` | Microsoft 365 (corporate) | Microsoft Graph API | ✅ Connected |
| Tax business | `<GMAIL_EMAIL>` | Gmail | Gmail API | ✅ Connected |
| Personal | `<PERSONAL_EMAIL>` | Exchange Online (Plan 1) | Microsoft Graph API | ✅ Connected — application permissions |

---

## 4. M365 account and Azure app details

### Personal account — ✅ Connected via application permissions

`<PERSONAL_EMAIL>` is a real member user in the Intertek Entra ID tenant (Object ID: `<AZURE_OBJECT_ID_PERSONAL>`). Exchange Online Plan 1. Highest volume inbox of the three accounts.

**Auth method:** Client credentials flow (`grant_type=client_credentials`) using `AZURE_CLIENT_ID_PERSONAL` + `AZURE_CLIENT_SECRET_PERSONAL`. No user sign-in or OAuth flow. Token cached as `personal_app` in `tokens_graph.json`. All Graph calls use `/users/<PERSONAL_EMAIL>/...`. Permissions: `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite` — all Application type with admin consent.

**If 403 after permission changes:** Delete `personal_app` key from `tokens_graph.json` to force fresh token acquisition.

### Tenant

| Item | Value |
|---|---|
| Tenant ID | `<AZURE_TENANT_ID>` |
| Tenant name | <TENANT_NAME> |

### Users in tenant

| Account | Email | Object ID |
|---|---|---|
| Financial planning | `<FINANCIAL_EMAIL>` | `<AZURE_OBJECT_ID_FINANCIAL>` |
| Personal | `<PERSONAL_EMAIL>` | `<AZURE_OBJECT_ID_PERSONAL>` |

### App registrations

| App | Purpose | Client ID key in .env | Secret key in .env |
|---|---|---|---|
| `email-agent` | M365 financial planning | `AZURE_CLIENT_ID_FINANCIAL` | `AZURE_CLIENT_SECRET_FINANCIAL` |
| `email-agent-personal` | M365 personal (parked) | `AZURE_CLIENT_ID_PERSONAL` | `AZURE_CLIENT_SECRET_PERSONAL` |

### API permissions — email-agent app (confirmed granted)

| Permission | Type | Purpose |
|---|---|---|
| `Mail.ReadWrite` | Delegated | Read, file, delete emails |
| `Mail.Send` | Delegated | Send approved drafts |
| `Calendars.ReadWrite` | Delegated | Read and create calendar events |
| `User.Read` | Delegated | Identify authenticated account |
| `offline_access` | Delegated | Maintain connection via refresh token |

### Redirect URIs — email-agent app (confirmed registered)
- `http://localhost:8000/auth/callback/financial`
- `http://localhost:8000/auth/callback/personal`

### Google Cloud — Gmail app registration

| Item | Key name in .env |
|---|---|
| Client ID | `GOOGLE_CLIENT_ID` |
| Client secret | `GOOGLE_CLIENT_SECRET` |

**Gmail OAuth app name:** `email-agent`
**Redirect URI:** `http://localhost:8000/auth/callback/gmail`
**Scopes:** `gmail.modify`, `calendar`, `userinfo.email`, `userinfo.profile`, `openid`

### Security notes
- Azure client secret (`AZURE_CLIENT_SECRET_FINANCIAL`) was visible in chat — rotate after personal account issue resolved
- Gmail client secret (`GOOGLE_CLIENT_SECRET`) was visible in chat — rotate after personal account issue resolved
- Steps to rotate Azure secret: App registrations → email-agent app → Certificates & secrets → delete old → New client secret → copy value → update `.env`
- Steps to rotate Gmail secret: Google Cloud Console → Credentials → edit email-agent client → Reset secret → update `.env`

---

## 5. Chosen tech stack

### Backend
- **FastAPI** 0.136.1 (Python, async) — agent core and API server
- **uvicorn** 0.46.0 — ASGI server to run FastAPI
- **httpx** 0.28.1 — async HTTP client for API calls
- **APScheduler** 3.11.2 — polling scheduler (every 5 minutes per inbox)
- **SQLite** — action log, draft history, learning loop data, per-account settings (prompts, footers)
- **aiosqlite** 0.22.1 — async SQLite driver
- **Jinja2** 3.1.6 — HTML templating
- **python-dotenv** 1.2.2 — loads .env secrets
- **ChromaDB** — semantic vector store (Phase 6, deferred)

### Python 3.14 compatibility note
Always use:
```python
templates.TemplateResponse(request, "template.html")
```
Never:
```python
templates.TemplateResponse("template.html", {"request": request})
```

### AI tiers

| Tier | Tool | Used for |
|---|---|---|
| Local | Ollama 0.23.0 + phi4-mini:latest | Classification only |
| Cloud | Claude API — claude-sonnet-4-6 | On-demand draft generation (primary) |
| Cloud | OpenAI API — gpt-4o | Draft generation fallback if Claude fails |
| Cloud | Gemini API | Not yet configured — skipped for now |

**Note on local model:** Must use `phi4-mini:latest` — not `phi4-mini`. Ollama runs as Mac App Store menu bar app — not a brew service, do not use `brew services` to manage it.
**Note on Claude API:** Model string is `claude-sonnet-4-6`. Separate Anthropic API account from Claude Pro. ANTHROPIC_API_KEY set in `.env`.
**Note on OpenAI API:** Model string is `gpt-4o`. Used as fallback only — same prompt as Claude. OPENAI_API_KEY set in `.env`.

**Considered and deferred — Anthropic Python SDK (`anthropic` package):**
Currently `drafter.py` and `learner.py` call the Anthropic API directly via `httpx`. Switching to the official SDK (`pip install anthropic`) was evaluated and deferred. Summary:

| | Detail |
|---|---|
| **Pro: Prompt caching** | SDK supports `cache_control: ephemeral` on system prompts. The per-account writing instruction (~500 tokens) could be cached across multiple draft calls, reducing input token cost. Benefit only realised if multiple drafts are generated within the 5-minute cache TTL — uncommon for a sole operator. |
| **Pro: Typed errors** | SDK raises `anthropic.APIStatusError`, `anthropic.APIConnectionError` etc. instead of `httpx.HTTPStatusError`. Marginally better log messages. |
| **Pro: Built-in retries** | SDK retries on transient errors by default. The current flow is not latency-sensitive enough for this to matter. |
| **Con: New dependency** | Adds `anthropic` + 5 transitive packages (`distro`, `docstring-parser`, `jiter`, `sniffio`, `anyio`). The current `httpx` approach has zero extra deps (httpx is already required for Graph/Gmail). |
| **Con: Working code replaced** | The raw httpx calls have no known bugs. Changing them introduced risk with no user-facing benefit. |
| **Verdict** | Not worth it at current usage volume. Revisit if: (a) multiple users, (b) batch draft generation is added, or (c) cost becomes a concern. If adopted, split `generate_draft()` into `system_prompt = base_prompt + voice_block` (cached) and `user_content = email details` (per-call). |

### Anthony's voice — draft prompt calibration
Per-account base prompts and footers are stored in the SQLite `settings` table and editable via ⚙ Settings → Account Configuration in the dashboard. Hardcoded defaults in `agent/drafter.py` are used only on first run before the user saves via the UI.

- Wise, calm, direct. Genuinely on the client's side.
- Short, precise sentences. One idea per sentence. No run-ons.
- Prose by default. Structure only when items are genuinely parallel or sequential.
- Leads with human moment when one exists.
- Cites ATO references and legislation when useful — not to impress.
- Frames as general considerations, not specific recommendations (compliance).
- Sign-off always "Kind Regards".
- Signature block appended by `generate_draft()` after the LLM response — never generated by the model. Footer loaded from `settings` table (`footer_financial` / `footer_gmail`).
- Financial account (Intertek): investment, super, insurance, cashflow, retirement focus. Authorised Representative under AFSL.
- Gmail account (Positive Tax): tax, BAS, tax planning, SMSF, business structuring focus. Registered Tax Agent.

### Frontend
- Served on Tailscale IP only, no public port
- Two-mode interface: Inbox (pending only) and History (sent/archived/deleted)
- Three-panel layout: sidebar / queue / detail — all panels resizable
- Vertical resizer between email body and draft panel

---

## 6. Agent behaviour — autonomous action rules

Current comfort level: **Moderate**

| Action | Autonomous? | Threshold | Notes |
|---|---|---|---|
| Hard delete spam | Yes — auto | Confidence > 0.95 | Permanently deleted — unrecoverable |
| Move to Junk | Yes — auto | Confidence ≤ 0.95 | Lower confidence spam |
| Move newsletters | Yes — auto | Any confidence | Moved to Newsletters folder |
| Move notifications | Yes — auto | Any confidence | Moved to Notifications folder |
| Generate draft reply | On-demand only | Any classification | User clicks Generate Draft button |
| Send any email | Never autonomous | — | Always requires explicit human action |

**Delete behaviour:**
- UI delete (🗑 button) → soft delete → moves to Deleted Items (M365) or Trash (Gmail). Recoverable for 30 days.
- Autonomous spam delete (confidence > 0.95) → hard delete → permanently removed from server.

**Poll behaviour:** Fetches ALL messages currently in inbox on every poll cycle (fully paginated, no count limit). Reconciles SQLite pending records against live inbox — anything pending in SQLite but no longer in inbox is automatically marked archived. Duplicate email_ids prevented by unique index.

**Inbox model:** If it's in the inbox it's pending. If it's not in the inbox it's not pending. Always reflects live mailbox state after each poll.

**Folders auto-created:** Newsletters, Notifications, Junk Email, Archive — on first use in both M365 and Gmail.

**Inbox view actions per classification:**

| Classification | Actions available |
|---|---|
| `action_required` | Send Reply, Archive, 📁 File…, Delete |
| `fyi` | Reply (draft panel), Archive, 📁 File…, Delete |
| `calendar` | Accept, Decline, Archive, 📁 File… |
| `notification` | Archive, 📁 File…, Delete |
| `newsletter` | Archive, 📁 File…, Delete |
| `spam` | Archive instead, 📁 File…, Delete |

File button always sits between Archive and Delete (or after Archive when there is no Delete). All classifications also have Flag and Reclassify controls.

**History view actions per status:**

| Status | Actions available |
|---|---|
| `sent` | Send Follow-up (editable To field, draft panel), Clear Record |
| `archived` | Move to Inbox (live unarchive + mark pending), Delete permanently, Clear Record |
| `deleted` | Restore to Queue (stub), Clear Record |

**Clear Record** — deletes single SQLite row by email_id via `DELETE /api/email/{id}/record`. Does not affect mailbox. Does not bulk-clear.

**Clear History button** (History mode sidebar) — bulk clear by scope: Sent / Archived / Deleted / All. Does not affect mailbox. Requires confirmation.

---

## 7. Phased implementation plan

| Phase | Description | Status |
|---|---|---|
| 1 | Foundation — Python venv, core libraries, FastAPI test page | ✅ Complete |
| 2 | Email connectivity — OAuth for all 3 accounts, read-only test | ✅ Complete (all 3 connected) |
| 3 | Agent core — polling loop, classification, SQLite logging | ✅ Complete |
| 4 | Web UI — FastAPI dashboard, action queue | ✅ Complete |
| 5A | Autonomous actions — delete spam, move newsletters/notifications | ✅ Complete |
| 5B | Draft generation — Claude API on-demand drafts | ✅ Complete |
| 5C | Live send, archive, mark-read, full workflow | ✅ Complete |
| 5D | History view, unarchive, restore, clear history, soft delete | ✅ Complete |
| 5E | Full inbox sync — paginated fetch, reconciliation, account filter | ✅ Complete |
| 5F | Clear Record fix, no confirmation dialogs, sender address parsing | ✅ Complete |
| 6 | Learning loop — synthesise voice profile from sent emails, inject into drafts | ✅ Complete |
| 7 | Telegram integration — mobile channel for travel | Not started |

---

## 8. Project folder structure (current)

```
~/email-agent/
├── email-agent              # Master control script (also at /opt/homebrew/bin/email-agent)
├── main.py                  # FastAPI app — all routes
├── agent/
│   ├── __init__.py
│   ├── actions.py           # Autonomous actions — hard_delete_email for spam, move for newsletters/notifications
│   ├── classifier.py        # Ollama/phi4-mini:latest classification
│   ├── drafter.py           # Claude API on-demand draft generation
│   └── poller.py            # Full paginated inbox fetch, reconciliation, no count limit
├── connectors/
│   ├── graph.py             # M365 — paginated full inbox, soft delete, hard_delete_email, archive, unarchive, send, mark read
│   └── gmail.py             # Gmail — paginated full inbox, soft delete, hard_delete_email, archive, unarchive, send, mark read
├── db/
│   ├── __init__.py
│   └── database.py          # SQLite — action_log, sent_examples, voice_profiles, settings, filing_history, sender_rules tables
├── ui/
│   ├── templates/
│   │   └── dashboard.html   # Two-mode UI, account filter, history view, no confirmation dialogs
│   └── static/
├── logs/                    # PID files and logs — 14 day retention, auto-purged on start
├── config/
│   ├── .env                 # All secrets and config
│   ├── agent.db             # SQLite database (auto-generated)
│   ├── tokens_graph.json    # M365 OAuth tokens (auto-generated)
│   └── tokens_gmail.json    # Gmail OAuth tokens (auto-generated)
└── requirements.txt
```

---

## 9. API credentials and secrets (locations, not values)

All secrets live in `~/email-agent/config/.env` — never hardcoded, never in git.

| Secret | Key name in .env | Status |
|---|---|---|
| Azure app client ID (M365 financial) | `AZURE_CLIENT_ID_FINANCIAL` | Set |
| Azure app client secret (M365 financial) | `AZURE_CLIENT_SECRET_FINANCIAL` | Set — rotate when personal resolved |
| Azure tenant ID | `AZURE_TENANT_ID` | Set |
| Azure object ID — financial planning user | `AZURE_OBJECT_ID_FINANCIAL` | Set |
| Azure app client ID (M365 personal) | `AZURE_CLIENT_ID_PERSONAL` | Set |
| Azure app client secret (M365 personal) | `AZURE_CLIENT_SECRET_PERSONAL` | Set |
| Azure object ID — personal user | `AZURE_OBJECT_ID_PERSONAL` | Set |
| Google OAuth client ID (Gmail tax) | `GOOGLE_CLIENT_ID` | Set |
| Google OAuth client secret (Gmail tax) | `GOOGLE_CLIENT_SECRET` | Set — rotate when personal resolved |
| Anthropic API key (Claude) | `ANTHROPIC_API_KEY` | ✅ Set |
| OpenAI API key | `OPENAI_API_KEY` | ✅ Set |
| Google Gemini API key | `GEMINI_API_KEY` | Not yet set |
| Default cloud provider for drafting | `DRAFT_DEFAULT_PROVIDER` | claude |
| Default cloud provider for analysis | `COMPLEX_ANALYSIS_DEFAULT_PROVIDER` | claude |
| Fallback cloud provider | `FALLBACK_PROVIDER` | gemini |
| Tailscale IP of active machine | `TAILSCALE_IP` | `<TAILSCALE_IP>` |

**Migration note:** When Mac mini arrives, update `TAILSCALE_IP`. Confirm Ollama menu bar app starts at login on Mac mini.

---

## 10. How to work with Claude on this project

### Session discipline
- Paste this document (or confirm Claude has read it from the Project) at the start of every session
- State the single goal for that session in one sentence before asking anything else
- End every session by asking Claude to produce the complete updated project reference document as a download

### File management rules
- Claude always provides complete files — never partial edits
- Project reference document always provided as a download
- Terminal commands shown as copyable code blocks
- To replace a file using vi: open with `vi`, type `ggdG` to clear, press `i`, paste, press `Esc` then `:wq`

### Prompting rules that prevent hallucinations
- Always paste the actual error message — never describe it in your own words
- Always paste the relevant code block — never summarise what it does
- Ask Claude to produce one file at a time, test it, then move to the next

### How to start each session
```
I am working on my email AI agent project.
Today's goal: [one sentence]
```

### Red flags — stop and clarify if Claude does any of these
- Suggests a library or tool not in the stack above without explaining why
- Produces code that references a file or function not yet created
- Gives instructions that differ from what worked in a previous session
- Asks you to modify Azure or Google app settings without explaining the exact change needed

---

## 11. Current state (update this after every session)

**Last updated:** 2026-05-10 (session 5). Major reliability and correctness pass across all three connectors and the entire action pipeline. All known bugs from the audit are fixed. System is feature-complete and stable.

### What changed in session 5

1. **Financial inbox showing 0 emails (silent TypeError)** — `get_valid_token()` compared `datetime.utcnow()` (naive, no tz) against an ISO string like `"2026-05-09T12:00:00+00:00"` (aware). Python raises `TypeError` which was silently caught by the outer `except Exception` in the poller, logging an auth error but no poll ever succeeding. Fixed by `_parse_expires_at()` helper and `datetime.now(timezone.utc)` throughout `graph.py` and `gmail.py`. Auth error keywords now include `"offset-naive"` and `"offset-aware"` to catch future recurrence.

2. **Archive / delete / calendar actions using wrong Graph ID** — action endpoints in `main.py` were passing `email_id` (the stable RFC 2822 `internetMessageId`) to Graph API, which requires the folder-scoped `id`. Fixed by storing `graph_id` at poll time in `action_log` and using `graph_id = email.get("graph_id") or email_id` at every action call site.

3. **Unarchive 404** — the stored `graph_id` goes stale after a message is moved to archive (Graph assigns a new ID per folder). Fixed by calling `get_message_graph_id(account, internet_message_id)` at unarchive time — searches Graph for the message by `internetMessageId` and returns the current folder-scoped ID regardless of where it's been moved.

4. **`refresh_token()` missing `raise_for_status()`** — a 4xx from the token endpoint was silently stored as the token dict. Fixed; all HTTP calls in both connectors now call `raise_for_status()`.

5. **Gmail sender rules mismatch** — raw `From` header was stored as the sender key (e.g. `"Acme <news@acme.com>"`). Sender rules keyed on bare address never matched. Fixed: `poll_gmail` now uses `parseaddr()` to extract the bare address before storing.

6. **Sender rules overhaul** — removed learned rules entirely. AI classifications are no longer written to `sender_rules`. Only manual reclassifications via `api_reclassify` create/increment rules. count=1 = pending (no effect), count≥2 = active (overrides AI). Reclassifying to a different category resets count to 1. Pollers check `rule['source'] == 'manual' and rule['count'] >= 2`. UI shows pending/active state clearly. Per-row delete fixed with `unquote()` (URL decoding) and `data.ok` check.

7. **Inbox as source of truth** — fundamental redesign: if an email is present in the live inbox, it must be pending in the app, regardless of prior DB status. `ensure_inbox_state(email_id, graph_id)` does a single UPDATE resetting status to pending and refreshing graph_id. Any email pending in DB but missing from inbox gets archived via `mark_missing_as_archived()`.

8. **Poll on page load / tab focus** — silent background poll triggered on page load (UI renders from cached DB first, refreshes after 8s) and on tab regain-focus (throttled to 60s minimum between polls via `_lastPollAt`). Reduces sync delay without infrastructure changes.

9. **Calendar implemented end-to-end** — Outlook: single Graph call (`/messages/{graph_id}/accept|decline`, `sendResponse: true`). Gmail: multi-step ICS extraction → Calendar API event lookup by UID → PATCH attendee responseStatus → `sendUpdates: all`. Both wired to UI buttons (✓ Accept / ✕ Decline) on calendar-classified emails. Requires `https://www.googleapis.com/auth/calendar` scope — included in Gmail OAuth flow.

**All phases 1–6: Complete. All three accounts: Complete.**

**email-agent script — important fixes (cumulative):**
- **`head()` → `hdr()`** — script defined `head()` as a bold-heading shell function, shadowing the system `head` command. In `cmd_stop`, every `pgrep ... | head -1` called the shell function instead of filtering output, so `running_pid` contained ANSI escape text instead of a real PID — stop never actually killed anything. Fixed by renaming to `hdr()` (session 1). Session 2: same rename applied to the remaining call sites.
- **`nohup` + `disown`** added to main server start. Ensures process survives script exit.
- **`tailscale up` attempt on start** — tries `tailscale up` before failing. Records `email-tailscale-started-by-us` flag if it brought Tailscale up.
- **Tailscale stop** — disconnects only if email-agent brought it up AND content-agent is not still running.
- **`content_agent_running()`** added — `pgrep -f "uvicorn server:app"`. Used for shared-resource coordination.
- **Ollama ownership flag written at start time** — flag `ollama_started_by_us` is now written immediately when `ollama serve` is launched, not at the end of `cmd_start`. Previously, if the server was already running (early-return path), the flag was never written even though Ollama had been started and `ollama.pid` had been written — leaving Ollama orphaned on next stop.
- **Ollama stop — ownership transfer model** — on stop, checks both `ollama_started_by_us` AND `~/content-agent/logs/content-ollama-started-by-us`. If either flag exists and content-agent IS running: write content-agent's flag (transfer ownership) so it will clean up. If content-agent is NOT running: kill Ollama (`pgrep -f "ollama serve"`) and remove all flags. If neither flag exists: Ollama was pre-existing, leave it alone.
- **Status shows Content Agent** — `email-agent status` reports whether content-agent is running.

**Shared resource coordination (email-agent ↔ content-agent):**
- email-agent detects content-agent via: `pgrep -f "uvicorn server:app"`
- content-agent detects email-agent via: `pgrep -f "uvicorn main:app"`
- **Ollama — ownership transfer model:** each agent checks both flag files on stop. Last agent out kills Ollama. If neither flag exists (Ollama was pre-existing), neither agent touches it. Full model documented in content-agent project reference.
- Tailscale: per-agent ownership. Neither disconnects if the other is still running. Pre-existing connections never touched.

**What is working:**
- `email-agent start/stop/restart/debug/status` — Tailscale CLI, Ollama, FastAPI all managed
- Tailscale fully controlled via CLI — menu bar app not required
- FastAPI on `<TAILSCALE_IP>:8000`
- All three accounts connected and polling — M365 financial (delegated OAuth), Gmail tax (delegated OAuth), M365 personal (application permissions / client credentials)
- Full paginated inbox fetch — every email in inbox fetched every poll cycle
- Reconciliation — pending emails no longer in inbox auto-archived each cycle
- Unique index on email_id — duplicates impossible
- phi4-mini:latest classifying accurately
- Autonomous actions: spam hard-deleted (>0.95), newsletters/notifications moved
- On-demand draft generation with voice profile injection
- Draft tone calibrated per account — base prompt read from `settings` table (`prompt_financial` / `prompt_gmail` / `prompt_personal`), falls back to hardcoded defaults in `agent/drafter.py` (`DEFAULT_PROMPT_FINANCIAL`, `DEFAULT_PROMPT_GMAIL`, `DEFAULT_PROMPT_PERSONAL`) if not yet saved. Personal prompt: warm, direct, no compliance framing, casual sign-off acceptable.
- Draft generation provider chain: Claude (`claude-sonnet-4-6`) primary → OpenAI (`gpt-4o`) fallback on any exception → empty string if both fail. Same prompt sent to whichever provider runs. Provider used logged at INFO level.
- Email footers appended to every draft — read from `settings` table (`footer_financial` / `footer_gmail`). No .env fallback.
- Learning loop — sent emails (3 months, 300+ char body) synthesised into a voice profile per account via Claude. Runs for financial, personal, and gmail. Stored in `voice_profiles` table. Manual trigger: `POST /api/voice/build`. Re-run to refresh.
- Editable To field in draft panel — pre-populated from sender, editable before sending
- Multiple recipients — separate with `,` or `;` in the To field. Both connectors handle arrays.
- Instruction field for draft guidance — type e.g. "make it shorter, focus on the CGT implications" before clicking Generate/Regenerate. Passed as additional constraint to Claude on top of base prompt and voice profile.
- Three-mode dashboard — Inbox, History, Settings (⚙)
- Settings — Voice Profile section: build status for all three accounts, Build and Refresh buttons, polls every 10s after build until all three complete
- Settings — Account Configuration section: per-account base prompt textarea + footer textarea + Save button for Financial Planning (Intertek), ***TAX_BUSINESS_NAME***, and Personal. Loaded from DB on entering Settings mode.
- `GET /api/settings` — returns all 6 settings (DB values or hardcoded defaults): `prompt_financial`, `prompt_gmail`, `prompt_personal`, `footer_financial`, `footer_gmail`, `footer_personal`
- `POST /api/settings` — saves any of the 6 keys (allowlist enforced)
- Queue architecture: `get_queue()` in `db/database.py` fetches ALL pending emails (no limit — inbox must always be complete), and up to 500 history records. Combined and sorted by classification priority. `api_queue()` in `main.py` calls `get_queue()` with no arguments. The old `limit=300` argument was removed.
- Sort buttons toggle direction on repeated click. Priority defaults ascending (Action Required at top). Date defaults descending (newest first). Arrow in button label shows current direction. Both have tiebreakers.
- Manual folder filing: "📁 File…" button on every inbox email opens a searchable folder picker grouped by account (Financial Planning, Personal, Positive Tax). Folders fetched from all three accounts simultaneously via `GET /api/folders`. Filing calls `POST /api/email/{id}/file`, moves email on the server, records the action in `filing_history` SQLite table, marks status as `filed`. Filed emails appear under History → 📁 Filed tab.
- Smart filing suggestions: after the first filing, future opens of the picker show a "Suggested" section at top for emails from the same sender domain, ranked by filing count. Powered by `GET /api/folders/suggestions?sender={raw_sender}`.
- Cross-account filing supported — MIME export/import pipeline. Source archived after successful copy. Graph imports land in draft state (API limitation — content intact).
- Account filter — All / Financial Planning / Positive Tax / Personal
- Classification filter within account in inbox mode
- History sub-tabs — Sent, Archived, Deleted, All
- Sent — Send Follow-up with editable To field
- Archived — Move to Inbox (live unarchive), Delete permanently, Clear Record
- Deleted — Restore to Queue (stub), Clear Record
- Clear Record — deletes single row by email_id (not bulk)
- Clear History — bulk by scope with confirmation
- Stub items — purple tag and notice banner
- Soft delete — Deleted Items (M365) / Trash (Gmail), recoverable 30 days
- Hard delete — autonomous spam only
- No confirmation dialogs on send or delete
- Resizable panels, auto-advance, all actions logged to SQLite
- Auth error warning: amber ⚠ triangle appears next to account name in sidebar when a poll cycle fails with an auth error (401, invalid_grant, etc.). Clicking it opens `http://localhost:8000/auth/login/{account}` in a new tab to re-authenticate. Personal account triangle tooltip directs to rotate client secret in .env instead. Triangle clears automatically on next successful poll. Powered by `GET /api/auth-errors`, `set_auth_error()` / `clear_auth_error()` in `db/database.py`, auth keyword detection in `agent/poller.py`.

- **User overrides stick permanently** — emails restored to inbox by the user are never re-actioned. The stable `internetMessageId` header is used as the DB key; Graph's folder-scoped `id` is used only for API move/delete calls. Graph accounts (financial, personal) send both IDs; Gmail IDs were already stable.
- **DB auto-pruning** — `prune_old_records(days=90)` runs after every full poll cycle (all three accounts). Deletes non-pending `action_log` records older than 90 days. `sender_rules` table is exempt. Logged at INFO level when records are deleted.
- **Sender classification rules** — `sender_rules` table stores one rule per sender email address. Learned rules (source='learned') strengthen on consistent LLM results and fire at count ≥ 2; conflicting results decay the count; eroded to zero deletes the rule. Manual rules (source='manual') fire immediately and are never overwritten by LLM data — only by another manual reclassify or explicit deletion. Created automatically when a user reclassifies an email via the UI.
- **Sender rule management endpoints** — `GET /api/sender-rules` lists all rules; `DELETE /api/sender-rules/{sender}` (URL-encode `@` as `%40`) clears one rule; `DELETE /api/sender-rules` clears all rules.
- **Sender Rules UI** — Settings panel has a new "Sender Rules" section (between AI & Voice and Account Configuration). Shows a table: sender address, classification badge, type badge (learned/manual), count (`—` for manual), last seen date, Delete button per row. Clear All Rules button at the bottom. Refreshes automatically on entering Settings. Loaded via `loadSenderRules()` called from `setMode('settings')`.

**Next session — start here:**
1. `email-agent start`
2. Confirm dashboard at `http://<TAILSCALE_IP>:8000`
3. Hit Poll Now — verify inbox counts match Outlook and Gmail
4. Choose next item from Section 12

**Known issues / gaps:**
- `GEMINI_API_KEY` not yet set (deferred — Claude + OpenAI fallback covers all draft needs)
- **Silent deletion failure** — `api_delete` catches all exceptions and returns `{"ok": True}` even if the Graph/Gmail delete call failed. DB status is still updated to 'deleted'. Low frequency but misleading.
- **Hardcoded folder names** — "Junk Email", "Newsletters", "Notifications" in `actions.py`. If a mailbox uses different names, autonomous moves silently fail (email stays pending). Not a current problem.
- **Send flow unconfirmed (Gmail)** — user reported clicking Send on a Gmail email produced no outbox entry. DB shows `status='sent'` so the code path completed without exception. Root cause unconfirmed; user was going to retry.
- **`stub` column** — exists in schema and filters queries (`WHERE stub = 0`) but is never set to 1. Dead column, harmless.

**All files:**
- `~/email-agent/email-agent` (also `/opt/homebrew/bin/email-agent`)
- `~/email-agent/auth_proxy.py`
- `~/email-agent/main.py`
- `~/email-agent/agent/__init__.py`
- `~/email-agent/agent/actions.py`
- `~/email-agent/agent/classifier.py`
- `~/email-agent/agent/drafter.py`
- `~/email-agent/agent/learner.py`
- `~/email-agent/agent/poller.py`
- `~/email-agent/db/__init__.py`
- `~/email-agent/db/database.py`
- `~/email-agent/connectors/graph.py`
- `~/email-agent/connectors/gmail.py`
- `~/email-agent/ui/templates/dashboard.html`
- `~/email-agent/config/.env`
- `~/email-agent/config/agent.db` (auto-generated)
- `~/email-agent/config/tokens_graph.json` (auto-generated)
- `~/email-agent/config/tokens_gmail.json` (auto-generated)

**Technical notes:**
- `deleteditems` is the M365 well-known folder name for soft delete (no space, lowercase)
- Gmail soft delete: add TRASH label, remove INBOX label
- M365 pagination: follow `@odata.nextLink`. Gmail pagination: follow `nextPageToken`
- Unique index on email_id — `INSERT OR IGNORE` used in log_action
- `extract_email_addresses()` in main.py parses one or more addresses, splits on `,` or `;`, handles "Display Name <email>" form, returns a list
- Graph `send_email()` accepts `to` as `str | list` — builds `toRecipients` array
- Gmail `send_email()` accepts comma-separated string in MIME To header natively
- `extractEmailFromSender()` in dashboard.html parses sender string client-side
- Voice profile: `sent_examples` table caches raw emails, `voice_profiles` table stores synthesised profile (one row per account, upserted). Profile injected into draft prompt via `VOICE_PROFILE_BLOCK` in `agent/drafter.py`.
- Draft guidance: optional `guidance` field in `GenerateDraftRequest` appended to prompt as "Additional instruction: …"
- Settings table: key/value store in SQLite. Keys: `prompt_financial`, `prompt_gmail`, `prompt_personal`, `footer_financial`, `footer_gmail`, `footer_personal`. Read by `get_setting()`, written by `set_setting()`. `GET /api/settings` returns all 6 (DB or defaults). `POST /api/settings` saves one key at a time (allowlist enforced).
- `.env` rule: credentials, API keys, IDs, secrets, and config values only. Footers and prompts are in SQLite. Never store operational content in `.env`.
- Personal account auth: client credentials flow (`grant_type=client_credentials`) using `AZURE_CLIENT_ID_PERSONAL` + `AZURE_CLIENT_SECRET_PERSONAL`. No user sign-in. Token cached as `personal_app` in `tokens_graph.json`. All Graph calls use `/users/<PERSONAL_EMAIL>/...`. If 403 after permission changes, delete `personal_app` key from `tokens_graph.json` to force fresh token acquisition.
- Personal account inbox filter shows with yellow dot in sidebar. Account label is `personal` in SQLite.
- `get_queue(history_limit=500)` signature: pending fetched with no LIMIT clause; history fetched with `LIMIT history_limit`. Do not pass a `limit` keyword — the old param is gone.
- `filing_history` table: `(sender_domain, target_folder_id)` unique index. `record_filing()` upserts and increments count. `get_filing_suggestions(sender_domain, limit=5)` returns top targets ordered by count DESC.
- **Stable email ID (Graph accounts):** `get_emails()` in `graph.py` requests `internetMessageId` in `$select`. In `poll_financial` and `poll_personal`: `stable_id = email.get("internetMessageId") or graph_id` is stored in `action_log.email_id` and used for all DB operations; `graph_id = email.get("id")` is passed to `execute_action` and all move/delete API calls. This means user overrides survive the email being moved and restored — the stable_id never changes.
- **DB pruning:** `prune_old_records(days=90)` in `db/database.py` deletes `action_log` records WHERE `status != 'pending' AND timestamp < cutoff`. Called from `poll_all()` after all three pollers complete. Does not touch `sender_rules`, `sent_examples`, `voice_profiles`, `filing_history`, or `settings`.
- **`sender_rules` table:** `sender TEXT PRIMARY KEY, classification TEXT, count INTEGER, source TEXT, created_at TEXT, last_seen TEXT`. Source is always `'manual'` — AI results are never stored. Never pruned. `upsert_sender_rule(sender, classification, source='manual')`: same classification → count+1; different classification → count resets to 1. `get_sender_rule(sender)` returns the row or None. In each poller: check rule before LLM; if `rule['source']=='manual' and rule['count']>=2` → use rule, skip LLM. count=1 is "pending" (no effect). Classifications of 'error', 'unknown', or '' are never stored. `POST /api/email/{id}/reclassify` calls `upsert_sender_rule(..., source='manual')`. UI shows count as "pending" (count=1) or "N ✓ active" (count≥2).
- `graph.list_folders(account)` — fetches `mailFolders?$expand=childFolders`, returns flat list including one level of subfolders as "Parent / Child". Sorted alphabetically.
- `gmail.list_labels()` — returns user-created labels only; excludes all `_GMAIL_SYSTEM_LABELS` and any `CATEGORY_` prefixed IDs.
- `graph.file_to_folder(account, email_id, folder_id)` and `gmail.file_to_label(email_id, label_id)` — move by ID directly, no name lookup. Distinct from `move_email()` which takes a folder name and creates if missing.
- Filing domain extraction: uses `extract_email_addresses()` to parse sender, then splits on `@` to get domain for `filing_history` key.
- Sort state: `sortDirs = { classification: 'asc', date: 'desc' }`. `setSort(mode)` toggles direction if mode is already active, else switches mode keeping its stored direction. `updateSortButtons()` syncs button text and active class.
- `auth_proxy.py` — asyncio TCP proxy listening on `127.0.0.1:8000`, forwards to `TAILSCALE_IP:8000`. Required because Azure OAuth only allows HTTP for localhost redirect URIs. Started by `email-agent start` using `nohup ... & disown` so it survives script exit. Stopped by `pgrep -f "auth_proxy.py"`. Health-checked with `kill -0` after 1s. PID saved in `logs/auth_proxy.pid`.

---

## 12. Known issues and next session work items

### 12.1 — Personal M365 account ✅ Connected

`<PERSONAL_EMAIL>` — confirmed real member user in the Intertek Entra ID tenant (Object ID: `<AZURE_OBJECT_ID_PERSONAL>`). Has Exchange Online Plan 1. Highest volume inbox of the three accounts.

**Attempts made (previous session — did not succeed):**
1. **Separate certificate for personal app registration** — attempted in Azure portal, was challenged during creation and did not complete.
2. **Financial app registration acting as global admin** — the financial account is a global admin and theoretically has delegate access to all mailboxes in the tenant. Attempted to use this to read the personal mailbox. Got `AADSTS50058` and other errors. Could not pull emails.
3. Account parked after both approaches failed.

**Error AADSTS50058:** "A silent sign-in request was sent but no user is signed in." Indicates the app attempted silent/cached token acquisition for the personal user but no session existed — interactive auth was needed but not triggered correctly, or MFA/Conditional Access blocked it.

**Resolution:** Application permissions (client credentials flow) on the `email-agent-personal` app registration. `Mail.ReadWrite`, `Mail.Send`, `Calendars.ReadWrite` added as Application type permissions with admin consent granted in Azure portal. Code uses `grant_type=client_credentials` — no user sign-in or OAuth flow required. All Graph calls use `/users/<PERSONAL_EMAIL>/...` instead of `/me/...`. Token cached in `tokens_graph.json` under `personal_app` key, auto-refreshed when expired. Confirmed working 2026-05-08.

**Note:** `get_user_profile("personal")` returns `displayName: null` because `User.Read.All` application permission is not granted — this only affects the `/auth/test/personal` display page, not any dashboard functionality.

### 12.2 — Calendar live actions ✅ Complete
M365: `POST /me/messages/{id}/accept` and `/decline` via Graph API — response sent to organiser automatically, `sendResponse: true`.
Gmail: extracts ICS from email payload (`text/calendar` MIME part), parses UID, finds event in Google Calendar via `iCalUID` query, patches attendee `responseStatus` with `sendUpdates: all`.
Both endpoints fall back gracefully — if the live call fails, SQLite status still updates and a warning is returned to the UI. Toast shows "Accepted (local only) — {reason}" on fallback.

### 12.3 — OpenAI fallback ✅ Complete. Gemini deferred.
OpenAI (`gpt-4o`) wired as fallback in `agent/drafter.py` — tried automatically if Claude throws any exception. Gemini skipped for now.

### 12.4 — Phase 6 — Learning loop ✅ Complete
Sent emails (3 months, min 300 chars) fetched from both accounts and stored in `sent_examples` table. Claude synthesises a compact voice profile per account from up to 50 most recent examples (400 chars each). Profile stored in `voice_profiles` table (one row per account, upserted on each run). Injected into draft prompt via `VOICE_PROFILE_BLOCK` in `agent/drafter.py`. No scheduler — trigger manually via `POST /api/voice/build` whenever a refresh is wanted. Signature stripped before storage via `strip_signature()` in `agent/learner.py`.

### 12.5 — Email footer / signature blocks ✅ Complete

Footers stored in SQLite `settings` table as `footer_financial` and `footer_gmail`. Appended by `generate_draft()` after the LLM response. Editable via ⚙ Settings → Account Configuration in the dashboard — changes take effect on the next draft with no restart. No `.env` involvement.

### 12.6 — Manual folder filing with smart suggestions ✅ Complete (same-account)

**Goal:** Add a "File" action as the last option on every email (after all other actions). Opens a folder picker that lets the user move the email into any folder — including folders in the other accounts.

**UI behaviour:**

- "File…" button appears last in the action row for every classification
- Clicking it opens a folder picker panel or modal (not a browser select — a searchable list)
- Folders are grouped by account: Financial Planning (Intertek) / ***TAX_BUSINESS_NAME*** (Gmail) / Personal
- Smart suggestions appear at the top: folders where similar emails have been filed before (based on sender domain + subject keywords, stored in SQLite)
- User selects a target folder → email is moved and marked `filed` in the action log

**Same-account filing:**

- M365: `POST /messages/{id}/move` with destination folder ID (already implemented as `move_email()` in `graph.py`)
- Gmail: `POST /messages/{id}/modify` — add target label, remove INBOX label (already implemented as `move_email()` in `gmail.py`)
- Need a folder/label listing endpoint per account: `GET /me/mailFolders` (Graph, recursive) and `GET /labels` (Gmail)

**Cross-account filing:**

- Not a native move — different providers. Approach: fetch full email content from source, create it in target account's folder, archive/delete from source
- M365 → M365 (financial ↔ personal): use Graph `MIME` export (`GET /messages/{id}/$value`) and `POST /messages` with raw MIME to target mailbox folder
- M365 → Gmail or Gmail → M365: same MIME approach — export and re-import. Treat as a copy+archive, not a true move. Fidelity may not be perfect (headers, threading).
- Cross-account filing is lower priority than same-account — implement same-account first

**Smart filing memory (SQLite):**

- New table: `filing_history` — columns: `sender_domain TEXT`, `subject_keywords TEXT`, `target_account TEXT`, `target_folder_id TEXT`, `target_folder_name TEXT`, `filed_at TEXT`, `count INTEGER`
- On each manual file action: upsert a row keyed on `(sender_domain, target_folder_id)`, incrementing `count`
- On folder picker open: query top 5 `target_folder_name` entries ordered by `count DESC` where `sender_domain` matches current email sender — surface these as "Suggested" at top of picker
- Subject keywords as secondary signal (optional refinement after basic version works)

**New API endpoints needed:**

- `GET /api/folders` — returns folder tree for all three accounts, grouped by account. Calls `list_folders()` on each connector.
- `POST /api/email/{id}/file` — body: `{target_account, folder_id, folder_name}`. Performs the move and logs to `filing_history`.

**New connector functions needed:**

- `graph.py`: `list_folders(account)` — recursive `GET /mailFolders?$expand=childFolders`
- `gmail.py`: `list_labels()` — `GET /labels` filtered to user-created labels only (exclude system labels like INBOX, SENT, TRASH)

**Implementation status:**

1. ✅ `filing_history` table in `db/database.py` — `record_filing()` + `get_filing_suggestions()`
2. ✅ `list_folders(account)` in `graph.py`, `list_labels()` in `gmail.py`
3. ✅ `GET /api/folders` — parallel fetch from all 3 accounts with exception isolation
4. ✅ `GET /api/folders/suggestions?sender=` — top 5 by filing count for sender domain
5. ✅ `POST /api/email/{id}/file` — same-account move + filing_history record
6. ✅ Folder picker modal in dashboard — searchable, grouped by account, session-cached
7. ✅ Smart suggestions shown at top of picker after first filing for a sender domain
8. ✅ Cross-account filing — MIME export/import. Graph↔Graph: $value export + mailFolders/{id}/messages import. Graph↔Gmail: Graph $value export + Gmail uploadType=media import + label apply. Gmail→Graph: format=raw export + base64url decode + Graph MIME import. Source archived after successful copy. Note: Graph imports land in draft state (API limitation).

### 12.7 — GitHub backup ✅ Complete

Private repo: **https://github.com/awolfend/email-agent**

Sensitive files excluded by `.gitignore`: `config/.env`, `config/tokens_graph.json`, `config/tokens_gmail.json`, `config/agent.db`, `logs/`, `__pycache__/`, `.DS_Store`.

**To push changes after each session:**
```
git add agent/ connectors/ db/ ui/ main.py email-agent email-agent-project-reference.md .gitignore
git commit -m "..."
git push
```

Never commit anything from `config/` — secrets and tokens live there.

### 12.8 — Phase 7 — Telegram integration (Priority 5)
Mobile channel for reviewing and actioning emails while travelling. Library: `python-telegram-bot`.

### 12.9 — Mac mini migration (When hardware arrives)
1. Update `TAILSCALE_IP` in `config/.env`
2. Confirm Ollama menu bar app starts at login
3. Re-authenticate both OAuth accounts (tokens are machine-bound)
4. Copy `email-agent` script to `/opt/homebrew/bin/` with `sudo`
5. Test full stack before decommissioning MacBook Pro as server
