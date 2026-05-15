# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Server management

The server is managed by launchd (`~/Library/LaunchAgents/com.awolfend.email-agent.plist`, `KeepAlive: true`). Killing the process causes an automatic restart.

```bash
# Restart (apply code changes)
launchctl unload ~/Library/LaunchAgents/com.awolfend.email-agent.plist
launchctl load ~/Library/LaunchAgents/com.awolfend.email-agent.plist

# Run manually for debug output (binds to 127.0.0.1 fallback if Tailscale is down)
cd ~/email-agent && source venv/bin/activate
uvicorn main:app --host 127.0.0.1 --port 8000 --reload

# Verify running
ps aux | grep uvicorn | grep email-agent
curl -s http://100.100.150.128:8000/api/stats

# Trigger a manual poll
curl -s -X POST http://100.100.150.128:8000/api/poll
```

The startup script (`email-agent-server`) resolves the Tailscale IP dynamically via `tailscale ip -4`. The server must **never** bind to `0.0.0.0`.

There is no test suite. Validation is done by running the server and exercising endpoints directly.

## Python environment

```bash
# Always use the venv
source ~/email-agent/venv/bin/activate

# Install deps
pip install -r requirements.txt

# Verify imports after changes
python -c "import main; print('OK')"
```

Python 3.14.5 via Homebrew. Critical Jinja2 compat — always use this form or it raises a TypeError:
```python
# Correct (Python 3.14 + Jinja2 3.1.6)
return templates.TemplateResponse(request, "dashboard.html")

# Wrong — raises TypeError
return templates.TemplateResponse("dashboard.html", {"request": request})
```

## Architecture

### Email ID duality (Graph accounts only)

Every email stored for the financial and personal accounts has **two IDs**:

- `email_id` in `action_log` = `internetMessageId` (RFC 2822 `Message-ID` header) — **stable across all folder moves**. Used as the SQLite primary key and all DB lookups.
- `graph_id` in `action_log` = Graph folder-scoped `id` — **changes when a message is moved**. Used for every Graph API call (move, delete, reply, archive).

Anywhere a Graph API call is made, use `graph_id = email.get("graph_id") or email_id`. After archive/move, the stored `graph_id` is stale — `get_message_graph_id(account, internet_message_id)` re-resolves the current one.

Gmail IDs are stable across label changes so `stable_id == op_id` for Gmail.

### Inbox-as-source-of-truth model

The poll cycle enforces this invariant in `agent/poller.py`:
1. For every email **present** in the live inbox: `ensure_inbox_state()` forces status to `pending` and refreshes `graph_id`, regardless of what the DB says. This means user-restored emails can't be re-actioned.
2. For every email **missing** from the live inbox but `pending` in the DB: `mark_missing_as_archived()` auto-archives it.

This runs on every poll for all three accounts.

### Three accounts, two auth patterns

| Account | Variable | Auth | Base URL |
|---|---|---|---|
| `financial` | `AZURE_CLIENT_ID_FINANCIAL` | Delegated OAuth (user sign-in) | `/me/...` |
| `personal` | `AZURE_CLIENT_ID_PERSONAL` | Client credentials (no sign-in) | `/users/<PERSONAL_EMAIL>/...` |
| `gmail` | `GOOGLE_CLIENT_ID` | Google OAuth (user sign-in) | Gmail REST API |

The personal account token is cached under `personal_app` in `tokens_graph.json`. If it gets a 403 after permission changes, delete that key to force fresh acquisition. `get_valid_token("personal")` calls `get_app_token()` instead of the standard refresh path.

### Request flow for any action

All action endpoints in `main.py` follow this pattern:
1. `get_email_by_id(email_id)` — load from SQLite
2. Resolve `graph_id = email.get("graph_id") or email_id`
3. Branch on `email["account"]` → call the appropriate connector
4. `update_email_status(email_id, ...)` — write result back

### Classification and AI tiers

```
New email → Ollama llama3.1:8b (local) → classify_email()
                ↓ (if sender rule active: manual + count≥2)
           Skip LLM, use rule directly

Draft generation → Claude claude-sonnet-4-6 (primary)
                 → OpenAI gpt-4o (fallback on any exception)
```

`VOICE_PROFILE_BLOCK` in `agent/drafter.py` is a **module-level f-string** with `{{profile}}` (double braces). This is intentional — `{_AUTHOR_FIRST}` is interpolated at load time; `{{profile}}` becomes `{profile}` for `.format(profile=profile)` inside `_get_voice_block()` at call time. Do not change `{{profile}}` to `{profile}`.

### Context assembly for draft generation

`generate_draft()` assembles prompt context from up to four sources:
1. Account base prompt (from `settings` table, or hardcoded defaults in `drafter.py`)
2. Voice profile (from `voice_profiles` table, via `_get_voice_block()`)
3. HubSpot CRM context (from `connectors/hubspot.py` — contact, notes, meetings, tasks)
4. Email history (from `graph.get_email_history()` or `gmail.get_email_history()`)

All four are fetched/assembled in `main.py:api_generate_draft()`. HubSpot and email history are gathered in parallel via `asyncio.gather()`. Failures in any context source are silent — never block draft generation.

### Sender rules

Only manual rules exist (`source='manual'`). Rules are created by user reclassification via the UI. A rule at `count=1` is "pending" (no effect on classification). At `count≥2` it is "active" and bypasses Ollama entirely. Reclassifying to a different category resets count to 1.

### Auth error detection

Each poller catches all exceptions, checks `str(e)` against `_AUTH_KEYWORDS` (in `agent/poller.py`), and calls `set_auth_error(account, message)` which writes to the `settings` table. The dashboard reads these via `GET /api/auth-errors` and shows a ⚠ triangle. The triangle clears automatically on the next successful poll.

### HubSpot CC / filing email

For the financial account, all outgoing email (replies via `api_send` and follow-ups via `api_send_followup`) is CC'd to `FILING_EMAIL_FINANCIAL` (from `.env`). Replies use `graph_reply_to_email(cc=...)`. Follow-ups use `graph_send_email(cc=...)`. The personal account is deliberately excluded — no filing address.

### Single-page frontend

`ui/templates/dashboard.html` is a self-contained SPA (~2,100 lines). All state lives in JS variables (`allEmails`, `currentEmail`, `currentMode`, etc.). No framework — vanilla JS with `fetch()`. The `showToast()` function is the only user feedback mechanism for async operations. Error strings come directly from `data.error` in API responses.

## Key config locations

| File | Purpose |
|---|---|
| `config/.env` | All secrets (never commit) |
| `config/agent.db` | SQLite — all operational data |
| `config/tokens_graph.json` | M365 OAuth tokens (`financial`, `personal_app`) |
| `config/tokens_gmail.json` | Gmail OAuth token |

Settings that live in the DB (not `.env`): per-account prompts (`prompt_financial`, `prompt_gmail`, `prompt_personal`), footers (`footer_financial`, `footer_gmail`, `footer_personal`). Read/write via `get_setting()` / `set_setting()` in `db/database.py`.
