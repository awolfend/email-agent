import asyncio
from contextlib import asynccontextmanager
from fastapi import FastAPI, Request
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import RedirectResponse, HTMLResponse, JSONResponse
from pydantic import BaseModel
import uvicorn
import re
import os
from dotenv import load_dotenv
from connectors.graph import (
    get_auth_url as graph_auth_url,
    exchange_code_for_token as graph_exchange,
    get_emails as graph_get_emails,
    get_user_profile as graph_get_profile,
    delete_email as graph_delete_email,
    archive_email as graph_archive_email,
    unarchive_email as graph_unarchive_email,
    get_message_graph_id as graph_get_message_id,
    send_email as graph_send_email,
    reply_to_email as graph_reply_to_email,
    mark_as_read as graph_mark_as_read,
    accept_calendar_event as graph_accept_calendar,
    decline_calendar_event as graph_decline_calendar,
    list_folders as graph_list_folders,
    file_to_folder as graph_file_to_folder,
    export_mime as graph_export_mime,
    import_mime as graph_import_mime,
)
from connectors.gmail import (
    get_auth_url as gmail_auth_url,
    exchange_code_for_token as gmail_exchange,
    get_emails as gmail_get_emails,
    get_user_profile as gmail_get_profile,
    delete_email as gmail_delete_email,
    archive_email as gmail_archive_email,
    unarchive_email as gmail_unarchive_email,
    send_email as gmail_send_email,
    mark_as_read as gmail_mark_as_read,
    accept_calendar_event as gmail_accept_calendar,
    decline_calendar_event as gmail_decline_calendar,
    list_labels as gmail_list_labels,
    file_to_label as gmail_file_to_label,
    export_mime as gmail_export_mime,
    import_mime as gmail_import_mime,
)
from db.database import (
    init_db, get_queue, get_stats,
    update_email_status, update_email_classification,
    update_draft_reply, toggle_flag, get_email_by_id,
    clear_history, delete_record, get_voice_profile_meta,
    get_setting, set_setting,
    record_filing, get_filing_suggestions,
    get_auth_errors,
    upsert_sender_rule, get_all_sender_rules,
    delete_sender_rule, clear_all_sender_rules,
)
from agent.poller import start_scheduler, poll_all
from agent.drafter import generate_draft
from agent.learner import build_voice_profiles
from connectors.hubspot import get_contact_context as hubspot_context
from connectors.graph import get_email_history as graph_email_history
from connectors.gmail import get_email_history as gmail_email_history

load_dotenv("config/.env")


def extract_email_addresses(raw: str) -> list[str]:
    """Parse one or more addresses from a To field. Accepts comma or semicolon separators.
    Handles 'Display Name <email>' and plain address forms."""
    if not raw:
        return []
    addresses = []
    for part in re.split(r'[,;]', raw):
        part = part.strip()
        if not part:
            continue
        match = re.search(r'<([^>]+)>', part)
        addresses.append(match.group(1).strip() if match else part)
    return [a for a in addresses if a]


@asynccontextmanager
async def lifespan(app: FastAPI):
    await init_db()
    start_scheduler()
    yield


app = FastAPI(title="Email Agent", lifespan=lifespan)
templates = Jinja2Templates(directory="ui/templates")
app.mount("/static", StaticFiles(directory="ui/static"), name="static")


class StatusUpdate(BaseModel):
    status: str

class ClassificationUpdate(BaseModel):
    classification: str

class DraftUpdate(BaseModel):
    draft: str

class SendRequest(BaseModel):
    to: str
    subject: str
    body: str

class ClearHistoryRequest(BaseModel):
    scope: str  # 'sent' | 'archived' | 'deleted' | 'all'

class GenerateDraftRequest(BaseModel):
    guidance: str = ""

class SettingUpdate(BaseModel):
    key: str
    value: str

class FileRequest(BaseModel):
    target_account: str
    folder_id: str
    folder_name: str


@app.get("/")
async def dashboard(request: Request):
    return templates.TemplateResponse(request, "dashboard.html")

@app.get("/api/queue")
async def api_queue():
    return JSONResponse(await get_queue())

@app.get("/api/stats")
async def api_stats():
    return JSONResponse(await get_stats())

@app.get("/api/auth-errors")
async def api_auth_errors():
    return JSONResponse(await get_auth_errors())

@app.post("/api/email/{email_id}/status")
async def api_status(email_id: str, body: StatusUpdate):
    await update_email_status(email_id, body.status)
    return {"ok": True}

@app.post("/api/email/{email_id}/reclassify")
async def api_reclassify(email_id: str, body: ClassificationUpdate):
    await update_email_classification(email_id, body.classification)
    email = await get_email_by_id(email_id)
    if email and email.get("sender"):
        await upsert_sender_rule(email["sender"], body.classification, source="manual")
    return {"ok": True}

@app.post("/api/email/{email_id}/draft")
async def api_draft(email_id: str, body: DraftUpdate):
    await update_draft_reply(email_id, body.draft)
    return {"ok": True}

@app.post("/api/email/{email_id}/flag")
async def api_flag(email_id: str):
    await toggle_flag(email_id)
    return {"ok": True}

@app.post("/api/email/{email_id}/generate-draft")
async def api_generate_draft(email_id: str, body: GenerateDraftRequest = GenerateDraftRequest()):
    email = await get_email_by_id(email_id)
    if not email:
        return JSONResponse({"draft": ""})
    account = email["account"]
    sender  = email["sender"] or ""

    crm_context = ""
    if sender:
        # Gather CRM context (financial only) + email history (all accounts) in parallel
        if account == "financial":
            hs_ctx, email_history = await asyncio.gather(
                hubspot_context(sender),
                graph_email_history("financial", sender, limit=8),
            )
        elif account == "personal":
            hs_ctx        = ""
            email_history = await graph_email_history("personal", sender, limit=8)
        else:  # gmail
            hs_ctx        = ""
            email_history = await gmail_email_history(sender, limit=8)

        parts = []
        if hs_ctx:
            parts.append(hs_ctx)
        if email_history:
            history_lines = ["--- Email history ---"]
            for msg in email_history:
                subject = msg["subject"] or "(no subject)"
                snippet = f" — {msg['snippet']}" if msg["snippet"] else ""
                history_lines.append(f"  [{msg['date']}] {msg['direction']} {subject}{snippet}")
            history_lines.append("--- end email history ---")
            parts.append("\n".join(history_lines))
        crm_context = "\n\n".join(parts)

    draft = await generate_draft(
        account=account,
        sender=sender,
        subject=email["subject"] or "",
        body=email["body"] or "",
        guidance=body.guidance,
        crm_context=crm_context,
    )
    if draft:
        await update_draft_reply(email_id, draft)
    return JSONResponse({"draft": draft, "crm_context_loaded": bool(crm_context)})

@app.post("/api/email/{email_id}/send")
async def api_send(email_id: str, body: SendRequest):
    email = await get_email_by_id(email_id)
    if not email:
        return JSONResponse({"ok": False, "error": "Email not found"}, status_code=404)
    recipients = extract_email_addresses(body.to)
    if not recipients:
        return JSONResponse({"ok": False, "error": "No valid recipient address"}, status_code=400)
    to = ", ".join(recipients)
    graph_id = email.get("graph_id") or email_id
    try:
        if email["account"] in ("financial", "personal"):
            _filing_email = os.getenv("FILING_EMAIL_FINANCIAL")
            cc = [_filing_email] if email["account"] == "financial" and _filing_email else None
            await graph_reply_to_email(email["account"], graph_id, recipients, body.body, cc=cc)
            await graph_archive_email(email["account"], graph_id)
        elif email["account"] == "gmail":
            await gmail_send_email(
                to, body.subject, body.body,
                thread_id=email.get("thread_id"),
                in_reply_to=email.get("orig_message_id"),
            )
            await gmail_archive_email(email_id)
        else:
            return JSONResponse({"ok": False, "error": f"Unknown account: {email['account']}"}, status_code=400)
        await update_email_status(email_id, "sent", "replied")
        return {"ok": True}
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)

@app.post("/api/email/{email_id}/send-followup")
async def api_send_followup(email_id: str, body: SendRequest):
    email = await get_email_by_id(email_id)
    if not email:
        return JSONResponse({"ok": False, "error": "Email not found"}, status_code=404)
    recipients = extract_email_addresses(body.to)
    if not recipients:
        return JSONResponse({"ok": False, "error": "No valid recipient address"}, status_code=400)
    to = ", ".join(recipients)
    try:
        if email["account"] in ("financial", "personal"):
            await graph_send_email(email["account"], recipients, body.subject, body.body)
        elif email["account"] == "gmail":
            await gmail_send_email(to, body.subject, body.body)
        await update_draft_reply(email_id, body.body)
        return {"ok": True}
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)

@app.post("/api/email/{email_id}/archive")
async def api_archive(email_id: str):
    email = await get_email_by_id(email_id)
    if not email:
        return JSONResponse({"ok": False, "error": "Email not found"}, status_code=404)
    graph_id = email.get("graph_id") or email_id
    try:
        if email["account"] in ("financial", "personal"):
            await graph_mark_as_read(email["account"], graph_id)
            await graph_archive_email(email["account"], graph_id)
        elif email["account"] == "gmail":
            await gmail_mark_as_read(email_id)
            await gmail_archive_email(email_id)
        await update_email_status(email_id, "archived", "archived")
        return {"ok": True}
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)

@app.post("/api/email/{email_id}/unarchive")
async def api_unarchive(email_id: str):
    email = await get_email_by_id(email_id)
    if not email:
        return JSONResponse({"ok": False, "error": "Email not found"}, status_code=404)
    try:
        if email["account"] in ("financial", "personal"):
            # Stored graph_id is the inbox-era ID and is stale after archive.
            # Look up the current folder-scoped ID via the stable internetMessageId.
            live_id = await graph_get_message_id(email["account"], email_id)
            if not live_id:
                raise Exception("Message not found in mailbox — may have been permanently deleted")
            await graph_unarchive_email(email["account"], live_id)
        elif email["account"] == "gmail":
            await gmail_unarchive_email(email_id)
        await update_email_status(email_id, "pending")
        return {"ok": True}
    except Exception as e:
        await update_email_status(email_id, "pending")
        return {"ok": True, "warning": str(e)}

@app.post("/api/email/{email_id}/calendar/accept")
async def api_calendar_accept(email_id: str):
    email = await get_email_by_id(email_id)
    if not email:
        return JSONResponse({"ok": False, "error": "Email not found"}, status_code=404)
    warning = None
    try:
        if email["account"] in ("financial", "personal"):
            graph_id = email.get("graph_id") or email_id
            await graph_accept_calendar(email["account"], graph_id)
        elif email["account"] == "gmail":
            await gmail_accept_calendar(email_id)
    except Exception as e:
        warning = str(e)
    await update_email_status(email_id, "approved", "calendar_accepted")
    return {"ok": True, **({"warning": warning} if warning else {})}


@app.post("/api/email/{email_id}/calendar/decline")
async def api_calendar_decline(email_id: str):
    email = await get_email_by_id(email_id)
    if not email:
        return JSONResponse({"ok": False, "error": "Email not found"}, status_code=404)
    warning = None
    try:
        if email["account"] in ("financial", "personal"):
            graph_id = email.get("graph_id") or email_id
            await graph_decline_calendar(email["account"], graph_id)
        elif email["account"] == "gmail":
            await gmail_decline_calendar(email_id)
    except Exception as e:
        warning = str(e)
    await update_email_status(email_id, "rejected", "calendar_declined")
    return {"ok": True, **({"warning": warning} if warning else {})}


@app.get("/api/folders")
async def api_list_folders():
    import asyncio
    results = await asyncio.gather(
        graph_list_folders("financial"),
        graph_list_folders("personal"),
        gmail_list_labels(),
        return_exceptions=True,
    )
    return JSONResponse({
        "financial": results[0] if not isinstance(results[0], Exception) else [],
        "personal": results[1] if not isinstance(results[1], Exception) else [],
        "gmail": results[2] if not isinstance(results[2], Exception) else [],
    })


@app.get("/api/folders/suggestions")
async def api_filing_suggestions(sender: str = ""):
    domain = ""
    if "@" in sender:
        addrs = extract_email_addresses(sender)
        if addrs and "@" in addrs[0]:
            domain = addrs[0].split("@")[-1].lower()
    suggestions = await get_filing_suggestions(domain) if domain else []
    return JSONResponse(suggestions)


@app.post("/api/email/{email_id}/file")
async def api_file(email_id: str, body: FileRequest):
    email = await get_email_by_id(email_id)
    if not email:
        return JSONResponse({"ok": False, "error": "Email not found"}, status_code=404)
    if body.target_account not in ("financial", "personal", "gmail"):
        return JSONResponse({"ok": False, "error": f"Unknown account: {body.target_account}"}, status_code=400)

    source = email["account"]
    target = body.target_account
    graph_id = email.get("graph_id") or email_id

    try:
        if source == target:
            # Same-account: direct move
            if target in ("financial", "personal"):
                await graph_file_to_folder(target, graph_id, body.folder_id)
            else:
                await gmail_file_to_label(email_id, body.folder_id)
        else:
            # Cross-account: export MIME from source, import into target, archive source
            if source in ("financial", "personal"):
                mime_bytes = await graph_export_mime(source, graph_id)
            else:
                mime_bytes = await gmail_export_mime(email_id)

            if target in ("financial", "personal"):
                await graph_import_mime(target, body.folder_id, mime_bytes)
            else:
                await gmail_import_mime(body.folder_id, mime_bytes)

            # Archive source after successful import
            if source in ("financial", "personal"):
                await graph_archive_email(source, graph_id)
            else:
                await gmail_archive_email(email_id)

        sender_addrs = extract_email_addresses(email.get("sender", ""))
        if sender_addrs and "@" in sender_addrs[0]:
            domain = sender_addrs[0].split("@")[-1].lower()
            await record_filing(domain, target, body.folder_id, body.folder_name)
        await update_email_status(email_id, "filed", "filed")
        return {"ok": True}
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e)}, status_code=500)


@app.post("/api/email/{email_id}/restore")
async def api_restore(email_id: str):
    email = await get_email_by_id(email_id)
    if not email:
        return JSONResponse({"ok": False, "error": "Email not found"}, status_code=404)
    await update_email_status(email_id, "pending")
    return {"ok": True}

@app.post("/api/email/{email_id}/delete")
async def api_delete(email_id: str):
    email = await get_email_by_id(email_id)
    if email:
        warning = None
        try:
            if email["account"] in ("financial", "personal"):
                graph_id = email.get("graph_id") or email_id
                await graph_delete_email(email["account"], graph_id)
            elif email["account"] == "gmail":
                await gmail_delete_email(email_id)
        except Exception as e:
            warning = str(e)
        await update_email_status(email_id, "deleted", "deleted")
        return {"ok": True, **({"warning": warning} if warning else {})}
    return {"ok": True}

@app.delete("/api/email/{email_id}/record")
async def api_delete_record(email_id: str):
    """Delete a single history record from SQLite. Does not affect the live mailbox."""
    await delete_record(email_id)
    return {"ok": True}

@app.post("/api/history/clear")
async def api_clear_history(body: ClearHistoryRequest):
    """Bulk clear history records by scope. Does not affect the live mailbox."""
    count = await clear_history(body.scope)
    return {"ok": True, "cleared": count}

@app.get("/api/sender-rules")
async def api_get_sender_rules():
    return JSONResponse(await get_all_sender_rules())

@app.delete("/api/sender-rules")
async def api_clear_sender_rules():
    count = await clear_all_sender_rules()
    return {"ok": True, "cleared": count}

@app.delete("/api/sender-rules/{sender:path}")
async def api_delete_sender_rule(sender: str):
    from urllib.parse import unquote
    deleted = await delete_sender_rule(unquote(sender))
    return {"ok": deleted}


@app.post("/api/poll")
async def api_poll():
    import asyncio
    asyncio.create_task(poll_all())
    return {"ok": True}

@app.get("/api/settings")
async def api_get_settings():
    from agent.drafter import DEFAULT_PROMPT_FINANCIAL, DEFAULT_PROMPT_GMAIL, DEFAULT_PROMPT_PERSONAL

    async def _get(key: str, fallback: str) -> str:
        val = await get_setting(key)
        return val if val is not None else fallback

    return JSONResponse({
        "prompt_financial": await _get("prompt_financial", DEFAULT_PROMPT_FINANCIAL),
        "prompt_gmail": await _get("prompt_gmail", DEFAULT_PROMPT_GMAIL),
        "prompt_personal": await _get("prompt_personal", DEFAULT_PROMPT_PERSONAL),
        "footer_financial": await _get("footer_financial", ""),
        "footer_gmail": await _get("footer_gmail", ""),
        "footer_personal": await _get("footer_personal", ""),
    })


@app.post("/api/settings")
async def api_save_setting(body: SettingUpdate):
    allowed = {"prompt_financial", "prompt_gmail", "prompt_personal",
               "footer_financial", "footer_gmail", "footer_personal"}
    if body.key not in allowed:
        return JSONResponse({"ok": False, "error": "Unknown setting key"}, status_code=400)
    await set_setting(body.key, body.value)
    return {"ok": True}


@app.get("/api/voice/status")
async def api_voice_status():
    financial = await get_voice_profile_meta("financial")
    gmail = await get_voice_profile_meta("gmail")
    personal = await get_voice_profile_meta("personal")
    return JSONResponse({
        "financial": financial or {"built": False},
        "gmail": gmail or {"built": False},
        "personal": personal or {"built": False},
    })

@app.post("/api/voice/build")
async def api_build_voice():
    import asyncio
    asyncio.create_task(build_voice_profiles())
    return {"ok": True}

@app.get("/auth/login/{account}")
async def login(account: str):
    if account == "gmail":
        return RedirectResponse(gmail_auth_url())
    if account == "financial":
        return RedirectResponse(graph_auth_url(account))
    if account == "personal":
        # Personal uses application credentials — no OAuth flow required
        return RedirectResponse("/auth/test/personal")
    return HTMLResponse("Unknown account", status_code=400)

@app.get("/auth/callback/gmail")
async def gmail_callback(request: Request):
    code = request.query_params.get("code")
    if not code:
        return HTMLResponse("No code returned from Google", status_code=400)
    await gmail_exchange(code)
    return RedirectResponse("/auth/test/gmail")

@app.get("/auth/callback/{account}")
async def graph_callback(account: str, request: Request):
    code = request.query_params.get("code")
    if not code:
        return HTMLResponse("No code returned from Microsoft", status_code=400)
    await graph_exchange(code, account)
    return RedirectResponse(f"/auth/test/{account}")

@app.get("/auth/test/gmail")
async def test_gmail():
    profile = await gmail_get_profile()
    emails = await gmail_get_emails()
    email_list = "".join(f"<li>{e['subject']} — {e['from']}</li>" for e in emails[:5])
    return HTMLResponse(
        f"<h2>Connected: {profile.get('name')} ({profile.get('email')})</h2>"
        f"<h3>Last 5 emails:</h3><ul>{email_list}</ul>"
        f"<br><a href='/'>Back to dashboard</a>"
    )

@app.get("/auth/test/{account}")
async def test_account(account: str):
    profile = await graph_get_profile(account)
    emails = await graph_get_emails(account)
    email_list = "".join(
        f"<li>{e['subject']} — {e['from']['emailAddress']['address']}</li>"
        for e in emails[:5]
    )
    return HTMLResponse(
        f"<h2>Connected: {profile.get('displayName')} ({profile.get('mail')})</h2>"
        f"<h3>Last 5 emails:</h3><ul>{email_list}</ul>"
        f"<br><a href='/'>Back to dashboard</a>"
    )

if __name__ == "__main__":
    uvicorn.run("main:app", host=os.getenv("TAILSCALE_IP", "127.0.0.1"), port=8000, reload=True)
