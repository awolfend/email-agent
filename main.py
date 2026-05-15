import asyncio
from contextlib import asynccontextmanager
from datetime import datetime, timedelta, timezone
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
    search_contact_history,
    save_compose_draft, get_compose_draft, clear_compose_draft,
    log_action,
    save_meeting_proposal, save_meeting_slot,
    get_meeting_proposal, get_slots_for_proposal,
    get_open_proposals_for_client,
    update_slot_status, update_proposal_status,
)
from agent.poller import start_scheduler, poll_all
from agent.drafter import generate_draft, generate_compose_draft
from agent.learner import build_voice_profiles
from connectors.hubspot import get_contact_context as hubspot_context, search_contacts as hubspot_search_contacts
from connectors.graph import (
    get_email_history as graph_email_history,
    search_contacts as graph_search_contacts,
    get_busy_windows as graph_get_busy_windows,
    create_calendar_hold as graph_create_hold,
    delete_calendar_event as graph_delete_event,
    create_confirmed_event as graph_create_confirmed,
)
from connectors.gmail import (
    get_email_history as gmail_email_history,
    get_busy_windows as gmail_get_busy_windows,
    create_calendar_hold as gmail_create_hold,
    delete_calendar_event as gmail_delete_event,
    create_confirmed_event as gmail_create_confirmed,
)

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

class ComposeDraftRequest(BaseModel):
    account: str = "financial"
    to: str = ""
    to_email: str = ""
    cc: str = ""
    subject: str = ""
    prompt: str = ""
    meeting_slots: list[str] = []
    duration_minutes: int = 60


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
            _filing_email = os.getenv("FILING_EMAIL_FINANCIAL")
            cc = [_filing_email] if email["account"] == "financial" and _filing_email else None
            await graph_send_email(email["account"], recipients, body.subject, body.body, cc=cc)
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


@app.get("/api/contacts/search")
async def api_contacts_search(q: str = "", account: str = "financial"):
    """
    Search contacts across HubSpot (financial only), M365 contacts, and email history.
    Returns [{ name, email, source }] deduped by email, priority: hubspot > contacts > history.
    Minimum query length: 2 characters.
    """
    if len(q) < 2:
        return JSONResponse([])

    searches = []
    if account == "financial":
        searches.append(hubspot_search_contacts(q))
    if account in ("financial", "personal"):
        searches.append(graph_search_contacts(account, q))
    searches.append(search_contact_history(account, q))

    batches = await asyncio.gather(*searches, return_exceptions=True)

    seen, results = set(), []
    for batch in batches:
        if isinstance(batch, Exception):
            continue
        for r in batch:
            key = r["email"].lower().strip()
            if key and key not in seen:
                seen.add(key)
                results.append(r)

    return JSONResponse(results[:20])


@app.get("/api/compose/draft-state")
async def api_get_compose_draft(account: str = "financial"):
    """Return the saved compose draft for the given account, or null."""
    draft = await get_compose_draft(account)
    return JSONResponse(draft or {})


@app.post("/api/compose/draft-state")
async def api_save_compose_draft(request: Request):
    """Auto-save compose draft state for an account."""
    body = await request.json()
    account = body.get("account", "financial")
    await save_compose_draft(
        account=account,
        to_address=body.get("to_address"),
        cc_address=body.get("cc_address"),
        subject=body.get("subject"),
        body=body.get("body"),
        prompt=body.get("prompt"),
    )
    return {"ok": True}


@app.post("/api/compose/draft")
async def api_compose_draft(body: ComposeDraftRequest):
    """
    Generate a new outgoing email draft (not a reply).
    Assembles CRM + email history context, then calls Claude/OpenAI.
    Returns { draft, subject, context_sources }.
    """
    account   = body.account
    to_email  = body.to_email.strip() or body.to.strip()
    to_display = body.to.strip() or to_email

    context_parts: list[str] = []
    context_sources: list[str] = []

    if to_email:
        if account == "financial":
            hs_ctx, email_history = await asyncio.gather(
                hubspot_context(to_email),
                graph_email_history("financial", to_email, limit=6),
            )
        elif account == "personal":
            hs_ctx        = ""
            email_history = await graph_email_history("personal", to_email, limit=6)
        else:  # gmail
            hs_ctx        = ""
            email_history = await gmail_email_history(to_email, limit=6)

        if hs_ctx:
            context_parts.append(hs_ctx)
            context_sources.append("HubSpot CRM")
        if email_history:
            lines = ["--- Email history ---"]
            for msg in email_history:
                subj    = msg["subject"] or "(no subject)"
                snippet = f" — {msg['snippet']}" if msg["snippet"] else ""
                lines.append(f"  [{msg['date']}] {msg['direction']} {subj}{snippet}")
            lines.append("--- end email history ---")
            context_parts.append("\n".join(lines))
            context_sources.append(f"Email history ({len(email_history)} msgs)")
    else:
        hs_ctx        = ""
        email_history = []

    crm_context = "\n\n".join(context_parts)

    draft, subject = await generate_compose_draft(
        account=account,
        to=to_display,
        subject=body.subject,
        prompt=body.prompt,
        crm_context=crm_context,
        meeting_slots=body.meeting_slots or [],
        duration_minutes=body.duration_minutes,
    )

    if not draft:
        return JSONResponse({"error": "Draft generation failed — check server logs"}, status_code=500)

    # Persist the generated draft so switching accounts doesn't lose it
    await save_compose_draft(
        account=account,
        to_address=body.to,
        cc_address=body.cc,
        subject=subject or body.subject,
        body=draft,
        prompt=body.prompt,
    )

    return JSONResponse({
        "draft": draft,
        "subject": subject,
        "context_sources": context_sources,
    })


class ComposeSendRequest(BaseModel):
    account: str = "financial"
    to: str = ""
    to_email: str = ""
    cc: str = ""
    subject: str
    body: str
    meeting_slots: list[dict] = []   # [{label: str, raw: {start, end} | None}]
    duration_minutes: int = 60


@app.post("/api/compose/send")
async def api_compose_send(body: ComposeSendRequest):
    """
    Send a composed email, create tentative calendar holds for any meeting slots,
    log to action_log, and clear the saved draft.
    """
    account  = body.account
    to_raw   = body.to.strip()
    to_email = body.to_email.strip() or to_raw
    recipients = extract_email_addresses(to_raw) or extract_email_addresses(to_email)
    if not recipients:
        return JSONResponse({"ok": False, "error": "No valid recipient address"}, status_code=400)
    if not body.subject.strip():
        return JSONResponse({"ok": False, "error": "Subject is required"}, status_code=400)
    if not body.body.strip():
        return JSONResponse({"ok": False, "error": "Body is empty"}, status_code=400)

    sent_at = datetime.now(timezone.utc)

    # ---- Send email ----
    try:
        if account in ("financial", "personal"):
            filing_cc = os.getenv("FILING_EMAIL_FINANCIAL")
            cc_list = []
            if account == "financial" and filing_cc:
                cc_list.append(filing_cc)
            if body.cc.strip():
                cc_list.extend(extract_email_addresses(body.cc))
            await graph_send_email(account, recipients, body.subject, body.body,
                                   cc=cc_list or None)
        elif account == "gmail":
            to_str = ", ".join(recipients)
            cc_str = body.cc.strip() or None
            await gmail_send_email(to_str, body.subject, body.body,
                                   cc=cc_str)
        else:
            return JSONResponse({"ok": False, "error": f"Unknown account: {account}"}, status_code=400)
    except Exception as e:
        return JSONResponse({"ok": False, "error": f"Send failed: {e}"}, status_code=500)

    # ---- Log to action_log so it appears in History > Sent ----
    import uuid
    compose_id = f"compose-{uuid.uuid4().hex[:16]}"
    await log_action(
        account=account,
        email_id=compose_id,
        subject=body.subject,
        sender=to_raw,           # store recipient as "sender" field so history shows who it was sent to
        action="sent_compose",
        body=body.body[:2000],
    )
    await update_email_status(compose_id, "sent", "sent_compose")

    # ---- Calendar holds for meeting slots ----
    proposal_id = None
    valid_slots = [s for s in body.meeting_slots if s.get("raw") and s["raw"].get("start") and s["raw"].get("end")]

    if valid_slots:
        no_response_hours = int(await get_setting("meeting_no_response_hours") or "36")
        post_slot_buffer  = int(await get_setting("meeting_post_slot_buffer_minutes") or "30")
        no_response_deadline = sent_at + timedelta(hours=no_response_hours)

        proposal_id = await save_meeting_proposal(
            outgoing_message_id=compose_id,
            account=account,
            client_email=to_email,
            client_name=to_raw,
            subject=body.subject,
            duration_minutes=body.duration_minutes,
            sent_at=sent_at.isoformat(),
        )

        for slot in valid_slots:
            raw   = slot["raw"]
            s_iso = raw["start"]
            e_iso = raw["end"]
            try:
                s_dt = datetime.fromisoformat(s_iso)
                e_dt = datetime.fromisoformat(e_iso)
            except ValueError:
                continue

            # Auto-release = earliest of (slot_end + buffer) and (sent + 36h)
            slot_deadline = e_dt.astimezone(timezone.utc) + timedelta(minutes=post_slot_buffer)
            auto_release  = min(slot_deadline, no_response_deadline).isoformat()

            # Create owning event + mirrors in parallel
            own_id = fin_mirror = tax_mirror = ""
            if account == "financial":
                own_id, tax_mirror = await asyncio.gather(
                    graph_create_hold("financial", s_iso, e_iso),
                    gmail_create_hold(s_iso, e_iso),
                )
            elif account == "gmail":
                tax_mirror, fin_mirror = await asyncio.gather(
                    gmail_create_hold(s_iso, e_iso),
                    graph_create_hold("financial", s_iso, e_iso),
                )
                own_id = tax_mirror
                tax_mirror = ""   # own_id IS the google event; fin_mirror is the M365 mirror
            elif account == "personal":
                own_id, fin_mirror, tax_mirror = await asyncio.gather(
                    graph_create_hold("personal", s_iso, e_iso),
                    graph_create_hold("financial", s_iso, e_iso),
                    gmail_create_hold(s_iso, e_iso),
                )

            await save_meeting_slot(
                proposal_id=proposal_id,
                slot_start=s_iso,
                slot_end=e_iso,
                auto_release_at=auto_release,
                owning_calendar_event_id=own_id or None,
                mirror_event_id_financial=fin_mirror or None,
                mirror_event_id_google_tax=tax_mirror or None,
            )

    # ---- Clear compose draft ----
    await clear_compose_draft(account)

    return JSONResponse({
        "ok": True,
        **({"proposal_id": proposal_id} if proposal_id else {}),
    })


@app.get("/api/meeting/by-sender")
async def api_meeting_by_sender(sender: str = ""):
    """Return all open (pending) meeting proposals for a given sender email address."""
    if not sender:
        return JSONResponse([])
    # Extract raw email address if in "Name <email>" format
    clean = re.search(r'<([^>]+)>', sender)
    email_addr = clean.group(1).strip().lower() if clean else sender.strip().lower()
    proposals = await get_open_proposals_for_client(email_addr)
    results = []
    for p in proposals:
        slots = await get_slots_for_proposal(p["id"])
        results.append({"proposal": p, "slots": slots})
    return JSONResponse(results)


@app.get("/api/meeting/{proposal_id}")
async def api_get_meeting(proposal_id: int):
    """Return a meeting proposal and its slots by proposal ID."""
    proposal = await get_meeting_proposal(proposal_id)
    if not proposal:
        return JSONResponse({"error": "Not found"}, status_code=404)
    slots = await get_slots_for_proposal(proposal_id)
    return JSONResponse({"proposal": proposal, "slots": slots})


async def _delete_slot_holds(slot: dict):
    """Delete all calendar hold events for a slot across owning + mirror calendars. Fails silently."""
    own_id    = slot.get("owning_calendar_event_id")
    fin_id    = slot.get("mirror_event_id_financial")
    tax_id    = slot.get("mirror_event_id_google_tax")
    # Determine which account owns this slot (from proposal, looked up by caller)
    account   = slot.get("_account", "financial")

    coros = []
    if own_id:
        if account == "gmail":
            coros.append(gmail_delete_event(own_id))
        else:
            coros.append(graph_delete_event(account, own_id))
    if fin_id:
        coros.append(graph_delete_event("financial", fin_id))
    if tax_id:
        coros.append(gmail_delete_event(tax_id))
    if coros:
        await asyncio.gather(*coros, return_exceptions=True)


@app.post("/api/meeting/{proposal_id}/confirm/{slot_id}")
async def api_meeting_confirm(proposal_id: int, slot_id: int):
    """
    Confirm a meeting slot:
    1. Create a confirmed calendar event in the owning calendar with the client as attendee.
    2. Delete tentative holds for all other slots (owning + mirrors).
    3. Mark the confirmed slot as 'confirmed', others as 'declined'.
    4. Mark the proposal as 'confirmed'.
    """
    proposal = await get_meeting_proposal(proposal_id)
    if not proposal:
        return JSONResponse({"ok": False, "error": "Proposal not found"}, status_code=404)
    if proposal["status"] != "pending":
        return JSONResponse({"ok": False, "error": f"Proposal is already {proposal['status']}"}, status_code=409)

    slots = await get_slots_for_proposal(proposal_id)
    confirmed_slot = next((s for s in slots if s["id"] == slot_id), None)
    if not confirmed_slot:
        return JSONResponse({"ok": False, "error": "Slot not found"}, status_code=404)

    account      = proposal["account"]
    client_email = proposal["client_email"]
    client_name  = proposal.get("client_name", "")
    # Strip "Name <email>" format if present
    name_match = re.search(r'^(.*?)\s*<[^>]+>$', client_name)
    if name_match:
        client_name = name_match.group(1).strip()

    subject = proposal.get("subject", "Meeting")

    # Create confirmed event in owning calendar
    s_iso, e_iso = confirmed_slot["slot_start"], confirmed_slot["slot_end"]
    if account == "gmail":
        event_id = await gmail_create_confirmed(s_iso, e_iso, subject, client_email, client_name)
    else:
        event_id = await graph_create_confirmed(account, s_iso, e_iso, subject, client_email, client_name)

    # Delete holds for all OTHER tentative slots
    other_slots = [s for s in slots if s["id"] != slot_id and s["status"] == "tentative"]
    for s in other_slots:
        s["_account"] = account
    await asyncio.gather(*[_delete_slot_holds(s) for s in other_slots], return_exceptions=True)

    # Update DB: confirmed slot → confirmed, others → declined, proposal → confirmed
    await update_slot_status(slot_id, "confirmed")
    for s in other_slots:
        await update_slot_status(s["id"], "declined")
    await update_proposal_status(proposal_id, "confirmed")

    return JSONResponse({"ok": True, "event_id": event_id or None})


@app.post("/api/meeting/{proposal_id}/decline-slot/{slot_id}")
async def api_meeting_decline_slot(proposal_id: int, slot_id: int):
    """
    Decline a single meeting slot: delete its calendar holds and mark it declined.
    If all slots are now declined, mark the proposal declined too.
    """
    proposal = await get_meeting_proposal(proposal_id)
    if not proposal:
        return JSONResponse({"ok": False, "error": "Proposal not found"}, status_code=404)

    slots = await get_slots_for_proposal(proposal_id)
    target = next((s for s in slots if s["id"] == slot_id), None)
    if not target:
        return JSONResponse({"ok": False, "error": "Slot not found"}, status_code=404)

    target["_account"] = proposal["account"]
    await _delete_slot_holds(target)
    await update_slot_status(slot_id, "declined")

    # If no tentative slots remain, close the proposal
    remaining = [s for s in slots if s["id"] != slot_id and s["status"] == "tentative"]
    if not remaining:
        await update_proposal_status(proposal_id, "declined")

    return JSONResponse({"ok": True})


@app.post("/api/meeting/{proposal_id}/decline")
async def api_meeting_decline(proposal_id: int):
    """
    Decline all slots for a proposal: delete all holds and mark everything declined.
    """
    proposal = await get_meeting_proposal(proposal_id)
    if not proposal:
        return JSONResponse({"ok": False, "error": "Proposal not found"}, status_code=404)
    if proposal["status"] != "pending":
        return JSONResponse({"ok": False, "error": f"Proposal is already {proposal['status']}"}, status_code=409)

    slots = await get_slots_for_proposal(proposal_id)
    account = proposal["account"]
    for s in slots:
        s["_account"] = account
    tentative = [s for s in slots if s["status"] == "tentative"]
    if tentative:
        await asyncio.gather(*[_delete_slot_holds(s) for s in tentative], return_exceptions=True)
        for s in tentative:
            await update_slot_status(s["id"], "declined")

    await update_proposal_status(proposal_id, "declined")
    return JSONResponse({"ok": True})


def _find_free_slots(
    busy_windows: list,
    biz_start_str: str,
    biz_end_str: str,
    duration_min: int,
    buffer_min: int,
    from_date: str = None,
    num_slots: int = 3,
    max_days: int = 14,
) -> list[dict]:
    """
    Walk business hours across up to max_days weekdays and return up to num_slots
    free windows of duration_min minutes, with buffer_min gaps around busy blocks.
    All busy_windows must be UTC-aware (datetime, datetime) tuples.
    Returned slots are ISO strings in Australia/Brisbane time.
    """
    from zoneinfo import ZoneInfo
    from datetime import date as date_type
    BRISBANE = ZoneInfo("Australia/Brisbane")

    duration = timedelta(minutes=duration_min)
    buffer   = timedelta(minutes=buffer_min)

    bh_start = tuple(int(x) for x in biz_start_str.split(":"))
    bh_end   = tuple(int(x) for x in biz_end_str.split(":"))

    if from_date:
        try:
            start_date = datetime.strptime(from_date, "%Y-%m-%d").date()
            day_offset_start = 0
        except ValueError:
            start_date = datetime.now(BRISBANE).date()
            day_offset_start = 1
    else:
        start_date = datetime.now(BRISBANE).date()
        day_offset_start = 1

    slots = []

    for offset in range(day_offset_start, max_days + day_offset_start + 1):
        day = start_date + timedelta(days=offset)
        if day.weekday() >= 5:
            continue

        biz_start = datetime(day.year, day.month, day.day, bh_start[0], bh_start[1], tzinfo=BRISBANE)
        biz_end   = datetime(day.year, day.month, day.day, bh_end[0],   bh_end[1],   tzinfo=BRISBANE)

        t = biz_start
        while t + duration <= biz_end:
            slot_end = t + duration

            # Find the furthest end of any overlapping busy window (with buffer)
            max_conflict_end = None
            for (bs, be) in busy_windows:
                if t < be + buffer and slot_end > bs - buffer:
                    if max_conflict_end is None or be > max_conflict_end:
                        max_conflict_end = be

            if max_conflict_end is None:
                slots.append({
                    "start": t.isoformat(),
                    "end":   slot_end.isoformat(),
                })
                if len(slots) >= num_slots:
                    return slots
                t = slot_end + buffer
            else:
                # Jump past the conflict and align to nearest 15-min boundary
                jump = max_conflict_end.astimezone(BRISBANE) + buffer
                rem  = (jump.hour * 60 + jump.minute) % 15
                if rem:
                    jump = jump + timedelta(minutes=15 - rem)
                t = jump.replace(second=0, microsecond=0)

    return slots


@app.get("/api/calendar/free-slots")
async def api_calendar_free_slots(
    account: str = "financial",
    duration: int = None,
    from_date: str = None,
):
    """
    Find free meeting slots by querying all relevant calendars and removing busy windows.
    financial/gmail: checks M365 financial + Google Calendar tax.
    personal: checks all three calendars.
    Returns up to 3 slot objects with ISO start/end strings (Brisbane time).
    """
    from zoneinfo import ZoneInfo
    BRISBANE = ZoneInfo("Australia/Brisbane")

    biz_start  = await get_setting("meeting_hours_start")  or "09:00"
    biz_end    = await get_setting("meeting_hours_end")    or "17:00"
    buffer_min = int(await get_setting("meeting_buffer_minutes")   or "15")
    if duration is None:
        duration = int(await get_setting("meeting_default_duration") or "60")

    # Search window: from_date (or today) through 16 days ahead
    if from_date:
        try:
            anchor = datetime.strptime(from_date, "%Y-%m-%d").replace(tzinfo=BRISBANE)
        except ValueError:
            anchor = datetime.now(BRISBANE)
    else:
        anchor = datetime.now(BRISBANE)

    window_start = anchor.replace(hour=0, minute=0, second=0, microsecond=0)
    window_end   = window_start + timedelta(days=18)

    # Always fetch both business calendars; add personal if needed
    fetches = [
        graph_get_busy_windows("financial", window_start, window_end),
        gmail_get_busy_windows(window_start, window_end),
    ]
    if account == "personal":
        fetches.append(graph_get_busy_windows("personal", window_start, window_end))

    batches = await asyncio.gather(*fetches, return_exceptions=True)

    busy = []
    for batch in batches:
        if not isinstance(batch, Exception):
            busy.extend(batch)
    busy.sort(key=lambda x: x[0])

    slots = _find_free_slots(
        busy, biz_start, biz_end, duration, buffer_min,
        from_date=from_date, num_slots=3,
    )

    return JSONResponse({"slots": slots, "duration": duration})


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
        "meeting_default_duration": await _get("meeting_default_duration", "60"),
        "meeting_hours_start": await _get("meeting_hours_start", "09:00"),
        "meeting_hours_end": await _get("meeting_hours_end", "17:00"),
        "meeting_buffer_minutes": await _get("meeting_buffer_minutes", "15"),
        "meeting_no_response_hours": await _get("meeting_no_response_hours", "36"),
        "meeting_post_slot_buffer_minutes": await _get("meeting_post_slot_buffer_minutes", "30"),
    })


@app.post("/api/settings")
async def api_save_setting(body: SettingUpdate):
    allowed = {"prompt_financial", "prompt_gmail", "prompt_personal",
               "footer_financial", "footer_gmail", "footer_personal",
               "meeting_default_duration", "meeting_hours_start", "meeting_hours_end",
               "meeting_buffer_minutes", "meeting_no_response_hours", "meeting_post_slot_buffer_minutes"}
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
