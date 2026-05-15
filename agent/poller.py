import asyncio
import logging
from email.utils import parsedate_to_datetime, parseaddr
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from connectors.graph import get_emails as graph_get_emails, delete_calendar_event as graph_delete_event
from connectors.gmail import get_emails as gmail_get_emails, delete_calendar_event as gmail_delete_event
from agent.classifier import classify_email
from agent.actions import execute_action
from db.database import log_action, get_email_by_id, update_email_status, ensure_inbox_state, mark_missing_as_archived, set_auth_error, clear_auth_error, prune_old_records, get_sender_rule, get_open_proposals_for_client, get_expired_tentative_slots, update_slot_status, expire_completed_proposals

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

_AUTH_KEYWORDS = (
    "401", "No token", "OAuth", "invalid_grant", "invalid_client",
    "AADSTS", "Unauthorized", "unauthorized", "offset-naive", "offset-aware",
)


def _parse_gmail_date(date_str: str) -> str:
    try:
        dt = parsedate_to_datetime(date_str)
        return dt.isoformat()
    except Exception:
        return None


def _normalize_graph_email(email: dict) -> dict:
    graph_id = email.get("id", "unknown")
    stable_id = email.get("internetMessageId") or graph_id
    sender = email.get("from", {}).get("emailAddress", {}).get("address", "unknown")
    body = email.get("fullBody") or email.get("bodyPreview", "")
    return {
        "stable_id": stable_id,
        "op_id": graph_id,
        "graph_id": graph_id,
        "subject": email.get("subject", "(no subject)"),
        "sender": sender,
        "body": body,
        "received_at": email.get("receivedDateTime"),
        "thread_id": None,
        "orig_message_id": None,
    }


def _normalize_gmail_email(email: dict) -> dict:
    email_id = email.get("id", "unknown")
    raw_from = email.get("from", "unknown")
    _, addr = parseaddr(raw_from)
    sender = addr.lower() if addr else raw_from
    body = email.get("fullBody") or email.get("snippet", "")
    return {
        "stable_id": email_id,
        "op_id": email_id,
        "graph_id": None,
        "subject": email.get("subject", "(no subject)"),
        "sender": sender,
        "body": body,
        "received_at": _parse_gmail_date(email.get("date", "")),
        "thread_id": email.get("threadId"),
        "orig_message_id": email.get("messageId"),
    }


async def _process_emails(account: str, emails: list, normalize_fn) -> set:
    inbox_ids = set()
    for raw_email in emails:
        n = normalize_fn(raw_email)
        inbox_ids.add(n["stable_id"])

        existing = await get_email_by_id(n["stable_id"])
        if existing:
            prev = existing.get("status")
            if prev not in (None, "pending"):
                logger.info(f"[{account}] Back in inbox (was {prev}): {n['subject'][:50]}")
            await ensure_inbox_state(n["stable_id"], n["op_id"])
            continue

        # Meeting response detection takes priority over all other classification
        open_proposals = await get_open_proposals_for_client(n["sender"])
        if open_proposals:
            count = len(open_proposals)
            result = {
                "classification": "meeting_response",
                "confidence": 0.95,
                "reason": f"Meeting response detected — {count} open proposal(s) from this sender",
            }
        else:
            rule = await get_sender_rule(n["sender"])
            if rule and rule["source"] == "manual" and rule["count"] >= 2:
                result = {
                    "classification": rule["classification"],
                    "confidence": 1.0,
                    "reason": "sender rule (manual, confirmed)",
                }
            else:
                result = await classify_email(n["subject"], n["sender"], n["body"][:1000])

        classification = result.get("classification")
        confidence = result.get("confidence", 0.0)

        action_label, status = await execute_action(
            account=account,
            email_id=n["op_id"],
            subject=n["subject"],
            sender=n["sender"],
            classification=classification,
            confidence=confidence,
        )

        await log_action(
            account=account,
            email_id=n["stable_id"],
            subject=n["subject"],
            sender=n["sender"],
            action=action_label,
            classification=classification,
            confidence=confidence,
            notes=result.get("reason"),
            body=n["body"],
            received_at=n["received_at"],
            graph_id=n["graph_id"],
            thread_id=n["thread_id"],
            orig_message_id=n["orig_message_id"],
        )

        if status != "pending":
            await update_email_status(n["stable_id"], status)

        logger.info(f"[{account}] {n['subject'][:50]} → {classification} ({confidence:.2f}) → {action_label}")

    return inbox_ids


async def _poll_account(account: str, fetch_fn, normalize_fn):
    logger.info(f"Polling {account} inbox (full)...")
    try:
        emails = await fetch_fn() if account == "gmail" else await fetch_fn(account)
        logger.info(f"[{account}] Found {len(emails)} emails in inbox")
        inbox_ids = await _process_emails(account, emails, normalize_fn)
        missing = await mark_missing_as_archived(account, inbox_ids)
        if missing:
            logger.info(f"[{account}] Reconciled {len(missing)} emails no longer in inbox")
        await clear_auth_error(account)
    except Exception as e:
        logger.error(f"Error polling {account} inbox: {e}")
        if any(k in str(e) for k in _AUTH_KEYWORDS):
            await set_auth_error(account, str(e))


async def poll_financial():
    await _poll_account("financial", graph_get_emails, _normalize_graph_email)


async def poll_gmail():
    await _poll_account("gmail", gmail_get_emails, _normalize_gmail_email)


async def poll_personal():
    await _poll_account("personal", graph_get_emails, _normalize_graph_email)


async def poll_all():
    await asyncio.gather(poll_financial(), poll_gmail(), poll_personal())
    pruned = await prune_old_records(days=90)
    if pruned:
        logger.info(f"Pruned {pruned} records older than 90 days")


async def release_expired_slots():
    """
    Delete calendar holds for any tentative meeting slots whose auto_release_at has passed,
    mark them declined, then expire proposals that have no remaining tentative slots.
    """
    try:
        expired = await get_expired_tentative_slots()
        if not expired:
            return

        logger.info(f"Auto-releasing {len(expired)} expired tentative meeting slot(s)")

        for slot in expired:
            account = slot.get("account", "financial")
            own_id  = slot.get("owning_calendar_event_id")
            fin_id  = slot.get("mirror_event_id_financial")
            tax_id  = slot.get("mirror_event_id_google_tax")

            deletes = []
            if own_id:
                if account == "gmail":
                    deletes.append(gmail_delete_event(own_id))
                else:
                    deletes.append(graph_delete_event(account, own_id))
            if fin_id:
                deletes.append(graph_delete_event("financial", fin_id))
            if tax_id:
                deletes.append(gmail_delete_event(tax_id))

            if deletes:
                results = await asyncio.gather(*deletes, return_exceptions=True)
                errors = [r for r in results if isinstance(r, Exception)]
                if errors:
                    logger.warning(f"Auto-release: {len(errors)} calendar deletion(s) failed for slot {slot['id']}")

            await update_slot_status(slot["id"], "declined")
            logger.info(f"Auto-released slot {slot['id']} (proposal {slot['proposal_id']})")

        await expire_completed_proposals()

    except Exception as e:
        logger.error(f"release_expired_slots failed: {e}")


def start_scheduler():
    scheduler = AsyncIOScheduler()
    scheduler.add_job(poll_all, "interval", minutes=5, id="poll_all")
    scheduler.add_job(release_expired_slots, "interval", minutes=15, id="release_expired_slots")
    scheduler.start()
    logger.info("Poller started — polling every 5 minutes, slot auto-release every 15 minutes")
    return scheduler
