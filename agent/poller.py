import asyncio
import logging
from email.utils import parsedate_to_datetime, parseaddr
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from connectors.graph import get_emails as graph_get_emails
from connectors.gmail import get_emails as gmail_get_emails
from agent.classifier import classify_email
from agent.actions import execute_action
from db.database import log_action, get_email_by_id, update_email_status, ensure_inbox_state, mark_missing_as_archived, set_auth_error, clear_auth_error, prune_old_records, get_sender_rule, get_open_proposals_for_client

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


async def poll_financial():
    logger.info("Polling financial inbox (full)...")
    try:
        emails = await graph_get_emails("financial")
        logger.info(f"[financial] Found {len(emails)} emails in inbox")
        inbox_ids = await _process_emails("financial", emails, _normalize_graph_email)
        missing = await mark_missing_as_archived("financial", inbox_ids)
        if missing:
            logger.info(f"[financial] Reconciled {len(missing)} emails no longer in inbox")
        await clear_auth_error("financial")
    except Exception as e:
        logger.error(f"Error polling financial inbox: {e}")
        if any(k in str(e) for k in _AUTH_KEYWORDS):
            await set_auth_error("financial", str(e))


async def poll_gmail():
    logger.info("Polling Gmail inbox (full)...")
    try:
        emails = await gmail_get_emails()
        logger.info(f"[gmail] Found {len(emails)} emails in inbox")
        inbox_ids = await _process_emails("gmail", emails, _normalize_gmail_email)
        missing = await mark_missing_as_archived("gmail", inbox_ids)
        if missing:
            logger.info(f"[gmail] Reconciled {len(missing)} emails no longer in inbox")
        await clear_auth_error("gmail")
    except Exception as e:
        logger.error(f"Error polling Gmail inbox: {e}")
        if any(k in str(e) for k in _AUTH_KEYWORDS):
            await set_auth_error("gmail", str(e))


async def poll_personal():
    logger.info("Polling personal inbox (full)...")
    try:
        emails = await graph_get_emails("personal")
        logger.info(f"[personal] Found {len(emails)} emails in inbox")
        inbox_ids = await _process_emails("personal", emails, _normalize_graph_email)
        missing = await mark_missing_as_archived("personal", inbox_ids)
        if missing:
            logger.info(f"[personal] Reconciled {len(missing)} emails no longer in inbox")
        await clear_auth_error("personal")
    except Exception as e:
        logger.error(f"Error polling personal inbox: {e}")
        if any(k in str(e) for k in _AUTH_KEYWORDS):
            await set_auth_error("personal", str(e))


async def poll_all():
    await asyncio.gather(poll_financial(), poll_gmail(), poll_personal())
    pruned = await prune_old_records(days=90)
    if pruned:
        logger.info(f"Pruned {pruned} records older than 90 days")


def start_scheduler():
    scheduler = AsyncIOScheduler()
    scheduler.add_job(poll_all, "interval", minutes=5, id="poll_all")
    scheduler.start()
    logger.info("Poller started — polling every 5 minutes")
    return scheduler
