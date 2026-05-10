import asyncio
import logging
from email.utils import parsedate_to_datetime, parseaddr
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from connectors.graph import get_emails as graph_get_emails
from connectors.gmail import get_emails as gmail_get_emails
from agent.classifier import classify_email
from agent.actions import execute_action
from db.database import log_action, get_email_by_id, update_email_status, ensure_inbox_state, mark_missing_as_archived, set_auth_error, clear_auth_error, prune_old_records, get_sender_rule, upsert_sender_rule

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def parse_gmail_date(date_str: str) -> str:
    try:
        dt = parsedate_to_datetime(date_str)
        return dt.isoformat()
    except Exception:
        return None


async def poll_financial():
    logger.info("Polling financial inbox (full)...")
    inbox_ids = set()
    try:
        emails = await graph_get_emails("financial")
        logger.info(f"[financial] Found {len(emails)} emails in inbox")

        for email in emails:
            graph_id = email.get("id", "unknown")
            stable_id = email.get("internetMessageId") or graph_id
            inbox_ids.add(stable_id)
            subject = email.get("subject", "(no subject)")
            sender = email.get("from", {}).get("emailAddress", {}).get("address", "unknown")
            body = email.get("fullBody") or email.get("bodyPreview", "")
            received_at = email.get("receivedDateTime")

            # Inbox is source of truth: if the email is here it must be pending.
            # For known emails, refresh graph_id and restore status if needed.
            # Only classify emails we have never seen before.
            existing = await get_email_by_id(stable_id)
            if existing:
                prev = existing.get("status")
                if prev not in (None, "pending"):
                    logger.info(f"[financial] Back in inbox (was {prev}): {subject[:50]}")
                await ensure_inbox_state(stable_id, graph_id)
                continue

            rule = await get_sender_rule(sender)
            if rule and rule['source'] == 'manual':
                result = {
                    "classification": rule['classification'],
                    "confidence": 1.0,
                    "reason": f"sender rule (manual)",
                }
            else:
                result = await classify_email(subject, sender, body[:1000])
                await upsert_sender_rule(sender, result.get("classification", ""))
            classification = result.get("classification")
            confidence = result.get("confidence", 0.0)

            action_label, status = await execute_action(
                account="financial",
                email_id=graph_id,  # Graph folder-scoped ID for API operations
                subject=subject,
                sender=sender,
                classification=classification,
                confidence=confidence
            )

            await log_action(
                account="financial",
                email_id=stable_id,  # stable RFC 2822 ID for DB lookups
                subject=subject,
                sender=sender,
                action=action_label,
                classification=classification,
                confidence=confidence,
                notes=result.get("reason"),
                body=body,
                received_at=received_at,
                graph_id=graph_id
            )

            if status != "pending":
                await update_email_status(stable_id, status)

            logger.info(f"[financial] {subject[:50]} → {classification} ({confidence:.2f}) → {action_label}")

        # Reconcile — anything pending in SQLite but not in inbox gets archived
        missing = await mark_missing_as_archived("financial", inbox_ids)
        if missing:
            logger.info(f"[financial] Reconciled {len(missing)} emails no longer in inbox")
        await clear_auth_error("financial")

    except Exception as e:
        logger.error(f"Error polling financial inbox: {e}")
        _AUTH_KEYWORDS = ("401", "No token", "OAuth", "invalid_grant", "invalid_client", "AADSTS", "Unauthorized", "unauthorized", "offset-naive", "offset-aware")
        if any(k in str(e) for k in _AUTH_KEYWORDS):
            await set_auth_error("financial", str(e))


async def poll_gmail():
    logger.info("Polling Gmail inbox (full)...")
    inbox_ids = set()
    try:
        emails = await gmail_get_emails()
        logger.info(f"[gmail] Found {len(emails)} emails in inbox")

        for email in emails:
            email_id = email.get("id", "unknown")
            inbox_ids.add(email_id)
            subject = email.get("subject", "(no subject)")
            raw_from = email.get("from", "unknown")
            _, addr = parseaddr(raw_from)
            sender = addr.lower() if addr else raw_from
            body = email.get("fullBody") or email.get("snippet", "")
            received_at = parse_gmail_date(email.get("date", ""))

            existing = await get_email_by_id(email_id)
            if existing:
                prev = existing.get("status")
                if prev not in (None, "pending"):
                    logger.info(f"[gmail] Back in inbox (was {prev}): {subject[:50]}")
                await ensure_inbox_state(email_id, email_id)
                continue

            rule = await get_sender_rule(sender)
            if rule and rule['source'] == 'manual':
                result = {
                    "classification": rule['classification'],
                    "confidence": 1.0,
                    "reason": f"sender rule (manual)",
                }
            else:
                result = await classify_email(subject, sender, body[:1000])
                await upsert_sender_rule(sender, result.get("classification", ""))
            classification = result.get("classification")
            confidence = result.get("confidence", 0.0)

            action_label, status = await execute_action(
                account="gmail",
                email_id=email_id,
                subject=subject,
                sender=sender,
                classification=classification,
                confidence=confidence
            )

            await log_action(
                account="gmail",
                email_id=email_id,
                subject=subject,
                sender=sender,
                action=action_label,
                classification=classification,
                confidence=confidence,
                notes=result.get("reason"),
                body=body,
                received_at=received_at
            )

            if status != "pending":
                await update_email_status(email_id, status)

            logger.info(f"[gmail] {subject[:50]} → {classification} ({confidence:.2f}) → {action_label}")

        # Reconcile — anything pending in SQLite but not in inbox gets archived
        missing = await mark_missing_as_archived("gmail", inbox_ids)
        if missing:
            logger.info(f"[gmail] Reconciled {len(missing)} emails no longer in inbox")
        await clear_auth_error("gmail")

    except Exception as e:
        logger.error(f"Error polling Gmail inbox: {e}")
        _AUTH_KEYWORDS = ("401", "No token", "OAuth", "invalid_grant", "invalid_client", "AADSTS", "Unauthorized", "unauthorized", "offset-naive", "offset-aware")
        if any(k in str(e) for k in _AUTH_KEYWORDS):
            await set_auth_error("gmail", str(e))


async def poll_personal():
    logger.info("Polling personal inbox (full)...")
    inbox_ids = set()
    try:
        emails = await graph_get_emails("personal")
        logger.info(f"[personal] Found {len(emails)} emails in inbox")

        for email in emails:
            graph_id = email.get("id", "unknown")
            stable_id = email.get("internetMessageId") or graph_id
            inbox_ids.add(stable_id)
            subject = email.get("subject", "(no subject)")
            sender = email.get("from", {}).get("emailAddress", {}).get("address", "unknown")
            body = email.get("fullBody") or email.get("bodyPreview", "")
            received_at = email.get("receivedDateTime")

            existing = await get_email_by_id(stable_id)
            if existing:
                prev = existing.get("status")
                if prev not in (None, "pending"):
                    logger.info(f"[personal] Back in inbox (was {prev}): {subject[:50]}")
                await ensure_inbox_state(stable_id, graph_id)
                continue

            rule = await get_sender_rule(sender)
            if rule and rule['source'] == 'manual':
                result = {
                    "classification": rule['classification'],
                    "confidence": 1.0,
                    "reason": f"sender rule (manual)",
                }
            else:
                result = await classify_email(subject, sender, body[:1000])
                await upsert_sender_rule(sender, result.get("classification", ""))
            classification = result.get("classification")
            confidence = result.get("confidence", 0.0)

            action_label, status = await execute_action(
                account="personal",
                email_id=graph_id,  # Graph folder-scoped ID for API operations
                subject=subject,
                sender=sender,
                classification=classification,
                confidence=confidence
            )

            await log_action(
                account="personal",
                email_id=stable_id,  # stable RFC 2822 ID for DB lookups
                subject=subject,
                sender=sender,
                action=action_label,
                classification=classification,
                confidence=confidence,
                notes=result.get("reason"),
                body=body,
                received_at=received_at,
                graph_id=graph_id
            )

            if status != "pending":
                await update_email_status(stable_id, status)

            logger.info(f"[personal] {subject[:50]} → {classification} ({confidence:.2f}) → {action_label}")

        missing = await mark_missing_as_archived("personal", inbox_ids)
        if missing:
            logger.info(f"[personal] Reconciled {len(missing)} emails no longer in inbox")
        await clear_auth_error("personal")

    except Exception as e:
        logger.error(f"Error polling personal inbox: {e}")
        _AUTH_KEYWORDS = ("401", "No token", "OAuth", "invalid_grant", "invalid_client", "AADSTS", "Unauthorized", "unauthorized", "offset-naive", "offset-aware")
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
