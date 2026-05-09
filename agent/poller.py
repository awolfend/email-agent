import asyncio
import logging
from email.utils import parsedate_to_datetime
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from connectors.graph import get_emails as graph_get_emails
from connectors.gmail import get_emails as gmail_get_emails
from agent.classifier import classify_email
from agent.actions import execute_action
from db.database import log_action, get_email_by_id, update_email_status, mark_missing_as_archived

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
            email_id = email.get("id", "unknown")
            inbox_ids.add(email_id)
            subject = email.get("subject", "(no subject)")
            sender = email.get("from", {}).get("emailAddress", {}).get("address", "unknown")
            body = email.get("fullBody") or email.get("bodyPreview", "")
            received_at = email.get("receivedDateTime")

            # Skip already actioned emails
            existing = await get_email_by_id(email_id)
            if existing and existing.get("status") not in (None, "pending"):
                logger.info(f"[financial] Skipping actioned: {subject[:50]} ({existing['status']})")
                continue

            result = await classify_email(subject, sender, body[:1000])
            classification = result.get("classification")
            confidence = result.get("confidence", 0.0)

            action_label, status = await execute_action(
                account="financial",
                email_id=email_id,
                subject=subject,
                sender=sender,
                classification=classification,
                confidence=confidence
            )

            await log_action(
                account="financial",
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

            logger.info(f"[financial] {subject[:50]} → {classification} ({confidence:.2f}) → {action_label}")

        # Reconcile — anything pending in SQLite but not in inbox gets archived
        missing = await mark_missing_as_archived("financial", inbox_ids)
        if missing:
            logger.info(f"[financial] Reconciled {len(missing)} emails no longer in inbox")

    except Exception as e:
        logger.error(f"Error polling financial inbox: {e}")


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
            sender = email.get("from", "unknown")
            body = email.get("fullBody") or email.get("snippet", "")
            received_at = parse_gmail_date(email.get("date", ""))

            # Skip already actioned emails
            existing = await get_email_by_id(email_id)
            if existing and existing.get("status") not in (None, "pending"):
                logger.info(f"[gmail] Skipping actioned: {subject[:50]} ({existing['status']})")
                continue

            result = await classify_email(subject, sender, body[:1000])
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

    except Exception as e:
        logger.error(f"Error polling Gmail inbox: {e}")


async def poll_personal():
    logger.info("Polling personal inbox (full)...")
    inbox_ids = set()
    try:
        emails = await graph_get_emails("personal")
        logger.info(f"[personal] Found {len(emails)} emails in inbox")

        for email in emails:
            email_id = email.get("id", "unknown")
            inbox_ids.add(email_id)
            subject = email.get("subject", "(no subject)")
            sender = email.get("from", {}).get("emailAddress", {}).get("address", "unknown")
            body = email.get("fullBody") or email.get("bodyPreview", "")
            received_at = email.get("receivedDateTime")

            existing = await get_email_by_id(email_id)
            if existing and existing.get("status") not in (None, "pending"):
                logger.info(f"[personal] Skipping actioned: {subject[:50]} ({existing['status']})")
                continue

            result = await classify_email(subject, sender, body[:1000])
            classification = result.get("classification")
            confidence = result.get("confidence", 0.0)

            action_label, status = await execute_action(
                account="personal",
                email_id=email_id,
                subject=subject,
                sender=sender,
                classification=classification,
                confidence=confidence
            )

            await log_action(
                account="personal",
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

            logger.info(f"[personal] {subject[:50]} → {classification} ({confidence:.2f}) → {action_label}")

        missing = await mark_missing_as_archived("personal", inbox_ids)
        if missing:
            logger.info(f"[personal] Reconciled {len(missing)} emails no longer in inbox")

    except Exception as e:
        logger.error(f"Error polling personal inbox: {e}")


async def poll_all():
    await asyncio.gather(poll_financial(), poll_gmail(), poll_personal())


def start_scheduler():
    scheduler = AsyncIOScheduler()
    scheduler.add_job(poll_all, "interval", minutes=5, id="poll_all")
    scheduler.start()
    logger.info("Poller started — polling every 5 minutes")
    return scheduler
