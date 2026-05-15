import logging

import httpx

logger = logging.getLogger(__name__)

SPAM_DELETE_THRESHOLD = 0.95

async def execute_action(account: str, email_id: str, subject: str, sender: str,
                         classification: str, confidence: float):
    """
    Decide and execute an autonomous action based on classification and confidence.
    Returns (action_label, status) tuple.
    """
    if classification == "spam":
        if confidence >= SPAM_DELETE_THRESHOLD:
            return await _hard_delete_email(account, email_id, subject)
        else:
            return await _move_to_folder(account, email_id, subject, "Junk Email")
    elif classification == "newsletter":
        return await _move_to_folder(account, email_id, subject, "Newsletters")
    elif classification == "notification":
        return await _move_to_folder(account, email_id, subject, "Notifications")

    return ("queued", "pending")


async def _hard_delete_email(account: str, email_id: str, subject: str):
    """Permanently delete — used only for high-confidence spam."""
    try:
        if account in ("financial", "personal"):
            from connectors.graph import hard_delete_email
            await hard_delete_email(account, email_id)
        elif account == "gmail":
            from connectors.gmail import hard_delete_email
            await hard_delete_email(email_id)
        logger.info(f"[{account}] HARD DELETED spam: {subject[:50]}")
        return ("deleted", "deleted")
    except httpx.HTTPStatusError as e:
        if e.response.status_code == 404:
            logger.info(f"[{account}] Spam already gone (404) — treating as deleted: {subject[:50]}")
            return ("deleted_externally", "deleted")
        logger.error(f"[{account}] Failed to hard delete {email_id}: {e}")
        return ("delete_failed", "pending")
    except Exception as e:
        logger.error(f"[{account}] Failed to hard delete {email_id}: {e}")
        return ("delete_failed", "pending")


async def _move_to_folder(account: str, email_id: str, subject: str, folder_name: str):
    try:
        if account in ("financial", "personal"):
            from connectors.graph import move_email
            await move_email(account, email_id, folder_name)
        elif account == "gmail":
            from connectors.gmail import move_email
            await move_email(email_id, folder_name)
        logger.info(f"[{account}] MOVED to {folder_name}: {subject[:50]}")
        return (f"moved_to_{folder_name.lower().replace(' ', '_')}", "archived")
    except httpx.HTTPStatusError as e:
        if e.response.status_code == 404:
            # Message no longer at this Graph ID — already moved or deleted externally
            logger.info(f"[{account}] Message {email_id} already moved (404) — treating as archived")
            return ("reconciled", "archived")
        logger.error(f"[{account}] Failed to move {email_id} to {folder_name}: {e}")
        return ("move_failed", "pending")
    except Exception as e:
        logger.error(f"[{account}] Failed to move {email_id} to {folder_name}: {e}")
        return ("move_failed", "pending")
