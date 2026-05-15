import os
import json
import logging
import tempfile
import httpx
import base64
import re
from datetime import datetime, timedelta, timezone
from email.utils import parseaddr, parsedate_to_datetime
from dotenv import load_dotenv
from email.mime.text import MIMEText
from connectors.utils import strip_html

logger = logging.getLogger(__name__)

load_dotenv("config/.env")

CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID")
CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET")

_APP_BASE_URL = os.getenv("APP_BASE_URL", "http://localhost:8000").rstrip("/")
REDIRECT_URI = f"{_APP_BASE_URL}/auth/callback/gmail"
AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
GMAIL_BASE = "https://gmail.googleapis.com/gmail/v1/users/me"
PEOPLE_BASE = "https://people.googleapis.com/v1"

SCOPES = " ".join([
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/contacts.readonly",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "openid",
])

TOKEN_FILE = "config/tokens_gmail.json"


def load_tokens() -> dict:
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r") as f:
            return json.load(f)
    return {}


def save_tokens(tokens: dict):
    dir_ = os.path.dirname(TOKEN_FILE)
    with tempfile.NamedTemporaryFile("w", dir=dir_, delete=False, suffix=".tmp") as f:
        json.dump(tokens, f, indent=2)
        tmp = f.name
    os.replace(tmp, TOKEN_FILE)


def get_auth_url() -> str:
    params = (
        f"?client_id={CLIENT_ID}"
        f"&redirect_uri={REDIRECT_URI}"
        f"&response_type=code"
        f"&scope={SCOPES.replace(' ', '%20')}"
        f"&access_type=offline"
        f"&prompt=consent"
    )
    return AUTH_URL + params


def _parse_expires_at(expires_str: str) -> datetime:
    dt = datetime.fromisoformat(expires_str)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


async def exchange_code_for_token(code: str) -> dict:
    async with httpx.AsyncClient() as client:
        response = await client.post(
            TOKEN_URL,
            data={
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "code": code,
                "redirect_uri": REDIRECT_URI,
                "grant_type": "authorization_code",
            },
        )
        response.raise_for_status()
        token_data = response.json()
        token_data["expires_at"] = (
            datetime.now(timezone.utc) + timedelta(seconds=token_data.get("expires_in", 3600))
        ).isoformat()
        tokens = load_tokens()
        tokens["gmail"] = token_data
        save_tokens(tokens)
        return token_data


async def refresh_token() -> dict:
    tokens = load_tokens()
    token_data = tokens.get("gmail")
    if not token_data:
        raise Exception("No Gmail token found — OAuth login required")
    async with httpx.AsyncClient() as client:
        response = await client.post(
            TOKEN_URL,
            data={
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "refresh_token": token_data["refresh_token"],
                "grant_type": "refresh_token",
            },
        )
        response.raise_for_status()
        new_token = response.json()
        new_token["expires_at"] = (
            datetime.now(timezone.utc) + timedelta(seconds=new_token.get("expires_in", 3600))
        ).isoformat()
        new_token["refresh_token"] = token_data["refresh_token"]
        tokens["gmail"] = new_token
        save_tokens(tokens)
        return new_token


async def get_valid_token() -> str:
    tokens = load_tokens()
    token_data = tokens.get("gmail")
    if not token_data:
        raise Exception("No Gmail token — OAuth login required")
    expires_at = _parse_expires_at(token_data["expires_at"])
    if datetime.now(timezone.utc) >= expires_at - timedelta(minutes=5):
        token_data = await refresh_token()
    return token_data["access_token"]


def extract_body_from_payload(payload: dict) -> str:
    mime_type = payload.get("mimeType", "")
    body_data = payload.get("body", {}).get("data", "")
    if mime_type == "text/plain" and body_data:
        return base64.urlsafe_b64decode(body_data + "==").decode("utf-8", errors="replace")
    if mime_type == "text/html" and body_data:
        html = base64.urlsafe_b64decode(body_data + "==").decode("utf-8", errors="replace")
        return strip_html(html)
    parts = payload.get("parts", [])
    for part in parts:
        if part.get("mimeType") == "text/plain":
            data = part.get("body", {}).get("data", "")
            if data:
                return base64.urlsafe_b64decode(data + "==").decode("utf-8", errors="replace")
    for part in parts:
        if part.get("mimeType") == "text/html":
            data = part.get("body", {}).get("data", "")
            if data:
                html = base64.urlsafe_b64decode(data + "==").decode("utf-8", errors="replace")
                return strip_html(html)
    for part in parts:
        result = extract_body_from_payload(part)
        if result:
            return result
    return ""


async def get_emails() -> list:
    """Fetch all messages currently in the inbox by following pagination."""
    token = await get_valid_token()
    all_message_ids = []

    async with httpx.AsyncClient(timeout=60.0) as client:
        # Step 1: get all message IDs in inbox (paginated)
        params = {"maxResults": 500, "labelIds": "INBOX"}
        next_page_token = None

        while True:
            if next_page_token:
                params["pageToken"] = next_page_token
            else:
                params.pop("pageToken", None)

            list_response = await client.get(
                f"{GMAIL_BASE}/messages",
                headers={"Authorization": f"Bearer {token}"},
                params=params,
            )
            list_response.raise_for_status()
            data = list_response.json()
            messages = data.get("messages", [])
            all_message_ids.extend(messages)
            next_page_token = data.get("nextPageToken")
            if not next_page_token:
                break

        # Step 2: fetch full details for each message
        emails = []
        for msg in all_message_ids:
            detail = await client.get(
                f"{GMAIL_BASE}/messages/{msg['id']}",
                headers={"Authorization": f"Bearer {token}"},
                params={"format": "full"},
            )
            detail.raise_for_status()
            data = detail.json()
            headers = {h["name"]: h["value"] for h in data.get("payload", {}).get("headers", [])}
            full_body = extract_body_from_payload(data.get("payload", {}))
            emails.append({
                "id": msg["id"],
                "threadId": data.get("threadId"),
                "messageId": headers.get("Message-ID", ""),
                "subject": headers.get("Subject", "(no subject)"),
                "from": headers.get("From", "unknown"),
                "date": headers.get("Date", ""),
                "snippet": data.get("snippet", ""),
                "fullBody": full_body,
            })

        return emails


async def get_sent_emails(days: int = 90) -> list:
    token = await get_valid_token()
    since_date = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y/%m/%d")
    all_message_ids = []

    async with httpx.AsyncClient(timeout=60.0) as client:
        params = {"maxResults": 500, "labelIds": "SENT", "q": f"after:{since_date}"}
        next_page_token = None

        while True:
            if next_page_token:
                params["pageToken"] = next_page_token
            else:
                params.pop("pageToken", None)
            list_response = await client.get(
                f"{GMAIL_BASE}/messages",
                headers={"Authorization": f"Bearer {token}"},
                params=params,
            )
            list_response.raise_for_status()
            data = list_response.json()
            all_message_ids.extend(data.get("messages", []))
            next_page_token = data.get("nextPageToken")
            if not next_page_token:
                break

        emails = []
        for msg in all_message_ids:
            detail = await client.get(
                f"{GMAIL_BASE}/messages/{msg['id']}",
                headers={"Authorization": f"Bearer {token}"},
                params={"format": "full"},
            )
            detail.raise_for_status()
            data = detail.json()
            headers = {h["name"]: h["value"] for h in data.get("payload", {}).get("headers", [])}
            full_body = extract_body_from_payload(data.get("payload", {}))
            emails.append({
                "id": msg["id"],
                "subject": headers.get("Subject", "(no subject)"),
                "body": full_body,
                "sent_at": headers.get("Date", ""),
            })

        return emails


async def get_email_history(address: str, limit: int = 8) -> list[dict]:
    """
    Search across all Gmail labels for recent messages to/from a specific address.
    Returns list of {subject, date, direction, snippet} dicts, newest first.
    Fails silently — returns [] on any error.
    """
    try:
        token = await get_valid_token()
        async with httpx.AsyncClient(timeout=20.0) as client:
            # Search across all mail (no labelIds = searches everything incl. archive)
            list_resp = await client.get(
                f"{GMAIL_BASE}/messages",
                headers={"Authorization": f"Bearer {token}"},
                params={"q": f"from:{address} OR to:{address}", "maxResults": limit},
            )
            if not list_resp.is_success:
                return []
            msg_ids = [m["id"] for m in list_resp.json().get("messages", [])]
            if not msg_ids:
                return []

            items = []
            for msg_id in msg_ids:
                detail = await client.get(
                    f"{GMAIL_BASE}/messages/{msg_id}",
                    headers={"Authorization": f"Bearer {token}"},
                    params={
                        "format":          "metadata",
                        "metadataHeaders": "Subject,From,Date",
                    },
                )
                if not detail.is_success:
                    continue
                data    = detail.json()
                headers = {h["name"]: h["value"] for h in data.get("payload", {}).get("headers", [])}
                from_hdr = headers.get("From", "").lower()
                direction = "←" if address.lower() in from_hdr else "→"
                raw_date  = headers.get("Date", "")
                try:
                    date = parsedate_to_datetime(raw_date).strftime("%Y-%m-%d")
                except Exception:
                    date = raw_date[:10]
                items.append({
                    "subject":   headers.get("Subject", "(no subject)"),
                    "date":      date,
                    "direction": direction,
                    "snippet":   data.get("snippet", "")[:200],
                })
            return items
    except Exception:
        return []


async def get_user_profile() -> dict:
    token = await get_valid_token()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            "https://www.googleapis.com/oauth2/v2/userinfo",
            headers={"Authorization": f"Bearer {token}"},
        )
        response.raise_for_status()
        return response.json()


async def delete_email(email_id: str):
    """Soft delete — moves to Trash. Recoverable from Gmail for 30 days."""
    token = await get_valid_token()
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GMAIL_BASE}/messages/{email_id}/modify",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"addLabelIds": ["TRASH"], "removeLabelIds": ["INBOX"]},
        )
        response.raise_for_status()


async def hard_delete_email(email_id: str):
    """Permanent delete — unrecoverable. Used only for autonomous spam deletion."""
    token = await get_valid_token()
    async with httpx.AsyncClient() as client:
        response = await client.delete(
            f"{GMAIL_BASE}/messages/{email_id}",
            headers={"Authorization": f"Bearer {token}"},
        )
        response.raise_for_status()


async def archive_email(email_id: str):
    """Remove INBOX label — moves to All Mail."""
    token = await get_valid_token()
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GMAIL_BASE}/messages/{email_id}/modify",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"removeLabelIds": ["INBOX"]},
        )
        response.raise_for_status()


async def unarchive_email(email_id: str):
    """Add INBOX label back."""
    token = await get_valid_token()
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GMAIL_BASE}/messages/{email_id}/modify",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"addLabelIds": ["INBOX"]},
        )
        response.raise_for_status()


async def mark_as_read(email_id: str):
    token = await get_valid_token()
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GMAIL_BASE}/messages/{email_id}/modify",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"removeLabelIds": ["UNREAD"]},
        )
        response.raise_for_status()


async def send_email(to: str, subject: str, body: str,
                     thread_id: str = None, in_reply_to: str = None,
                     cc: str = None) -> str:
    """Send an email and return the Gmail message ID of the sent message."""
    token = await get_valid_token()
    profile = await get_user_profile()
    from_addr = profile.get("email", "")

    mime = MIMEText(body, "plain", "utf-8")
    mime["From"] = from_addr
    mime["To"] = to
    mime["Subject"] = subject
    if cc:
        mime["Cc"] = cc
    if in_reply_to:
        mime["In-Reply-To"] = in_reply_to
        mime["References"] = in_reply_to

    raw = base64.urlsafe_b64encode(mime.as_bytes()).decode("utf-8")
    payload: dict = {"raw": raw}
    if thread_id:
        payload["threadId"] = thread_id

    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GMAIL_BASE}/messages/send",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json=payload,
        )
        if not response.is_success:
            raise Exception(f"Gmail send failed {response.status_code}: {response.text}")
        msg_id = response.json().get("id", "")
        logger.info(f"[gmail] sent message id={msg_id} thread={thread_id}")
        return msg_id


async def get_or_create_label(label_name: str) -> str:
    token = await get_valid_token()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GMAIL_BASE}/labels",
            headers={"Authorization": f"Bearer {token}"},
        )
        response.raise_for_status()
        labels = response.json().get("labels", [])
        for label in labels:
            if label["name"].lower() == label_name.lower():
                return label["id"]
        response = await client.post(
            f"{GMAIL_BASE}/labels",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"name": label_name},
        )
        response.raise_for_status()
        return response.json()["id"]


CALENDAR_BASE = "https://www.googleapis.com/calendar/v3"


def _extract_ics_from_payload(payload: dict) -> str:
    if payload.get("mimeType") == "text/calendar":
        data = payload.get("body", {}).get("data", "")
        if data:
            return base64.urlsafe_b64decode(data + "==").decode("utf-8", errors="replace")
    for part in payload.get("parts", []):
        result = _extract_ics_from_payload(part)
        if result:
            return result
    return ""


def _extract_uid_from_ics(ics: str) -> str:
    for line in ics.splitlines():
        if line.upper().startswith("UID:"):
            return line[4:].strip()
    return ""


async def _respond_to_calendar_event(email_id: str, response_status: str):
    """Find a calendar event via the invite ICS and update attendee response."""
    token = await get_valid_token()

    async with httpx.AsyncClient(timeout=30.0) as client:
        # Re-fetch message to get ICS payload
        detail = await client.get(
            f"{GMAIL_BASE}/messages/{email_id}",
            headers={"Authorization": f"Bearer {token}"},
            params={"format": "full"},
        )
        detail.raise_for_status()
        ics = _extract_ics_from_payload(detail.json().get("payload", {}))
        if not ics:
            raise Exception("No ICS data found in this email — cannot respond via Calendar API")

        uid = _extract_uid_from_ics(ics)
        if not uid:
            raise Exception("Could not extract event UID from ICS")

        # Find event in primary calendar by UID
        search = await client.get(
            f"{CALENDAR_BASE}/calendars/primary/events",
            headers={"Authorization": f"Bearer {token}"},
            params={"iCalUID": uid, "singleEvents": "true"},
        )
        search.raise_for_status()
        items = search.json().get("items", [])
        if not items:
            raise Exception(f"Event UID {uid} not found in Google Calendar")

        event = items[0]
        event_id = event["id"]

        # Get authenticated user's email to find their attendee entry
        profile = await get_user_profile()
        user_email = profile.get("email", "").lower()

        attendees = event.get("attendees", [])
        matched = False
        for attendee in attendees:
            if attendee.get("email", "").lower() == user_email:
                attendee["responseStatus"] = response_status
                matched = True
                break
        if not matched:
            attendees.append({"email": user_email, "responseStatus": response_status})

        patch = await client.patch(
            f"{CALENDAR_BASE}/calendars/primary/events/{event_id}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"attendees": attendees},
            params={"sendUpdates": "all"},
        )
        patch.raise_for_status()


async def accept_calendar_event(email_id: str):
    await _respond_to_calendar_event(email_id, "accepted")


async def decline_calendar_event(email_id: str):
    await _respond_to_calendar_event(email_id, "declined")


async def get_busy_windows(start_dt: datetime, end_dt: datetime) -> list[tuple]:
    """
    Return (start, end) UTC datetime tuples for all non-free, non-cancelled,
    non-all-day events in the given window from Google Calendar primary.
    Fails silently — returns [] on any error.
    """
    try:
        token = await get_valid_token()
        params = {
            "timeMin":      start_dt.astimezone(timezone.utc).isoformat(),
            "timeMax":      end_dt.astimezone(timezone.utc).isoformat(),
            "singleEvents": "true",
            "maxResults":   250,
            "fields":       "nextPageToken,items(start,end,status,transparency)",
        }
        busy = []
        async with httpx.AsyncClient(timeout=20.0) as client:
            while True:
                resp = await client.get(
                    f"{CALENDAR_BASE}/calendars/primary/events",
                    headers={"Authorization": f"Bearer {token}"},
                    params=params,
                )
                resp.raise_for_status()
                data = resp.json()
                for ev in data.get("items", []):
                    if ev.get("status") == "cancelled":
                        continue
                    if ev.get("transparency") == "transparent":
                        continue
                    s = ev.get("start", {})
                    e = ev.get("end", {})
                    if "dateTime" not in s:
                        continue  # all-day event
                    s_dt = datetime.fromisoformat(s["dateTime"]).astimezone(timezone.utc)
                    e_dt = datetime.fromisoformat(e["dateTime"]).astimezone(timezone.utc)
                    busy.append((s_dt, e_dt))
                next_page = data.get("nextPageToken")
                if not next_page:
                    break
                params["pageToken"] = next_page
        return busy
    except Exception as e:
        logger.debug(f"gmail get_busy_windows failed: {e}")
        return []


async def create_calendar_hold(start_iso: str, end_iso: str, title: str = "Hold") -> str:
    """
    Create a tentative calendar hold on Google Calendar primary.
    Returns the Google Calendar event id, or "" on failure.
    """
    try:
        token = await get_valid_token()
        s_dt = datetime.fromisoformat(start_iso).astimezone(timezone.utc)
        e_dt = datetime.fromisoformat(end_iso).astimezone(timezone.utc)
        event = {
            "summary": title,
            "status": "tentative",
            "transparency": "opaque",
            "start": {"dateTime": s_dt.isoformat(), "timeZone": "UTC"},
            "end":   {"dateTime": e_dt.isoformat(), "timeZone": "UTC"},
        }
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.post(
                f"{CALENDAR_BASE}/calendars/primary/events",
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json=event,
            )
            resp.raise_for_status()
            return resp.json().get("id", "")
    except Exception as e:
        logger.warning(f"gmail create_calendar_hold failed: {e}")
        return ""


async def create_confirmed_event(start_iso: str, end_iso: str,
                                  title: str, client_email: str, client_name: str = "") -> str:
    """
    Create a confirmed Google Calendar event with the client as attendee.
    sendUpdates='all' causes Google to email the invite. Returns event id, or "" on failure.
    """
    try:
        token = await get_valid_token()
        s_dt  = datetime.fromisoformat(start_iso).astimezone(timezone.utc)
        e_dt  = datetime.fromisoformat(end_iso).astimezone(timezone.utc)
        event = {
            "summary": title,
            "status": "confirmed",
            "transparency": "opaque",
            "start": {"dateTime": s_dt.isoformat(), "timeZone": "UTC"},
            "end":   {"dateTime": e_dt.isoformat(), "timeZone": "UTC"},
            "attendees": [{"email": client_email, "displayName": client_name or client_email}],
        }
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.post(
                f"{CALENDAR_BASE}/calendars/primary/events",
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json=event,
                params={"sendUpdates": "all"},
            )
            resp.raise_for_status()
            return resp.json().get("id", "")
    except Exception as e:
        logger.warning(f"gmail create_confirmed_event failed: {e}")
        return ""


async def delete_calendar_event(event_id: str) -> bool:
    """Delete a Google Calendar event by id. Returns True on success."""
    try:
        token = await get_valid_token()
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.delete(
                f"{CALENDAR_BASE}/calendars/primary/events/{event_id}",
                headers={"Authorization": f"Bearer {token}"},
            )
            return resp.status_code in (200, 204)
    except Exception as e:
        logger.warning(f"gmail delete_calendar_event failed: {e}")
        return False


async def move_email(email_id: str, folder_name: str):
    label_id = await get_or_create_label(folder_name)
    token = await get_valid_token()
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GMAIL_BASE}/messages/{email_id}/modify",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"addLabelIds": [label_id], "removeLabelIds": ["INBOX"]},
        )
        response.raise_for_status()


_GMAIL_SYSTEM_LABELS = {
    "INBOX", "SENT", "TRASH", "DRAFT", "SPAM", "STARRED", "IMPORTANT", "UNREAD",
    "CATEGORY_SOCIAL", "CATEGORY_UPDATES", "CATEGORY_FORUMS",
    "CATEGORY_PROMOTIONS", "CATEGORY_PERSONAL",
}


async def list_labels() -> list:
    """Return user-created labels only (excludes Gmail system labels)."""
    token = await get_valid_token()
    async with httpx.AsyncClient(timeout=15.0) as client:
        resp = await client.get(
            f"{GMAIL_BASE}/labels",
            headers={"Authorization": f"Bearer {token}"},
        )
        resp.raise_for_status()
        labels = []
        for label in resp.json().get("labels", []):
            if label["id"] not in _GMAIL_SYSTEM_LABELS and not label["id"].startswith("CATEGORY_"):
                labels.append({"id": label["id"], "name": label["name"]})
        return sorted(labels, key=lambda l: l["name"].lower())


async def file_to_label(email_id: str, label_id: str):
    """Apply a specific label by ID and remove from inbox — used for manual filing."""
    token = await get_valid_token()
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GMAIL_BASE}/messages/{email_id}/modify",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"addLabelIds": [label_id], "removeLabelIds": ["INBOX"]},
        )
        response.raise_for_status()


async def export_mime(email_id: str) -> bytes:
    """Export a message as raw MIME bytes — used for cross-account filing."""
    token = await get_valid_token()
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.get(
            f"{GMAIL_BASE}/messages/{email_id}",
            headers={"Authorization": f"Bearer {token}"},
            params={"format": "raw"},
        )
        resp.raise_for_status()
        raw = resp.json().get("raw", "")
        return base64.urlsafe_b64decode(raw + "==")


async def import_mime(label_id: str, mime_bytes: bytes) -> str:
    """Import raw MIME bytes as a new Gmail message and apply a target label.
    Used for cross-account filing — creates a copy in the target label."""
    token = await get_valid_token()
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.post(
            "https://www.googleapis.com/upload/gmail/v1/users/me/messages",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "message/rfc822",
            },
            params={"uploadType": "media"},
            content=mime_bytes,
        )
        resp.raise_for_status()
        msg_id = resp.json().get("id", "")
        if msg_id and label_id:
            await client.post(
                f"{GMAIL_BASE}/messages/{msg_id}/modify",
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json={"addLabelIds": [label_id], "removeLabelIds": ["INBOX"]},
            )
        return msg_id


async def search_contacts(q: str, limit: int = 10) -> list[dict]:
    """
    Search Gmail contacts via the People API.
    Returns [{ name, email, source: 'contacts' }].
    Fails silently — returns [] if scope not granted or any error.
    """
    if not q:
        return []
    try:
        token = await get_valid_token()
        async with httpx.AsyncClient(timeout=10.0) as client:
            resp = await client.get(
                f"{PEOPLE_BASE}/people:searchContacts",
                headers={"Authorization": f"Bearer {token}"},
                params={
                    "query": q,
                    "readMask": "names,emailAddresses,organizations",
                    "pageSize": limit,
                },
            )
        if not resp.is_success:
            return []
        results = []
        for item in resp.json().get("results", []):
            person = item.get("person", {})
            names = person.get("names", [])
            emails = person.get("emailAddresses", [])
            orgs = person.get("organizations", [])
            name = next((n.get("displayName", "") for n in names if n.get("displayName")), "")
            email = next((e.get("value", "") for e in emails if e.get("value")), "")
            company = next((o.get("name", "") for o in orgs if o.get("name")), "")
            if not email:
                continue
            entry = {"name": name or email, "email": email.strip(), "source": "contacts"}
            if company:
                entry["company"] = company
            results.append(entry)
        return results
    except Exception:
        return []
