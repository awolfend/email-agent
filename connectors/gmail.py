import os
import json
import httpx
import base64
import re
from datetime import datetime, timedelta, timezone
from email.utils import parseaddr
from dotenv import load_dotenv
from email.mime.text import MIMEText

load_dotenv("config/.env")

CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID")
CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET")

REDIRECT_URI = "http://localhost:8000/auth/callback/gmail"
AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
GMAIL_BASE = "https://gmail.googleapis.com/gmail/v1/users/me"

SCOPES = " ".join([
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/calendar",
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
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f, indent=2)


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


def strip_html(html: str) -> str:
    text = re.sub(r'<style[^>]*>.*?</style>', '', html, flags=re.DOTALL)
    text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL)
    text = re.sub(r'<br\s*/?>', '\n', text)
    text = re.sub(r'<p[^>]*>', '\n', text)
    text = re.sub(r'</p>', '', text)
    text = re.sub(r'<[^>]+>', '', text)
    text = re.sub(r'&nbsp;', ' ', text)
    text = re.sub(r'&amp;', '&', text)
    text = re.sub(r'&lt;', '<', text)
    text = re.sub(r'&gt;', '>', text)
    text = re.sub(r'&quot;', '"', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


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


async def get_emails(count: int = None) -> list:
    """
    Fetch ALL messages currently in the inbox by following pagination.
    The count parameter is ignored — full inbox is always returned.
    """
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


async def send_email(to: str, subject: str, body: str):
    token = await get_valid_token()
    mime = MIMEText(body, "plain")
    mime["to"] = to
    mime["subject"] = subject
    raw = base64.urlsafe_b64encode(mime.as_bytes()).decode("utf-8")
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{GMAIL_BASE}/messages/send",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"raw": raw},
        )
        response.raise_for_status()


async def get_or_create_label(label_name: str) -> str:
    token = await get_valid_token()
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{GMAIL_BASE}/labels",
            headers={"Authorization": f"Bearer {token}"},
        )
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
