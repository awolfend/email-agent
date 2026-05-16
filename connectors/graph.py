import asyncio
import os
import logging
import httpx
import json
import re
import tempfile
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv
from connectors.utils import strip_html

logger = logging.getLogger(__name__)

load_dotenv("config/.env")

CLIENT_ID = os.getenv("AZURE_CLIENT_ID_FINANCIAL")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET_FINANCIAL")
CLIENT_ID_PERSONAL = os.getenv("AZURE_CLIENT_ID_PERSONAL")
CLIENT_SECRET_PERSONAL = os.getenv("AZURE_CLIENT_SECRET_PERSONAL")
TENANT_ID = os.getenv("AZURE_TENANT_ID")
PERSONAL_EMAIL = os.getenv("PERSONAL_EMAIL")

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = "Mail.ReadWrite Mail.Send Calendars.ReadWrite User.Read offline_access"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

TOKEN_FILE = "config/tokens_graph.json"


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


def _mailbox_base(account: str) -> str:
    """Return the Graph API base path for a given account's mailbox."""
    if account == "personal":
        return f"{GRAPH_BASE}/users/{PERSONAL_EMAIL}"
    return f"{GRAPH_BASE}/me"


_APP_BASE_URL = os.getenv("APP_BASE_URL", "http://localhost:8000").rstrip("/")


def get_auth_url(account: str) -> str:
    redirect_uri = f"{_APP_BASE_URL}/auth/callback/{account}"
    return (
        f"{AUTHORITY}/oauth2/v2.0/authorize"
        f"?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={redirect_uri}"
        f"&response_mode=query"
        f"&scope={SCOPES.replace(' ', '%20')}"
        f"&state={account}"
        f"&prompt=select_account"
    )


async def exchange_code_for_token(code: str, account: str) -> dict:
    redirect_uri = f"{_APP_BASE_URL}/auth/callback/{account}"
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{AUTHORITY}/oauth2/v2.0/token",
            data={
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "code": code,
                "redirect_uri": redirect_uri,
                "grant_type": "authorization_code",
                "scope": SCOPES,
            },
        )
        response.raise_for_status()
        token_data = response.json()
        token_data["account"] = account
        token_data["expires_at"] = (
            datetime.now(timezone.utc) + timedelta(seconds=token_data.get("expires_in", 3600))
        ).isoformat()
        tokens = load_tokens()
        tokens[account] = token_data
        save_tokens(tokens)
        return token_data


async def refresh_token(account: str) -> dict:
    if account != "financial":
        raise Exception(f"refresh_token called for non-delegated account: {account}")
    tokens = load_tokens()
    token_data = tokens.get(account)
    if not token_data:
        raise Exception(f"No token found for account: {account}")
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{AUTHORITY}/oauth2/v2.0/token",
            data={
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "refresh_token": token_data["refresh_token"],
                "grant_type": "refresh_token",
                "scope": SCOPES,
            },
        )
        response.raise_for_status()
        new_token = response.json()
        new_token["account"] = account
        new_token["expires_at"] = (
            datetime.now(timezone.utc) + timedelta(seconds=new_token.get("expires_in", 3600))
        ).isoformat()
        tokens[account] = new_token
        save_tokens(tokens)
        return new_token


def _parse_expires_at(expires_str: str) -> datetime:
    """Parse an ISO expires_at string, always returning a timezone-aware UTC datetime."""
    dt = datetime.fromisoformat(expires_str)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


async def get_app_token() -> str:
    """Client credentials flow for the personal account — no user sign-in required."""
    tokens = load_tokens()
    cached = tokens.get("personal_app", {})
    if cached.get("access_token") and cached.get("expires_at"):
        if datetime.now(timezone.utc) < _parse_expires_at(cached["expires_at"]) - timedelta(minutes=5):
            return cached["access_token"]
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{AUTHORITY}/oauth2/v2.0/token",
            data={
                "client_id": CLIENT_ID_PERSONAL,
                "client_secret": CLIENT_SECRET_PERSONAL,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            },
        )
        response.raise_for_status()
        token_data = response.json()
        token_data["expires_at"] = (
            datetime.now(timezone.utc) + timedelta(seconds=token_data.get("expires_in", 3600))
        ).isoformat()
        tokens["personal_app"] = token_data
        save_tokens(tokens)
        return token_data["access_token"]


async def get_valid_token(account: str) -> str:
    if account == "personal":
        return await get_app_token()
    tokens = load_tokens()
    token_data = tokens.get(account)
    if not token_data:
        raise Exception(f"No token for {account} — OAuth login required")
    expires_at = _parse_expires_at(token_data["expires_at"])
    if datetime.now(timezone.utc) >= expires_at - timedelta(minutes=5):
        token_data = await refresh_token(account)
    return token_data["access_token"]


async def get_emails(account: str) -> list:
    from connectors.ical import parse_ical_string
    import base64 as _base64
    token = await get_valid_token(account)
    base = _mailbox_base(account)
    all_emails = []

    async with httpx.AsyncClient(timeout=60.0) as client:
        url = f"{base}/mailFolders/inbox/messages"
        params = {
            "$top": 100,
            "$select": "id,internetMessageId,subject,from,receivedDateTime,bodyPreview,isRead,body,hasAttachments",
        }

        while url:
            response = await client.get(
                url,
                headers={"Authorization": f"Bearer {token}"},
                params=params if url == f"{base}/mailFolders/inbox/messages" else None,
            )
            response.raise_for_status()
            data = response.json()
            for email in data.get("value", []):
                raw_body = email.get("body", {})
                content = raw_body.get("content", "")
                content_type = raw_body.get("contentType", "text")
                email["fullBody"] = strip_html(content) if content_type == "html" else content.strip()
                all_emails.append(email)
            url = data.get("@odata.nextLink")

    # For emails with attachments, fetch attachment list and parse any text/calendar part
    emails_with_attachments = [e for e in all_emails if e.get("hasAttachments")]
    if emails_with_attachments:
        async def _fetch_ical(email: dict) -> None:
            try:
                async with httpx.AsyncClient(timeout=15.0) as client:
                    resp = await client.get(
                        f"{base}/messages/{email['id']}/attachments",
                        headers={"Authorization": f"Bearer {token}"},
                        params={"$select": "contentType,contentBytes,name"},
                    )
                    if resp.status_code != 200:
                        return
                    for att in resp.json().get("value", []):
                        ct = att.get("contentType", "")
                        if ct.startswith("text/calendar") or ct.startswith("application/ics"):
                            raw = _base64.b64decode(att.get("contentBytes", ""))
                            parsed = parse_ical_string(raw.decode("utf-8", errors="replace"))
                            if parsed:
                                email["ical_event"] = parsed
                            return
            except Exception as e:
                logger.debug(f"get_emails: attachment fetch failed for {email.get('id')}: {e}")

        await asyncio.gather(*[_fetch_ical(e) for e in emails_with_attachments], return_exceptions=True)

    return all_emails


async def get_sent_emails(account: str, days: int = 90) -> list:
    token = await get_valid_token(account)
    base = _mailbox_base(account)
    since = (datetime.now(timezone.utc) - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
    all_emails = []

    async with httpx.AsyncClient(timeout=60.0) as client:
        url = f"{base}/mailFolders/sentitems/messages"
        params = {
            "$top": 100,
            "$select": "id,subject,sentDateTime,body",
            "$filter": f"sentDateTime ge {since}",
            "$orderby": "sentDateTime desc",
        }

        while url:
            response = await client.get(
                url,
                headers={"Authorization": f"Bearer {token}"},
                params=params if url == f"{base}/mailFolders/sentitems/messages" else None,
            )
            response.raise_for_status()
            data = response.json()
            for email in data.get("value", []):
                raw_body = email.get("body", {})
                content = raw_body.get("content", "")
                content_type = raw_body.get("contentType", "text")
                body_text = strip_html(content) if content_type == "html" else content.strip()
                all_emails.append({
                    "id": email["id"],
                    "subject": email.get("subject", "(no subject)"),
                    "body": body_text,
                    "sent_at": email.get("sentDateTime"),
                })
            url = data.get("@odata.nextLink")

    return all_emails


async def get_email_history(account: str, address: str, limit: int = 8) -> list[dict]:
    """
    Search across all mail folders for recent messages to/from a specific address.
    Returns list of {subject, date, direction, snippet} dicts, newest first.
    Fails silently — returns [] on any error.
    """
    try:
        token = await get_valid_token(account)
        base  = _mailbox_base(account)
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.get(
                f"{base}/messages",
                headers={"Authorization": f"Bearer {token}"},
                params={
                    "$search":  f'"{address}"',
                    "$select":  "subject,receivedDateTime,sentDateTime,from,toRecipients,bodyPreview,isDraft",
                    "$top":     limit,
                },
            )
            if not resp.is_success:
                return []
            items = []
            for msg in resp.json().get("value", []):
                if msg.get("isDraft"):
                    continue
                from_addr = (msg.get("from") or {}).get("emailAddress", {}).get("address", "").lower()
                date      = (msg.get("receivedDateTime") or msg.get("sentDateTime") or "")[:10]
                direction = "←" if from_addr == address.lower() else "→"
                subject   = (msg.get("subject") or "").strip()
                snippet   = (msg.get("bodyPreview") or "").strip()[:200]
                items.append({
                    "subject":   subject,
                    "date":      date,
                    "direction": direction,
                    "snippet":   snippet,
                })
            return items
    except Exception:
        return []


async def get_user_profile(account: str) -> dict:
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.get(
            _mailbox_base(account),
            headers={"Authorization": f"Bearer {token}"},
        )
        response.raise_for_status()
        return response.json()


async def search_contacts(account: str, query: str, limit: int = 10) -> list[dict]:
    """
    Search M365 contacts by display name or email using $search.
    Returns [{ name, email, source: 'contacts' }].
    Fails silently — never raises.
    """
    if not query:
        return []
    try:
        token = await get_valid_token(account)
        base  = _mailbox_base(account)
        async with httpx.AsyncClient(timeout=15.0) as client:
            resp = await client.get(
                f"{base}/contacts",
                headers={
                    "Authorization": f"Bearer {token}",
                    "ConsistencyLevel": "eventual",
                },
                params={
                    "$search": f'"{query}"',
                    "$top": limit,
                    "$select": "displayName,emailAddresses",
                },
            )
        if not resp.is_success:
            return []
        results = []
        for item in resp.json().get("value", []):
            name = (item.get("displayName") or "").strip()
            addresses = item.get("emailAddresses") or []
            email = next((a.get("address", "") for a in addresses if a.get("address")), "")
            email = email.strip()
            if not email:
                continue
            results.append({"name": name or email, "email": email, "source": "contacts"})
        return results
    except Exception as e:
        logger.debug(f"Graph contact search failed ({account}): {e}")
        return []


async def delete_email(account: str, email_id: str):
    """Soft delete — moves to Deleted Items."""
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{_mailbox_base(account)}/messages/{email_id}/move",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"destinationId": "deleteditems"},
        )
        response.raise_for_status()


async def get_message_graph_id(account: str, internet_message_id: str) -> str | None:
    """Find the current folder-scoped Graph ID for a message using its stable
    internetMessageId. Searches across all folders so works after archive/move."""
    token = await get_valid_token(account)
    safe_id = internet_message_id.strip('<>').replace("'", "''")
    async with httpx.AsyncClient(timeout=15.0) as client:
        resp = await client.get(
            f"{_mailbox_base(account)}/messages",
            headers={"Authorization": f"Bearer {token}"},
            params={
                "$filter": f"internetMessageId eq '{safe_id}'",
                "$select": "id,internetMessageId",
                "$top": 1,
            },
        )
        resp.raise_for_status()
        items = resp.json().get("value", [])
        return items[0]["id"] if items else None


async def get_message_event(account: str, graph_id: str) -> dict | None:
    """Fetch structured event data for a single eventMessage via $expand=event."""
    token = await get_valid_token(account)
    async with httpx.AsyncClient(timeout=15.0) as client:
        resp = await client.get(
            f"{_mailbox_base(account)}/messages/{graph_id}",
            headers={"Authorization": f"Bearer {token}"},
            params={"$select": "subject,meetingMessageType", "$expand": "event"},
        )
        if resp.status_code != 200:
            return None
        data = resp.json()
        event_obj = data.get("event")
        if not event_obj:
            return None
        s   = event_obj.get("start", {})
        e   = event_obj.get("end", {})
        org = (event_obj.get("organizer") or {}).get("emailAddress", {})
        return {
            "summary":   data.get("subject", ""),
            "start":     s.get("dateTime"),
            "start_tz":  s.get("timeZone", "UTC"),
            "end":       e.get("dateTime"),
            "end_tz":    e.get("timeZone", "UTC"),
            "organizer": org.get("address", ""),
            "location":  (event_obj.get("location") or {}).get("displayName", ""),
            "source":    "graph",
        }


async def hard_delete_email(account: str, email_id: str):
    """Permanent delete — used only for autonomous spam deletion."""
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.delete(
            f"{_mailbox_base(account)}/messages/{email_id}",
            headers={"Authorization": f"Bearer {token}"},
        )
        response.raise_for_status()


async def archive_email(account: str, email_id: str):
    """Move to Archive folder."""
    folder_id = await get_or_create_folder(account, "Archive")
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{_mailbox_base(account)}/messages/{email_id}/move",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"destinationId": folder_id},
        )
        response.raise_for_status()


async def unarchive_email(account: str, email_id: str):
    """Move from Archive back to Inbox."""
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{_mailbox_base(account)}/messages/{email_id}/move",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"destinationId": "inbox"},
        )
        response.raise_for_status()


async def mark_as_read(account: str, email_id: str):
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.patch(
            f"{_mailbox_base(account)}/messages/{email_id}",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"isRead": True},
        )
        response.raise_for_status()


async def send_email(account: str, to: str | list, subject: str, body: str, cc: list[str] = None):
    token = await get_valid_token(account)
    addresses = to if isinstance(to, list) else [to]
    recipients = [{"emailAddress": {"address": a}} for a in addresses]
    message = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body},
        "toRecipients": recipients,
    }
    if cc:
        message["ccRecipients"] = [{"emailAddress": {"address": a}} for a in cc]
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{_mailbox_base(account)}/sendMail",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"message": message, "saveToSentItems": True},
        )
        response.raise_for_status()


async def reply_to_email(account: str, original_graph_id: str, to: list[str], body: str, cc: list[str] = None):
    """Reply to a message using Graph's /reply action.

    Unlike sendMail, this inherits the conversation context so the reply
    appears in the correct thread in Outlook and Sent Items.
    """
    token = await get_valid_token(account)
    message = {
        "toRecipients": [{"emailAddress": {"address": a}} for a in to],
        "body": {"contentType": "Text", "content": body},
    }
    if cc:
        message["ccRecipients"] = [{"emailAddress": {"address": a}} for a in cc]
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{_mailbox_base(account)}/messages/{original_graph_id}/reply",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"message": message},
        )
        response.raise_for_status()


async def get_or_create_folder(account: str, folder_name: str) -> str:
    token = await get_valid_token(account)
    base = _mailbox_base(account)
    async with httpx.AsyncClient() as client:
        response = await client.get(
            f"{base}/mailFolders",
            headers={"Authorization": f"Bearer {token}"},
            params={"$top": 50},
        )
        response.raise_for_status()
        folders = response.json().get("value", [])
        for folder in folders:
            if folder["displayName"].lower() == folder_name.lower():
                return folder["id"]
        response = await client.post(
            f"{base}/mailFolders",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"displayName": folder_name},
        )
        response.raise_for_status()
        return response.json()["id"]


async def accept_calendar_event(account: str, email_id: str):
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{_mailbox_base(account)}/messages/{email_id}/accept",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"sendResponse": True},
        )
        response.raise_for_status()


async def decline_calendar_event(account: str, email_id: str):
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{_mailbox_base(account)}/messages/{email_id}/decline",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"sendResponse": True},
        )
        response.raise_for_status()


def _parse_graph_dt(dt_str: str, tz_str: str = "UTC") -> datetime | None:
    """Parse a Graph API dateTime string to a UTC-aware datetime."""
    if not dt_str:
        return None
    try:
        s = re.sub(r'\.\d+', '', dt_str).replace('Z', '+00:00')
        if '+' in s[10:] or s.endswith('+00:00'):
            dt = datetime.fromisoformat(s)
        else:
            from zoneinfo import ZoneInfo
            try:
                dt = datetime.fromisoformat(s).replace(tzinfo=ZoneInfo(tz_str or "UTC"))
            except Exception:
                dt = datetime.fromisoformat(s).replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return None


async def get_busy_windows(account: str, start_dt: datetime, end_dt: datetime) -> list[tuple]:
    """
    Return (start, end) UTC datetime tuples for all non-free, non-cancelled events
    in the given window. Follows @odata.nextLink for pagination.
    Fails silently — returns [] on any error.
    """
    try:
        token = await get_valid_token(account)
        base  = _mailbox_base(account)
        url   = f"{base}/calendarView"
        params = {
            "startDateTime": start_dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "endDateTime":   end_dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "$select": "start,end,showAs,isCancelled",
            "$top": 100,
        }
        busy = []
        async with httpx.AsyncClient(timeout=20.0) as client:
            while url:
                resp = await client.get(url, headers={"Authorization": f"Bearer {token}"}, params=params)
                resp.raise_for_status()
                data = resp.json()
                for ev in data.get("value", []):
                    if ev.get("isCancelled"):
                        continue
                    if ev.get("showAs", "busy") == "free":
                        continue
                    s = ev.get("start", {})
                    e = ev.get("end", {})
                    s_dt = _parse_graph_dt(s.get("dateTime"), s.get("timeZone"))
                    e_dt = _parse_graph_dt(e.get("dateTime"), e.get("timeZone"))
                    if s_dt and e_dt:
                        busy.append((s_dt, e_dt))
                url = data.get("@odata.nextLink")
                params = {}
        return busy
    except Exception as e:
        logger.debug(f"get_busy_windows failed ({account}): {e}")
        return []


async def create_calendar_hold(account: str, start_iso: str, end_iso: str, title: str = "Hold") -> str:
    """
    Create a tentative calendar hold with no client-identifiable details.
    Returns the Graph event id, or "" on failure.
    """
    try:
        from zoneinfo import ZoneInfo
        token = await get_valid_token(account)
        base  = _mailbox_base(account)
        # Parse ISO strings to UTC-aware datetimes then format for Graph
        s_dt = datetime.fromisoformat(start_iso).astimezone(timezone.utc)
        e_dt = datetime.fromisoformat(end_iso).astimezone(timezone.utc)
        body_payload = {
            "subject": title,
            "showAs": "tentative",
            "isAllDay": False,
            "start": {"dateTime": s_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
            "end":   {"dateTime": e_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
        }
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.post(
                f"{base}/calendar/events",
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json=body_payload,
            )
            resp.raise_for_status()
            return resp.json().get("id", "")
    except Exception as e:
        logger.warning(f"create_calendar_hold failed ({account}): {e}")
        return ""


async def create_confirmed_event(account: str, start_iso: str, end_iso: str,
                                  title: str, client_email: str, client_name: str = "") -> tuple[str, str]:
    """
    Create a confirmed calendar event with a Teams meeting and the client as a required attendee.
    Graph sends them an invite automatically. Returns (event_id, teams_join_url), or ("", "") on failure.
    """
    try:
        token = await get_valid_token(account)
        base  = _mailbox_base(account)
        s_dt  = datetime.fromisoformat(start_iso).astimezone(timezone.utc)
        e_dt  = datetime.fromisoformat(end_iso).astimezone(timezone.utc)
        attendee = {"emailAddress": {"address": client_email, "name": client_name or client_email}, "type": "required"}
        payload = {
            "subject": title,
            "showAs": "busy",
            "isAllDay": False,
            "start": {"dateTime": s_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
            "end":   {"dateTime": e_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
            "attendees": [attendee],
            "isOnlineMeeting": True,
            "onlineMeetingProvider": "teamsForBusiness",
        }
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.post(
                f"{base}/calendar/events",
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json=payload,
            )
            resp.raise_for_status()
            data     = resp.json()
            event_id = data.get("id", "")
            join_url = (data.get("onlineMeeting") or {}).get("joinUrl", "")
            return event_id, join_url
    except Exception as e:
        logger.warning(f"create_confirmed_event failed ({account}): {e}")
        return "", ""


async def create_online_hold(account: str, start_iso: str, end_iso: str, title: str = "Hold") -> tuple[str, str]:
    """
    Create a tentative calendar hold with a Teams meeting link (no attendees).
    Used when proposing a specific time so the join URL can be included in the email.
    Returns (event_id, teams_join_url), or ("", "") on failure.
    """
    try:
        token = await get_valid_token(account)
        base  = _mailbox_base(account)
        s_dt  = datetime.fromisoformat(start_iso).astimezone(timezone.utc)
        e_dt  = datetime.fromisoformat(end_iso).astimezone(timezone.utc)
        payload = {
            "subject": title,
            "showAs": "tentative",
            "isAllDay": False,
            "start": {"dateTime": s_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
            "end":   {"dateTime": e_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
            "isOnlineMeeting": True,
            "onlineMeetingProvider": "teamsForBusiness",
        }
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.post(
                f"{base}/calendar/events",
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json=payload,
            )
            resp.raise_for_status()
            data     = resp.json()
            event_id = data.get("id", "")
            join_url = (data.get("onlineMeeting") or {}).get("joinUrl", "")
            return event_id, join_url
    except Exception as e:
        logger.warning(f"create_online_hold failed ({account}): {e}")
        return "", ""


async def delete_calendar_event(account: str, event_id: str) -> bool:
    """Delete a calendar event by its Graph event id. Returns True on success."""
    try:
        token = await get_valid_token(account)
        async with httpx.AsyncClient(timeout=20.0) as client:
            resp = await client.delete(
                f"{_mailbox_base(account)}/calendar/events/{event_id}",
                headers={"Authorization": f"Bearer {token}"},
            )
            return resp.status_code in (200, 204)
    except Exception as e:
        logger.warning(f"delete_calendar_event failed ({account}): {e}")
        return False


async def move_email(account: str, email_id: str, folder_name: str):
    folder_id = await get_or_create_folder(account, folder_name)
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{_mailbox_base(account)}/messages/{email_id}/move",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"destinationId": folder_id},
        )
        response.raise_for_status()


async def list_folders(account: str) -> list:
    """Return a flat list of all mail folders for the account, including one level of subfolders."""
    token = await get_valid_token(account)
    base = _mailbox_base(account)
    folders = []
    first = True
    url = f"{base}/mailFolders"
    async with httpx.AsyncClient(timeout=30.0) as client:
        while url:
            resp = await client.get(
                url,
                headers={"Authorization": f"Bearer {token}"},
                params={"$top": 100, "$expand": "childFolders"} if first else None,
            )
            first = False
            resp.raise_for_status()
            data = resp.json()
            for folder in data.get("value", []):
                folders.append({"id": folder["id"], "name": folder["displayName"]})
                children = folder.get("childFolders", [])
                if isinstance(children, dict):
                    children = children.get("value", [])
                for child in children:
                    folders.append({
                        "id": child["id"],
                        "name": f"{folder['displayName']} / {child['displayName']}",
                    })
            url = data.get("@odata.nextLink")
    return sorted(folders, key=lambda f: f["name"].lower())


async def file_to_folder(account: str, email_id: str, folder_id: str):
    """Move an email to a folder by ID — used for manual filing where folder_id is already known."""
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{_mailbox_base(account)}/messages/{email_id}/move",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"destinationId": folder_id},
        )
        response.raise_for_status()


async def export_mime(account: str, email_id: str) -> bytes:
    """Export a message as raw RFC 5322 MIME bytes — used for cross-account filing."""
    token = await get_valid_token(account)
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.get(
            f"{_mailbox_base(account)}/messages/{email_id}/$value",
            headers={"Authorization": f"Bearer {token}"},
        )
        resp.raise_for_status()
        return resp.content


async def import_mime(account: str, folder_id: str, mime_bytes: bytes) -> str:
    """Import raw MIME bytes into a specific folder — used for cross-account filing.
    Creates a copy of the message in the target folder."""
    token = await get_valid_token(account)
    async with httpx.AsyncClient(timeout=30.0) as client:
        resp = await client.post(
            f"{_mailbox_base(account)}/mailFolders/{folder_id}/messages",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "text/plain",
            },
            content=mime_bytes,
        )
        resp.raise_for_status()
        return resp.json().get("id", "")
