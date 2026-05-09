import os
import httpx
import json
import re
from datetime import datetime, timedelta
from dotenv import load_dotenv

load_dotenv("config/.env")

CLIENT_ID = os.getenv("AZURE_CLIENT_ID_FINANCIAL")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET_FINANCIAL")
CLIENT_ID_PERSONAL = os.getenv("AZURE_CLIENT_ID_PERSONAL")
CLIENT_SECRET_PERSONAL = os.getenv("AZURE_CLIENT_SECRET_PERSONAL")
TENANT_ID = os.getenv("AZURE_TENANT_ID")
PERSONAL_EMAIL = os.getenv("PERSONAL_EMAIL", "***PERSONAL_EMAIL***")

TAILSCALE_IP = os.getenv("TAILSCALE_IP", "localhost")
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
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f, indent=2)


def _mailbox_base(account: str) -> str:
    """Return the Graph API base path for a given account's mailbox."""
    if account == "personal":
        return f"{GRAPH_BASE}/users/{PERSONAL_EMAIL}"
    return f"{GRAPH_BASE}/me"


def get_auth_url(account: str) -> str:
    return (
        f"{AUTHORITY}/oauth2/v2.0/authorize"
        f"?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri=http://localhost:8000/auth/callback/{account}"
        f"&response_mode=query"
        f"&scope={SCOPES.replace(' ', '%20')}"
        f"&state={account}"
        f"&prompt=select_account"
    )


async def exchange_code_for_token(code: str, account: str) -> dict:
    redirect_uri = f"http://localhost:8000/auth/callback/{account}"
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
        token_data = response.json()
        token_data["account"] = account
        token_data["expires_at"] = (
            datetime.utcnow() + timedelta(seconds=token_data.get("expires_in", 3600))
        ).isoformat()
        tokens = load_tokens()
        tokens[account] = token_data
        save_tokens(tokens)
        return token_data


async def refresh_token(account: str) -> dict:
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
        new_token = response.json()
        new_token["account"] = account
        new_token["expires_at"] = (
            datetime.utcnow() + timedelta(seconds=new_token.get("expires_in", 3600))
        ).isoformat()
        tokens[account] = new_token
        save_tokens(tokens)
        return new_token


async def get_app_token() -> str:
    """Client credentials flow for the personal account — no user sign-in required."""
    tokens = load_tokens()
    cached = tokens.get("personal_app", {})
    if cached.get("access_token") and cached.get("expires_at"):
        if datetime.utcnow() < datetime.fromisoformat(cached["expires_at"]) - timedelta(minutes=5):
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
            datetime.utcnow() + timedelta(seconds=token_data.get("expires_in", 3600))
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
    expires_at = datetime.fromisoformat(token_data["expires_at"])
    if datetime.utcnow() >= expires_at - timedelta(minutes=5):
        token_data = await refresh_token(account)
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


async def get_emails(account: str, count: int = None) -> list:
    token = await get_valid_token(account)
    base = _mailbox_base(account)
    all_emails = []

    async with httpx.AsyncClient(timeout=60.0) as client:
        url = f"{base}/mailFolders/inbox/messages"
        params = {
            "$top": 100,
            "$select": "id,internetMessageId,subject,from,receivedDateTime,bodyPreview,isRead,body",
            "$orderby": "receivedDateTime desc",
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

    return all_emails


async def get_sent_emails(account: str, days: int = 90) -> list:
    token = await get_valid_token(account)
    base = _mailbox_base(account)
    since = (datetime.utcnow() - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
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


async def get_user_profile(account: str) -> dict:
    token = await get_valid_token(account)
    async with httpx.AsyncClient() as client:
        response = await client.get(
            _mailbox_base(account),
            headers={"Authorization": f"Bearer {token}"},
        )
        return response.json()


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


async def send_email(account: str, to: str | list, subject: str, body: str):
    token = await get_valid_token(account)
    addresses = to if isinstance(to, list) else [to]
    recipients = [{"emailAddress": {"address": a}} for a in addresses]
    async with httpx.AsyncClient() as client:
        response = await client.post(
            f"{_mailbox_base(account)}/sendMail",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={
                "message": {
                    "subject": subject,
                    "body": {"contentType": "Text", "content": body},
                    "toRecipients": recipients,
                },
                "saveToSentItems": True,
            },
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
