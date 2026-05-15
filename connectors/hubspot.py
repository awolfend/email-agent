"""
HubSpot CRM context for financial email drafting.

Uses:
  - CRM v3 contacts API  — contact search and properties
  - Engagements v1 API   — all activity (notes, meetings, calls, emails, tasks)
    in a single call per contact, no extra scopes required beyond contacts.read
"""
import asyncio
import logging
import os
import re
from datetime import datetime, timezone

import httpx

logger = logging.getLogger(__name__)

_BASE_CRM = "https://api.hubapi.com/crm/v3"
_BASE_ENG = "https://api.hubapi.com/engagements/v1"
_TIMEOUT  = 15.0

# Max recent records per engagement type shown in context
_MAX = {"NOTE": 4, "MEETING": 3, "CALL": 3, "TASK": 5}


def _headers() -> dict:
    return {
        "Authorization": f"Bearer {os.getenv('HUBSPOT_API_KEY', '')}",
        "Content-Type": "application/json",
    }


def _ms_to_date(ms) -> str:
    """Convert HubSpot millisecond timestamp to YYYY-MM-DD."""
    try:
        return datetime.fromtimestamp(int(ms) / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
    except Exception:
        return ""


def _strip_html(text: str) -> str:
    """Remove HTML tags and decode common entities."""
    text = re.sub(r"<[^>]+>", " ", text)
    text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    return re.sub(r"\s+", " ", text).strip()


async def _search_contact(client: httpx.AsyncClient, email: str) -> tuple[str | None, dict]:
    """Return (contact_id, properties) or (None, {})."""
    resp = await client.post(
        f"{_BASE_CRM}/objects/contacts/search",
        headers=_headers(),
        json={
            "filterGroups": [{"filters": [{"propertyName": "email", "operator": "EQ", "value": email}]}],
            "properties": ["firstname", "lastname", "company", "jobtitle", "lifecyclestage"],
            "limit": 1,
        },
    )
    if not resp.is_success:
        return None, {}
    results = resp.json().get("results", [])
    if not results:
        return None, {}
    row = results[0]
    return row["id"], row.get("properties", {})


async def _get_engagements(client: httpx.AsyncClient, contact_id: str) -> list[dict]:
    """Fetch all engagements for a contact via the v1 API (single call)."""
    resp = await client.get(
        f"{_BASE_ENG}/engagements/associated/CONTACT/{contact_id}/paged",
        headers=_headers(),
        params={"limit": 100},
    )
    if not resp.is_success:
        return []
    return resp.json().get("results", [])


def _format_engagements(items: list[dict]) -> dict[str, list[str]]:
    """
    Parse engagement list into per-type formatted strings, capped by _MAX.
    Returns dict keyed by type with list of formatted lines.
    """
    buckets: dict[str, list[tuple[int, str]]] = {}

    for item in items:
        eng  = item.get("engagement", {})
        meta = item.get("metadata", {})
        etype = eng.get("type", "")
        ts    = eng.get("timestamp") or eng.get("createdAt") or 0
        date  = _ms_to_date(ts)

        text = ""
        if etype == "NOTE":
            raw  = (meta.get("body") or "").strip()
            text = _strip_html(raw)[:400]
        elif etype in ("EMAIL", "INCOMING_EMAIL"):
            # Content is redacted by HubSpot API — show direction only
            direction = "outbound email" if etype == "EMAIL" else "inbound email"
            text = direction
        elif etype == "CALL":
            raw  = (meta.get("body") or meta.get("title") or "").strip()
            text = _strip_html(raw)[:200]
        elif etype == "MEETING":
            raw  = (meta.get("title") or meta.get("body") or "").strip()
            text = _strip_html(raw)[:200]
        elif etype == "TASK":
            subject = (meta.get("subject") or "").strip()
            status  = (meta.get("status") or "").strip().lower()
            text    = _strip_html(subject)[:120]
            if status and status != "completed":
                text += f" ({status})"
        else:
            continue

        if not text:
            continue

        if etype not in buckets:
            buckets[etype] = []
        buckets[etype].append((ts, f"  [{date}] {text}"))

    # Sort each bucket newest-first, cap at _MAX
    out = {}
    for etype, entries in buckets.items():
        entries.sort(key=lambda x: x[0], reverse=True)
        limit = _MAX.get(etype, 3)
        out[etype] = [line for _, line in entries[:limit]]

    return out


async def search_contacts(query: str, limit: int = 10) -> list[dict]:
    """
    Full-text search across HubSpot contacts by name or email.
    Returns [{ name, email, company, source: 'hubspot' }].
    Fails silently — never raises.
    """
    if not os.getenv("HUBSPOT_API_KEY", "") or not query:
        return []
    try:
        async with httpx.AsyncClient(timeout=_TIMEOUT) as client:
            resp = await client.post(
                f"{_BASE_CRM}/objects/contacts/search",
                headers=_headers(),
                json={
                    "query": query,
                    "properties": ["firstname", "lastname", "email", "company"],
                    "filterGroups": [{"filters": [
                        {"propertyName": "email", "operator": "HAS_PROPERTY"}
                    ]}],
                    "limit": limit,
                },
            )
        if not resp.is_success:
            return []
        results = []
        for row in resp.json().get("results", []):
            props = row.get("properties", {})
            email = (props.get("email") or "").strip()
            if not email:
                continue
            first = (props.get("firstname") or "").strip()
            last  = (props.get("lastname") or "").strip()
            name  = " ".join(p for p in [first, last] if p) or email
            results.append({
                "name": name,
                "email": email,
                "company": (props.get("company") or "").strip(),
                "source": "hubspot",
            })
        return results
    except Exception as e:
        logger.debug(f"HubSpot contact search failed: {e}")
        return []


async def get_contact_context(sender_email: str) -> str:
    """
    Look up sender_email in HubSpot and return a formatted CRM context block.
    Pulls: contact details, notes, meetings, emails (in/out), calls, tasks.
    Returns "" if HUBSPOT_API_KEY not set, contact not found, or any error.
    Always fails silently — never blocks draft generation.
    """
    if not os.getenv("HUBSPOT_API_KEY", ""):
        return ""
    try:
        async with httpx.AsyncClient(timeout=_TIMEOUT) as client:
            contact_id, props = await _search_contact(client, sender_email)
            if not contact_id:
                return ""
            engagements = await _get_engagements(client, contact_id)

        first    = (props.get("firstname") or "").strip()
        last     = (props.get("lastname") or "").strip()
        name     = " ".join(p for p in [first, last] if p)
        company  = (props.get("company") or "").strip()
        jobtitle = (props.get("jobtitle") or "").strip()
        lifecycle = (props.get("lifecyclestage") or "").strip()

        lines = ["--- CRM context (HubSpot) ---"]

        identity = name
        if jobtitle:
            identity += f", {jobtitle}"
        if company:
            identity += f" — {company}"
        if lifecycle:
            identity += f" [{lifecycle}]"
        if identity:
            lines.append(f"Client: {identity}")

        by_type = _format_engagements(engagements)

        # Section labels — order matters: tasks first (actionable), then history
        for etype, label in [("TASK", "Open tasks"), ("NOTE", "Notes"),
                              ("MEETING", "Meetings"), ("CALL", "Calls")]:
            entries = by_type.get(etype, [])
            if entries:
                lines.append(f"{label}:")
                lines.extend(entries)

        lines.append("--- end CRM context ---")

        if len(lines) <= 3:
            return ""

        total = sum(len(v) for v in by_type.values())
        logger.info(f"HubSpot context loaded for {sender_email}: {name or '(unnamed)'} — {total} engagements")
        return "\n".join(lines)

    except Exception as e:
        logger.debug(f"HubSpot lookup failed for {sender_email}: {e}")
        return ""
