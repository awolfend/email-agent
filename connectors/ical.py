"""
Parse iCal (RFC 5545) strings into structured dicts.
Used to extract event details from inbound calendar invite emails.
"""
import logging
from datetime import datetime, date, timezone

logger = logging.getLogger(__name__)


def parse_ical_string(ics_text: str) -> dict | None:
    """
    Parse a VCALENDAR string and return the first VEVENT as a normalised dict.
    Returns None if parsing fails or no VEVENT is found.
    """
    if not ics_text or not ics_text.strip():
        return None
    try:
        from icalendar import Calendar
        cal = Calendar.from_ical(ics_text)
        for component in cal.walk():
            if component.name != "VEVENT":
                continue

            dtstart = component.get("DTSTART")
            dtend   = component.get("DTEND") or component.get("DURATION")

            start_dt = dtstart.dt if dtstart else None
            end_dt   = dtend.dt   if dtend   else None

            # Normalise to ISO strings — both date and datetime objects have .isoformat()
            def _iso(val):
                if val is None:
                    return None
                if isinstance(val, datetime):
                    # Make timezone-aware if naive (assume UTC)
                    if val.tzinfo is None:
                        val = val.replace(tzinfo=timezone.utc)
                    return val.isoformat()
                if isinstance(val, date):
                    return val.isoformat()
                return str(val)

            organizer = component.get("ORGANIZER", "")
            org_email = str(organizer).lower().replace("mailto:", "").strip() if organizer else ""

            return {
                "summary":     str(component.get("SUMMARY", "")).strip(),
                "start":       _iso(start_dt),
                "end":         _iso(end_dt),
                "organizer":   org_email,
                "location":    str(component.get("LOCATION", "")).strip(),
                "uid":         str(component.get("UID", "")).strip(),
                "description": str(component.get("DESCRIPTION", "")).strip()[:500],
            }
    except Exception as e:
        logger.debug(f"ical parse failed: {e}")
    return None
