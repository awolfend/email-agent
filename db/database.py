import aiosqlite
import os
import re
from datetime import datetime, timezone, timedelta


def _extract_addr(raw: str) -> str:
    """Extract bare email address from 'Name <addr>' or plain address string."""
    m = re.search(r'<([^>]+)>', raw or '')
    return (m.group(1) if m else (raw or '')).strip().lower()

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "config", "agent.db")

async def init_db():
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("PRAGMA journal_mode=WAL")
        await db.execute("""
            CREATE TABLE IF NOT EXISTS sent_examples (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                account TEXT NOT NULL,
                message_id TEXT NOT NULL,
                subject TEXT,
                body TEXT NOT NULL,
                sent_at TEXT,
                char_count INTEGER,
                imported_at TEXT NOT NULL
            )
        """)
        try:
            await db.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_sent_example_msg_id ON sent_examples (message_id)"
            )
        except aiosqlite.OperationalError:
            pass
        await db.execute("""
            CREATE TABLE IF NOT EXISTS voice_profiles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                account TEXT NOT NULL UNIQUE,
                profile TEXT NOT NULL,
                example_count INTEGER,
                generated_at TEXT NOT NULL
            )
        """)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS action_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                timestamp TEXT NOT NULL,
                received_at TEXT,
                account TEXT NOT NULL,
                email_id TEXT NOT NULL,
                subject TEXT,
                sender TEXT,
                action TEXT NOT NULL,
                classification TEXT,
                confidence REAL,
                notes TEXT,
                status TEXT DEFAULT 'pending',
                flagged INTEGER DEFAULT 0,
                draft_reply TEXT,
                body TEXT,
                stub INTEGER DEFAULT 0
            )
        """)
        for col, definition in [
            ("status", "TEXT DEFAULT 'pending'"),
            ("flagged", "INTEGER DEFAULT 0"),
            ("draft_reply", "TEXT"),
            ("body", "TEXT"),
            ("received_at", "TEXT"),
            ("stub", "INTEGER DEFAULT 0"),
            ("graph_id", "TEXT"),
            ("thread_id", "TEXT"),
            ("orig_message_id", "TEXT"),
        ]:
            try:
                await db.execute(f"ALTER TABLE action_log ADD COLUMN {col} {definition}")
            except aiosqlite.OperationalError:
                pass
        try:
            await db.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_email_id ON action_log (email_id)"
            )
        except aiosqlite.OperationalError:
            pass
        await db.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS sender_rules (
                sender          TEXT PRIMARY KEY,
                classification  TEXT NOT NULL,
                count           INTEGER DEFAULT 1,
                source          TEXT DEFAULT 'learned',
                created_at      TEXT NOT NULL,
                last_seen       TEXT NOT NULL
            )
        """)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS filing_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sender_domain TEXT NOT NULL,
                target_account TEXT NOT NULL,
                target_folder_id TEXT NOT NULL,
                target_folder_name TEXT NOT NULL,
                count INTEGER DEFAULT 1,
                last_filed_at TEXT NOT NULL
            )
        """)
        try:
            await db.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_filing_domain_folder "
                "ON filing_history (sender_domain, target_folder_id)"
            )
        except aiosqlite.OperationalError:
            pass
        await db.execute("""
            CREATE TABLE IF NOT EXISTS compose_drafts (
                id INTEGER PRIMARY KEY,
                account TEXT NOT NULL UNIQUE,
                to_address TEXT,
                cc_address TEXT,
                subject TEXT,
                body TEXT,
                prompt TEXT,
                updated_at TEXT
            )
        """)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS meeting_proposals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                outgoing_message_id TEXT,
                account TEXT NOT NULL,
                client_email TEXT,
                client_name TEXT,
                subject TEXT,
                duration_minutes INTEGER,
                sent_at TEXT,
                status TEXT DEFAULT 'pending'
            )
        """)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS meeting_slots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                proposal_id INTEGER NOT NULL REFERENCES meeting_proposals(id),
                slot_start TEXT NOT NULL,
                slot_end TEXT NOT NULL,
                auto_release_at TEXT NOT NULL,
                owning_calendar_event_id TEXT,
                mirror_event_id_financial TEXT,
                mirror_event_id_google_tax TEXT,
                status TEXT DEFAULT 'tentative'
            )
        """)
        # Schema migrations — additive only, safe to run every startup
        for stmt in [
            "ALTER TABLE meeting_proposals ADD COLUMN direction TEXT DEFAULT 'outbound'",
            "ALTER TABLE meeting_proposals ADD COLUMN triggering_email_id TEXT",
            "ALTER TABLE meeting_slots ADD COLUMN proposed_by TEXT DEFAULT 'us'",
            "ALTER TABLE action_log ADD COLUMN ical_data TEXT",
        ]:
            try:
                await db.execute(stmt)
            except aiosqlite.OperationalError:
                pass  # column already exists

        # Seed meeting settings defaults (INSERT OR IGNORE — never overwrite user values)
        for key, value in [
            ("meeting_hours_start", "09:00"),
            ("meeting_hours_end", "17:00"),
            ("meeting_buffer_minutes", "15"),
            ("meeting_default_duration", "60"),
            ("meeting_no_response_hours", "36"),
            ("meeting_post_slot_buffer_minutes", "30"),
        ]:
            await db.execute(
                "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (key, value)
            )
        await db.commit()


async def log_action(account: str, email_id: str, subject: str, sender: str,
                     action: str, classification: str = None,
                     confidence: float = None, notes: str = None,
                     body: str = None, received_at: str = None,
                     graph_id: str = None, thread_id: str = None,
                     orig_message_id: str = None, ical_data: str = None):
    timestamp = datetime.now(timezone.utc).isoformat()
    async with aiosqlite.connect(DB_PATH) as db:
        existing = await db.execute(
            "SELECT id FROM action_log WHERE email_id = ?", (email_id,)
        )
        row = await existing.fetchone()
        if row:
            await db.execute("""
                UPDATE action_log SET
                    timestamp = ?, classification = ?, confidence = ?,
                    notes = ?, action = ?, body = ?,
                    received_at = COALESCE(received_at, ?),
                    graph_id = COALESCE(?, graph_id),
                    thread_id = COALESCE(?, thread_id),
                    orig_message_id = COALESCE(?, orig_message_id),
                    ical_data = COALESCE(?, ical_data)
                WHERE email_id = ?
            """, (timestamp, classification, confidence, notes, action, body,
                  received_at, graph_id, thread_id, orig_message_id, ical_data, email_id))
        else:
            await db.execute("""
                INSERT OR IGNORE INTO action_log
                    (timestamp, received_at, account, email_id, subject, sender, action,
                     classification, confidence, notes, status, flagged, body, stub,
                     graph_id, thread_id, orig_message_id, ical_data)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', 0, ?, 0, ?, ?, ?, ?)
            """, (timestamp, received_at, account, email_id, subject, sender, action,
                  classification, confidence, notes, body,
                  graph_id, thread_id, orig_message_id, ical_data))
        await db.commit()


async def get_pending_ids_for_account(account: str) -> set:
    async with aiosqlite.connect(DB_PATH) as db:
        async with db.execute(
            "SELECT email_id FROM action_log WHERE account = ? AND status = 'pending' AND stub = 0",
            (account,)
        ) as cursor:
            rows = await cursor.fetchall()
            return {row[0] for row in rows}


async def mark_missing_as_archived(account: str, inbox_ids: set):
    pending_ids = await get_pending_ids_for_account(account)
    missing = pending_ids - inbox_ids
    if missing:
        async with aiosqlite.connect(DB_PATH) as db:
            for email_id in missing:
                await db.execute(
                    "UPDATE action_log SET status = 'archived', action = 'reconciled' "
                    "WHERE email_id = ? AND status = 'pending'",
                    (email_id,)
                )
            await db.commit()
    return missing


async def get_queue(history_limit: int = 1000):
    priority = {
        "action_required": 1, "calendar": 2, "fyi": 3,
        "notification": 4, "newsletter": 5, "spam": 6, "unknown": 7, "error": 8
    }
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        # All pending emails — no limit, inbox must always be complete
        async with db.execute("""
            SELECT * FROM action_log WHERE status = 'pending'
            ORDER BY COALESCE(received_at, timestamp) DESC
        """) as cursor:
            pending = [dict(row) for row in await cursor.fetchall()]
        # Recent history — bounded, sufficient for review
        async with db.execute("""
            SELECT * FROM action_log WHERE status != 'pending'
            ORDER BY COALESCE(received_at, timestamp) DESC
            LIMIT ?
        """, (history_limit,)) as cursor:
            history = [dict(row) for row in await cursor.fetchall()]
        items = pending + history
        items.sort(key=lambda x: priority.get(x.get("classification", "unknown"), 7))
        return items


async def get_email_by_id(email_id: str) -> dict:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT * FROM action_log WHERE email_id = ?", (email_id,)
        ) as cursor:
            row = await cursor.fetchone()
            return dict(row) if row else None


async def ensure_inbox_state(email_id: str, graph_id: str):
    """Inbox is source of truth. If an email is present in the live inbox,
    force its DB record to pending and refresh its graph_id."""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "UPDATE action_log SET status = 'pending', graph_id = COALESCE(?, graph_id) WHERE email_id = ?",
            (graph_id, email_id)
        )
        await db.commit()


async def update_ical_data(email_id: str, ical_json: str):
    """Write ical_data only when the existing row has NULL — used to backfill on retry polls."""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "UPDATE action_log SET ical_data = ? WHERE email_id = ? AND ical_data IS NULL",
            (ical_json, email_id)
        )
        await db.commit()


async def update_email_status(email_id: str, status: str, action: str = None):
    async with aiosqlite.connect(DB_PATH) as db:
        if action:
            await db.execute(
                "UPDATE action_log SET status = ?, action = ? WHERE email_id = ?",
                (status, action, email_id)
            )
        else:
            await db.execute(
                "UPDATE action_log SET status = ? WHERE email_id = ?",
                (status, email_id)
            )
        await db.commit()


async def update_email_classification(email_id: str, classification: str):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "UPDATE action_log SET classification = ? WHERE email_id = ?",
            (classification, email_id)
        )
        await db.commit()


async def update_draft_reply(email_id: str, draft: str):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "UPDATE action_log SET draft_reply = ? WHERE email_id = ?",
            (draft, email_id)
        )
        await db.commit()


async def toggle_flag(email_id: str):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "UPDATE action_log SET flagged = CASE WHEN flagged = 1 THEN 0 ELSE 1 END "
            "WHERE email_id = ?",
            (email_id,)
        )
        await db.commit()


async def delete_record(email_id: str):
    """Delete a single SQLite record by email_id. Does not affect the live mailbox."""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "DELETE FROM action_log WHERE email_id = ?", (email_id,)
        )
        await db.commit()


async def prune_old_records(days: int = 90) -> int:
    """Delete non-pending records older than `days` days. Returns count deleted."""
    cutoff = (datetime.now(timezone.utc) - timedelta(days=days)).isoformat()
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            "DELETE FROM action_log WHERE status != 'pending' AND timestamp < ?",
            (cutoff,)
        )
        count = cursor.rowcount
        await db.commit()
    return count


async def clear_history(scope: str) -> int:
    """Delete non-pending history records by scope. Returns count deleted."""
    valid_scopes = {"sent", "archived", "deleted", "all"}
    if scope not in valid_scopes:
        return 0
    async with aiosqlite.connect(DB_PATH) as db:
        if scope == "all":
            cursor = await db.execute(
                "DELETE FROM action_log WHERE status != 'pending'"
            )
        else:
            cursor = await db.execute(
                "DELETE FROM action_log WHERE status = ?", (scope,)
            )
        count = cursor.rowcount
        await db.commit()
        return count


async def save_sent_example(account: str, message_id: str, subject: str, body: str, sent_at: str = None) -> bool:
    imported_at = datetime.now(timezone.utc).isoformat()
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            "INSERT OR IGNORE INTO sent_examples (account, message_id, subject, body, sent_at, char_count, imported_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (account, message_id, subject, body, sent_at, len(body), imported_at)
        )
        await db.commit()
        return cursor.rowcount > 0


async def get_all_sent_examples(account: str) -> list:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT subject, body FROM sent_examples WHERE account = ? ORDER BY sent_at DESC",
            (account,)
        ) as cursor:
            rows = await cursor.fetchall()
            return [dict(row) for row in rows]


async def save_voice_profile(account: str, profile: str, example_count: int):
    generated_at = datetime.now(timezone.utc).isoformat()
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "INSERT INTO voice_profiles (account, profile, example_count, generated_at) VALUES (?, ?, ?, ?) "
            "ON CONFLICT(account) DO UPDATE SET profile=excluded.profile, example_count=excluded.example_count, generated_at=excluded.generated_at",
            (account, profile, example_count, generated_at)
        )
        await db.commit()


async def get_voice_profile(account: str) -> str | None:
    async with aiosqlite.connect(DB_PATH) as db:
        async with db.execute(
            "SELECT profile FROM voice_profiles WHERE account = ?", (account,)
        ) as cursor:
            row = await cursor.fetchone()
            return row[0] if row else None


async def get_voice_profile_meta(account: str) -> dict | None:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT account, example_count, generated_at FROM voice_profiles WHERE account = ?", (account,)
        ) as cursor:
            row = await cursor.fetchone()
            return dict(row) if row else None


async def get_setting(key: str) -> str | None:
    async with aiosqlite.connect(DB_PATH) as db:
        async with db.execute("SELECT value FROM settings WHERE key = ?", (key,)) as cursor:
            row = await cursor.fetchone()
            return row[0] if row else None


async def set_setting(key: str, value: str):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value)
        )
        await db.commit()


async def get_stats():
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        stats = {"by_classification": {}, "by_account": {}, "by_status": {}}
        async with db.execute("""
            SELECT classification, COUNT(*) as count
            FROM action_log WHERE status = 'pending'
            GROUP BY classification
        """) as cursor:
            for row in await cursor.fetchall():
                stats["by_classification"][row["classification"]] = row["count"]
        async with db.execute("""
            SELECT account, COUNT(*) as count
            FROM action_log WHERE status = 'pending'
            GROUP BY account
        """) as cursor:
            for row in await cursor.fetchall():
                stats["by_account"][row["account"]] = row["count"]
        async with db.execute("""
            SELECT status, COUNT(*) as count FROM action_log GROUP BY status
        """) as cursor:
            for row in await cursor.fetchall():
                stats["by_status"][row["status"]] = row["count"]
        return stats


async def upsert_sender_rule(sender: str, classification: str, source: str = 'manual') -> None:
    """Create or update a manual sender classification rule.

    Only manual rules are stored. A rule fires automatically (overrides AI) once
    the same classification has been manually confirmed twice (count >= 2).
    Setting a different classification resets the count to 1.
    Classifications of 'error' or 'unknown' are never stored.
    """
    if not sender or source != 'manual' or classification in ('error', 'unknown', ''):
        return
    now = datetime.now(timezone.utc).isoformat()
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT classification, count FROM sender_rules WHERE sender = ?", (sender,)
        ) as cur:
            row = await cur.fetchone()
        row = dict(row) if row else None

        if not row:
            await db.execute(
                "INSERT INTO sender_rules (sender, classification, count, source, created_at, last_seen) "
                "VALUES (?, ?, 1, 'manual', ?, ?)",
                (sender, classification, now, now)
            )
        elif row['classification'] == classification:
            # Same classification confirmed again — strengthen
            await db.execute(
                "UPDATE sender_rules SET count = count + 1, last_seen = ? WHERE sender = ?",
                (now, sender)
            )
        else:
            # Classification changed — reset count to 1
            await db.execute(
                "UPDATE sender_rules SET classification = ?, count = 1, last_seen = ? WHERE sender = ?",
                (classification, now, sender)
            )
        await db.commit()


async def get_sender_rule(sender: str) -> dict | None:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT * FROM sender_rules WHERE sender = ?", (sender,)
        ) as cur:
            row = await cur.fetchone()
    return dict(row) if row else None


async def get_all_sender_rules() -> list:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT * FROM sender_rules ORDER BY last_seen DESC"
        ) as cur:
            return [dict(r) for r in await cur.fetchall()]


async def delete_sender_rule(sender: str) -> bool:
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute(
            "DELETE FROM sender_rules WHERE sender = ?", (sender,)
        )
        await db.commit()
    return cursor.rowcount > 0


async def clear_all_sender_rules() -> int:
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute("DELETE FROM sender_rules")
        await db.commit()
    return cursor.rowcount


async def record_filing(sender_domain: str, target_account: str, target_folder_id: str, target_folder_name: str):
    now = datetime.now(timezone.utc).isoformat()
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            INSERT INTO filing_history (sender_domain, target_account, target_folder_id, target_folder_name, count, last_filed_at)
            VALUES (?, ?, ?, ?, 1, ?)
            ON CONFLICT(sender_domain, target_folder_id) DO UPDATE SET
                count = count + 1,
                target_folder_name = excluded.target_folder_name,
                last_filed_at = excluded.last_filed_at
        """, (sender_domain, target_account, target_folder_id, target_folder_name, now))
        await db.commit()


async def set_auth_error(account: str, message: str):
    await set_setting(f"auth_error_{account}", message[:500])


async def clear_auth_error(account: str):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("DELETE FROM settings WHERE key = ?", (f"auth_error_{account}",))
        await db.commit()


async def get_auth_errors() -> dict:
    result = {}
    async with aiosqlite.connect(DB_PATH) as db:
        for account in ("financial", "gmail", "personal"):
            async with db.execute(
                "SELECT value FROM settings WHERE key = ?", (f"auth_error_{account}",)
            ) as cursor:
                row = await cursor.fetchone()
                result[account] = row[0] if row else None
    return result


async def get_filing_suggestions(sender_domain: str, limit: int = 5) -> list:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute("""
            SELECT target_account, target_folder_id, target_folder_name, count
            FROM filing_history
            WHERE sender_domain = ?
            ORDER BY count DESC
            LIMIT ?
        """, (sender_domain, limit)) as cursor:
            rows = await cursor.fetchall()
            return [dict(row) for row in rows]


# ---------------------------------------------------------------------------
# compose_drafts — one row per account, restored when modal reopens
# ---------------------------------------------------------------------------

async def save_compose_draft(account: str, to_address: str = None, cc_address: str = None,
                              subject: str = None, body: str = None, prompt: str = None):
    now = datetime.now(timezone.utc).isoformat()
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            INSERT INTO compose_drafts (account, to_address, cc_address, subject, body, prompt, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(account) DO UPDATE SET
                to_address = excluded.to_address,
                cc_address = excluded.cc_address,
                subject = excluded.subject,
                body = excluded.body,
                prompt = excluded.prompt,
                updated_at = excluded.updated_at
        """, (account, to_address, cc_address, subject, body, prompt, now))
        await db.commit()


async def get_compose_draft(account: str) -> dict | None:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT * FROM compose_drafts WHERE account = ?", (account,)
        ) as cursor:
            row = await cursor.fetchone()
            return dict(row) if row else None


async def clear_compose_draft(account: str):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("DELETE FROM compose_drafts WHERE account = ?", (account,))
        await db.commit()


# ---------------------------------------------------------------------------
# meeting_proposals
# ---------------------------------------------------------------------------

async def save_meeting_proposal(outgoing_message_id: str, account: str, client_email: str,
                                 client_name: str, subject: str, duration_minutes: int,
                                 sent_at: str, direction: str = "outbound",
                                 triggering_email_id: str = None) -> int:
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute("""
            INSERT INTO meeting_proposals
                (outgoing_message_id, account, client_email, client_name, subject,
                 duration_minutes, sent_at, direction, triggering_email_id)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (outgoing_message_id, account, client_email, client_name, subject,
              duration_minutes, sent_at, direction, triggering_email_id))
        await db.commit()
        return cursor.lastrowid


async def get_meeting_proposal(proposal_id: int) -> dict | None:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT * FROM meeting_proposals WHERE id = ?", (proposal_id,)
        ) as cursor:
            row = await cursor.fetchone()
            return dict(row) if row else None


async def get_open_proposals_for_client(client_email: str) -> list:
    """Return pending proposals for a given client email — fuzzy address match."""
    addr = _extract_addr(client_email)
    if not addr or '@' not in addr:
        return []
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT * FROM meeting_proposals WHERE status = 'pending' ORDER BY sent_at DESC"
        ) as cursor:
            rows = [dict(r) for r in await cursor.fetchall()]
    def _matches(stored: str) -> bool:
        return _extract_addr(stored) == addr
    return [r for r in rows if _matches(r.get('client_email', ''))]


async def get_all_proposals() -> list:
    """Return all proposals with their slots, ordered pending-first."""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute("""
            SELECT * FROM meeting_proposals
            ORDER BY
                CASE status WHEN 'pending' THEN 0 WHEN 'confirmed' THEN 1
                            WHEN 'expired' THEN 2 WHEN 'declined' THEN 3 ELSE 4 END,
                sent_at DESC
        """) as cursor:
            proposals = [dict(r) for r in await cursor.fetchall()]
        if not proposals:
            return []
        ids = [p["id"] for p in proposals]
        placeholders = ",".join("?" * len(ids))
        async with db.execute(
            f"SELECT * FROM meeting_slots WHERE proposal_id IN ({placeholders}) ORDER BY slot_start ASC",
            ids
        ) as cursor:
            all_slots = [dict(r) for r in await cursor.fetchall()]
    slots_by = {}
    for s in all_slots:
        slots_by.setdefault(s["proposal_id"], []).append(s)
    for p in proposals:
        p["slots"] = slots_by.get(p["id"], [])
    return proposals


async def get_pending_calendar_invites() -> list[dict]:
    """Return pending emails that have an iCal attachment, ordered newest-first."""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            """SELECT email_id, account, subject, sender, received_at, ical_data, graph_id
               FROM action_log
               WHERE ical_data IS NOT NULL AND status = 'pending'
               ORDER BY received_at DESC"""
        ) as cursor:
            rows = await cursor.fetchall()
    result = []
    for r in rows:
        d = dict(r)
        if d.get("ical_data"):
            try:
                import json as _json
                d["ical_event"] = _json.loads(d["ical_data"])
            except Exception:
                d["ical_event"] = None
        result.append(d)
    return result


_VALID_PROPOSAL_STATUSES = {'pending', 'negotiating', 'confirmed', 'cancelled', 'expired', 'declined'}
_VALID_SLOT_STATUSES = {'tentative', 'confirmed', 'declined', 'released'}

async def update_proposal_status(proposal_id: int, status: str):
    if status not in _VALID_PROPOSAL_STATUSES:
        raise ValueError(f"Invalid proposal status: {status!r}")
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "UPDATE meeting_proposals SET status = ? WHERE id = ?", (status, proposal_id)
        )
        await db.commit()


# ---------------------------------------------------------------------------
# meeting_slots
# ---------------------------------------------------------------------------

async def save_meeting_slot(proposal_id: int, slot_start: str, slot_end: str,
                             auto_release_at: str, owning_calendar_event_id: str = None,
                             mirror_event_id_financial: str = None,
                             mirror_event_id_google_tax: str = None,
                             proposed_by: str = "us") -> int:
    async with aiosqlite.connect(DB_PATH) as db:
        cursor = await db.execute("""
            INSERT INTO meeting_slots
                (proposal_id, slot_start, slot_end, auto_release_at,
                 owning_calendar_event_id, mirror_event_id_financial, mirror_event_id_google_tax,
                 proposed_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (proposal_id, slot_start, slot_end, auto_release_at,
              owning_calendar_event_id, mirror_event_id_financial, mirror_event_id_google_tax,
              proposed_by))
        await db.commit()
        return cursor.lastrowid


async def get_slots_for_proposal(proposal_id: int) -> list:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT * FROM meeting_slots WHERE proposal_id = ? ORDER BY slot_start ASC",
            (proposal_id,)
        ) as cursor:
            return [dict(r) for r in await cursor.fetchall()]


async def update_slot_status(slot_id: int, status: str):
    if status not in _VALID_SLOT_STATUSES:
        raise ValueError(f"Invalid slot status: {status!r}")
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "UPDATE meeting_slots SET status = ? WHERE id = ?", (status, slot_id)
        )
        await db.commit()


async def get_expired_tentative_slots() -> list:
    """Return tentative slots whose auto_release_at has passed — for background expiry task."""
    now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S+00:00')
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute("""
            SELECT ms.*, mp.account
            FROM meeting_slots ms
            JOIN meeting_proposals mp ON ms.proposal_id = mp.id
            WHERE ms.status = 'tentative' AND ms.auto_release_at < ?
        """, (now,)) as cursor:
            return [dict(r) for r in await cursor.fetchall()]


async def search_contact_history(account: str, query: str, limit: int = 10) -> list[dict]:
    """
    Search action_log sender field for contacts matching query.
    Returns [{ name, email, source: 'history' }], deduped by email.
    """
    import re
    q = f"%{query}%"
    async with aiosqlite.connect(DB_PATH) as db:
        async with db.execute("""
            SELECT DISTINCT sender FROM action_log
            WHERE account = ? AND sender LIKE ? AND sender != ''
            ORDER BY timestamp DESC
            LIMIT ?
        """, (account, q, limit * 3)) as cursor:
            rows = [r[0] for r in await cursor.fetchall()]

    seen, results = set(), []
    for raw in rows:
        raw = raw.strip()
        # Parse "Display Name <email>" or bare email
        m = re.match(r'^(.*?)\s*<([^>]+)>$', raw)
        if m:
            name  = m.group(1).strip().strip('"')
            email = m.group(2).strip().lower()
        elif "@" in raw:
            email = raw.lower()
            name  = email
        else:
            continue
        if email and email not in seen:
            seen.add(email)
            results.append({"name": name or email, "email": email, "source": "history"})
        if len(results) >= limit:
            break
    return results


async def expire_completed_proposals():
    """Mark proposals as expired when all their slots have been released."""
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            UPDATE meeting_proposals SET status = 'expired'
            WHERE status = 'pending'
            AND id NOT IN (
                SELECT DISTINCT proposal_id FROM meeting_slots WHERE status = 'tentative'
            )
        """)
        await db.commit()
