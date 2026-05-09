import aiosqlite
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "config", "agent.db")

async def init_db():
    async with aiosqlite.connect(DB_PATH) as db:
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
        except Exception:
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
        ]:
            try:
                await db.execute(f"ALTER TABLE action_log ADD COLUMN {col} {definition}")
            except Exception:
                pass
        try:
            await db.execute(
                "CREATE UNIQUE INDEX IF NOT EXISTS idx_email_id ON action_log (email_id)"
            )
        except Exception:
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
        except Exception:
            pass
        await db.commit()


async def log_action(account: str, email_id: str, subject: str, sender: str,
                     action: str, classification: str = None,
                     confidence: float = None, notes: str = None,
                     body: str = None, received_at: str = None):
    from datetime import datetime, timezone
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
                    received_at = COALESCE(received_at, ?)
                WHERE email_id = ?
            """, (timestamp, classification, confidence, notes, action, body,
                  received_at, email_id))
        else:
            await db.execute("""
                INSERT OR IGNORE INTO action_log
                    (timestamp, received_at, account, email_id, subject, sender, action,
                     classification, confidence, notes, status, flagged, body, stub)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', 0, ?, 0)
            """, (timestamp, received_at, account, email_id, subject, sender, action,
                  classification, confidence, notes, body))
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
    from datetime import datetime, timezone, timedelta
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
    from datetime import datetime, timezone
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
    from datetime import datetime, timezone
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


async def upsert_sender_rule(sender: str, classification: str, source: str = 'learned') -> None:
    """Create or update a sender classification rule.

    Learned rules strengthen on consistency (fire at count >= 2) and decay on conflict.
    Manual rules (source='manual') fire immediately and are not overwritten by learned data.
    Classifications of 'error' or 'unknown' are never stored.
    """
    if not sender or classification in ('error', 'unknown', ''):
        return
    from datetime import datetime, timezone
    now = datetime.now(timezone.utc).isoformat()
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute(
            "SELECT classification, count, source FROM sender_rules WHERE sender = ?", (sender,)
        ) as cur:
            row = await cur.fetchone()
        row = dict(row) if row else None

        if source == 'manual':
            if not row:
                await db.execute(
                    "INSERT INTO sender_rules (sender, classification, count, source, created_at, last_seen) "
                    "VALUES (?, ?, 1, 'manual', ?, ?)",
                    (sender, classification, now, now)
                )
            else:
                await db.execute(
                    "UPDATE sender_rules SET classification = ?, count = 1, source = 'manual', last_seen = ? "
                    "WHERE sender = ?",
                    (classification, now, sender)
                )
        else:
            if not row:
                await db.execute(
                    "INSERT INTO sender_rules (sender, classification, count, source, created_at, last_seen) "
                    "VALUES (?, ?, 1, 'learned', ?, ?)",
                    (sender, classification, now, now)
                )
            elif row['source'] == 'manual':
                # Manual rules are sticky — only update last_seen
                await db.execute(
                    "UPDATE sender_rules SET last_seen = ? WHERE sender = ?", (now, sender)
                )
            elif row['classification'] == classification:
                # Consistent signal — strengthen the rule
                await db.execute(
                    "UPDATE sender_rules SET count = count + 1, last_seen = ? WHERE sender = ?",
                    (now, sender)
                )
            else:
                # Conflicting signal — decay the rule; delete if fully eroded
                new_count = row['count'] - 1
                if new_count <= 0:
                    await db.execute("DELETE FROM sender_rules WHERE sender = ?", (sender,))
                else:
                    await db.execute(
                        "UPDATE sender_rules SET count = ?, last_seen = ? WHERE sender = ?",
                        (new_count, now, sender)
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
    from datetime import datetime, timezone
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
