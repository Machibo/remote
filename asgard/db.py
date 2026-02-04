from __future__ import annotations

import sqlite3
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional


def get_connection(db_path: str = "data.db") -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def init_db(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,
            source_type TEXT,
            url TEXT NOT NULL UNIQUE,
            title TEXT,
            content TEXT,
            language TEXT,
            published_at TEXT,
            category TEXT,
            status TEXT NOT NULL DEFAULT 'new',
            summary TEXT,
            image_path TEXT,
            scheduled_at TEXT,
            notified_at TEXT,
            admin_message_id TEXT,
            channel_message_id TEXT,
            error TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
        """
    )
    conn.commit()


def now_iso() -> str:
    return datetime.utcnow().isoformat()


def upsert_item(conn: sqlite3.Connection, item: Dict[str, Any]) -> None:
    conn.execute(
        """
        INSERT INTO items (
            source, source_type, url, title, content, language,
            published_at, category, status, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(url) DO UPDATE SET
            title=excluded.title,
            content=excluded.content,
            published_at=excluded.published_at,
            updated_at=excluded.updated_at
        """,
        (
            item.get("source"),
            item.get("source_type"),
            item.get("url"),
            item.get("title"),
            item.get("content"),
            item.get("language"),
            item.get("published_at"),
            item.get("category"),
            item.get("status", "new"),
            now_iso(),
            now_iso(),
        ),
    )
    conn.commit()


def fetch_by_status(
    conn: sqlite3.Connection, status: str, limit: int = 20
) -> List[sqlite3.Row]:
    cursor = conn.execute(
        """
        SELECT * FROM items
        WHERE status = ?
        ORDER BY published_at DESC, id DESC
        LIMIT ?
        """,
        (status, limit),
    )
    return list(cursor.fetchall())


def fetch_prepared_unnotified(conn: sqlite3.Connection, limit: int = 10) -> List[sqlite3.Row]:
    cursor = conn.execute(
        """
        SELECT * FROM items
        WHERE status = 'prepared' AND notified_at IS NULL
        ORDER BY published_at DESC, id DESC
        LIMIT ?
        """,
        (limit,),
    )
    return list(cursor.fetchall())


def mark_prepared(
    conn: sqlite3.Connection,
    item_id: int,
    summary: str,
    image_path: Optional[str],
    category: Optional[str],
) -> None:
    conn.execute(
        """
        UPDATE items
        SET status = 'prepared',
            summary = ?,
            image_path = ?,
            category = COALESCE(?, category),
            updated_at = ?
        WHERE id = ?
        """,
        (summary, image_path, category, now_iso(), item_id),
    )
    conn.commit()


def mark_notified(conn: sqlite3.Connection, item_id: int, message_id: str | None) -> None:
    conn.execute(
        """
        UPDATE items
        SET notified_at = ?,
            admin_message_id = COALESCE(?, admin_message_id),
            updated_at = ?
        WHERE id = ?
        """,
        (now_iso(), message_id, now_iso(), item_id),
    )
    conn.commit()


def set_status(
    conn: sqlite3.Connection,
    item_id: int,
    status: str,
    scheduled_at: Optional[str] = None,
) -> None:
    conn.execute(
        """
        UPDATE items
        SET status = ?,
            scheduled_at = COALESCE(?, scheduled_at),
            updated_at = ?
        WHERE id = ?
        """,
        (status, scheduled_at, now_iso(), item_id),
    )
    conn.commit()


def mark_posted(conn: sqlite3.Connection, item_id: int, message_id: str | None) -> None:
    conn.execute(
        """
        UPDATE items
        SET status = 'posted',
            channel_message_id = COALESCE(?, channel_message_id),
            updated_at = ?
        WHERE id = ?
        """,
        (message_id, now_iso(), item_id),
    )
    conn.commit()


def fetch_due_posts(conn: sqlite3.Connection, now_iso_str: str, limit: int = 10) -> List[sqlite3.Row]:
    cursor = conn.execute(
        """
        SELECT * FROM items
        WHERE status = 'approved'
          AND scheduled_at IS NOT NULL
          AND scheduled_at <= ?
        ORDER BY scheduled_at ASC, id ASC
        LIMIT ?
        """,
        (now_iso_str, limit),
    )
    return list(cursor.fetchall())


def set_setting(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute(
        """
        INSERT INTO settings (key, value)
        VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value = excluded.value
        """,
        (key, value),
    )
    conn.commit()


def get_setting(conn: sqlite3.Connection, key: str) -> Optional[str]:
    cursor = conn.execute("SELECT value FROM settings WHERE key = ?", (key,))
    row = cursor.fetchone()
    if not row:
        return None
    return row["value"]


def fetch_items_by_ids(conn: sqlite3.Connection, ids: Iterable[int]) -> List[sqlite3.Row]:
    ids_list = list(ids)
    if not ids_list:
        return []
    placeholders = ",".join(["?"] * len(ids_list))
    cursor = conn.execute(
        f"SELECT * FROM items WHERE id IN ({placeholders})",
        ids_list,
    )
    return list(cursor.fetchall())
