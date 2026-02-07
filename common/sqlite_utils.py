from __future__ import annotations

import sqlite3
from typing import Iterable, List, Tuple


def connect(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON;")
    conn.execute("PRAGMA journal_mode = WAL;")
    conn.execute("PRAGMA synchronous = NORMAL;")
    return conn


def table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?;",
        (table_name,),
    ).fetchone()
    return row is not None


def clear_table(conn: sqlite3.Connection, table_name: str) -> None:
    if table_exists(conn, table_name):
        conn.execute(f"DELETE FROM {table_name};")


def executemany(conn: sqlite3.Connection, sql: str, rows: Iterable[Tuple]) -> None:
    conn.executemany(sql, list(rows))
