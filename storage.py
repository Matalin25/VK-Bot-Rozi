# -*- coding: utf-8 -*-
"""
Постоянное хранилище найденных заказов на SQLite.

Путь к базе:
    /data/orders.db  — на Amvera (постоянное хранилище)
    ./data/orders.db — локально / на прочих хостингах
"""
import os
import sqlite3
import threading
from datetime import datetime

_LOCK = threading.Lock()


def _db_path():
    base = "/data" if os.path.isdir("/data") and os.access("/data", os.W_OK) else "data"
    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "orders.db")


def _connect():
    conn = sqlite3.connect(_db_path(), timeout=10)
    conn.row_factory = sqlite3.Row
    return conn


def init():
    """Создаёт таблицы, если их ещё нет."""
    with _LOCK, _connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS orders (
                key         TEXT PRIMARY KEY,
                source      TEXT,
                name        TEXT,
                desc        TEXT,
                price       TEXT,
                link        TEXT,
                categories  TEXT,
                price_num   INTEGER,
                first_seen  TEXT,
                last_seen   TEXT
            )
            """
        )
        conn.execute(
            "CREATE TABLE IF NOT EXISTS runs ("
            " id INTEGER PRIMARY KEY AUTOINCREMENT,"
            " ts TEXT, sources TEXT, query TEXT, pages TEXT, found INTEGER, new INTEGER)"
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_last_seen ON orders(last_seen)")


def _row_key(r):
    return r.get("link") or f"{r.get('source','')}:{r.get('name','')}"


def save_run(rows, meta):
    """
    Сохраняет заказы. Возвращает список ключей, которые встретились ВПЕРВЫЕ
    (для подсветки «новых» заказов).
    """
    now = datetime.utcnow().isoformat(timespec="seconds")
    new_keys = []
    with _LOCK, _connect() as conn:
        for r in rows:
            key = _row_key(r)
            if not key:
                continue
            cats = ",".join(r.get("categories", []))
            cur = conn.execute("SELECT 1 FROM orders WHERE key=?", (key,))
            exists = cur.fetchone() is not None
            if exists:
                conn.execute("UPDATE orders SET last_seen=? WHERE key=?", (now, key))
            else:
                new_keys.append(key)
                conn.execute(
                    "INSERT INTO orders (key, source, name, desc, price, link,"
                    " categories, price_num, first_seen, last_seen)"
                    " VALUES (?,?,?,?,?,?,?,?,?,?)",
                    (key, r.get("source", ""), r.get("name", ""), r.get("desc", ""),
                     r.get("price", ""), r.get("link", ""), cats,
                     r.get("price_num"), now, now),
                )
        conn.execute(
            "INSERT INTO runs (ts, sources, query, pages, found, new)"
            " VALUES (?,?,?,?,?,?)",
            (now, ",".join(meta.get("sources", [])), meta.get("query", ""),
             str(meta.get("pages", "")), len(rows), len(new_keys)),
        )
    return new_keys


def known_keys():
    """Множество всех ключей заказов, когда-либо сохранённых."""
    with _LOCK, _connect() as conn:
        return {row["key"] for row in conn.execute("SELECT key FROM orders")}


def history(limit=500, query=""):
    """Последние сохранённые заказы (для восстановления сессии)."""
    with _LOCK, _connect() as conn:
        if query:
            like = f"%{query}%"
            cur = conn.execute(
                "SELECT * FROM orders WHERE name LIKE ? OR desc LIKE ?"
                " ORDER BY last_seen DESC LIMIT ?",
                (like, like, limit),
            )
        else:
            cur = conn.execute(
                "SELECT * FROM orders ORDER BY last_seen DESC LIMIT ?", (limit,))
        rows = []
        for row in cur.fetchall():
            rows.append({
                "source": row["source"],
                "name": row["name"],
                "desc": row["desc"],
                "price": row["price"],
                "link": row["link"],
                "categories": [c for c in (row["categories"] or "").split(",") if c],
                "price_num": row["price_num"],
                "first_seen": row["first_seen"],
                "last_seen": row["last_seen"],
            })
        return rows


def stats():
    with _LOCK, _connect() as conn:
        total = conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
        runs = conn.execute("SELECT COUNT(*) FROM runs").fetchone()[0]
        last = conn.execute("SELECT MAX(last_seen) FROM orders").fetchone()[0]
        return {"total": total, "runs": runs, "last_seen": last}


def clear():
    with _LOCK, _connect() as conn:
        conn.execute("DELETE FROM orders")
        conn.execute("DELETE FROM runs")
