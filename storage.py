# -*- coding: utf-8 -*-
"""
Постоянное хранилище найденных заказов.

Бэкенд выбирается автоматически:
    DATABASE_URL задан  -> PostgreSQL (Render/Railway и др.);
    иначе                -> SQLite (/data/orders.db на Amvera, ./data/orders.db локально).

Публичный API одинаков для обоих бэкендов:
    init(), save_run(rows, meta), known_keys(),
    history(limit, query), stats(), clear()
"""
import os
import threading
from datetime import datetime

_LOCK = threading.Lock()

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

# Render/Heroku иногда отдают "postgres://" — psycopg2 понимает только "postgresql://"
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = "postgresql://" + DATABASE_URL[len("postgres://"):]

USE_PG = bool(DATABASE_URL)


def backend_name():
    return "postgresql" if USE_PG else "sqlite"


# ---------------------------------------------------------------------------
# SQLite
# ---------------------------------------------------------------------------
def _sqlite_path():
    base = "/data" if os.path.isdir("/data") and os.access("/data", os.W_OK) else "data"
    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "orders.db")


def _connect_sqlite():
    import sqlite3
    conn = sqlite3.connect(_sqlite_path(), timeout=10)
    conn.row_factory = sqlite3.Row
    return conn


# ---------------------------------------------------------------------------
# Postgres
# ---------------------------------------------------------------------------
def _connect_pg():
    import psycopg2
    import psycopg2.extras
    conn = psycopg2.connect(DATABASE_URL)
    return conn


def _dict_cursor_pg(conn):
    import psycopg2.extras
    return conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)


def _connect():
    return _connect_pg() if USE_PG else _connect_sqlite()


def _q(sql):
    """Приводит плейсхолдеры к нужному стилю: %s для Postgres, ? для SQLite."""
    return sql.replace("?", "%s") if USE_PG else sql


def _sql_upsert():
    if USE_PG:
        return (
            "INSERT INTO orders (key, source, name, \"desc\", price, link,"
            " categories, price_num, first_seen, last_seen)"
            " VALUES (?,?,?,?,?,?,?,?,?,?)"
            " ON CONFLICT (key) DO UPDATE SET last_seen = EXCLUDED.last_seen"
        )
    return (
        "INSERT INTO orders (key, source, name, desc, price, link,"
        " categories, price_num, first_seen, last_seen)"
        " VALUES (?,?,?,?,?,?,?,?,?,?)"
        " ON CONFLICT (key) DO UPDATE SET last_seen = excluded.last_seen"
    )


def _sql_run_insert():
    if USE_PG:
        return ("INSERT INTO runs (ts, sources, query, pages, found, new)"
                " VALUES (?,?,?,?,?,?) RETURNING id")
    return ("INSERT INTO runs (ts, sources, query, pages, found, new)"
            " VALUES (?,?,?,?,?,?)")


# ---------------------------------------------------------------------------
# Схема
# ---------------------------------------------------------------------------
def init():
    """Создаёт таблицы, если их ещё нет."""
    with _LOCK:
        conn = _connect()
        try:
            cur = conn.cursor()
            if USE_PG:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS orders (
                        key         TEXT PRIMARY KEY,
                        source      TEXT,
                        name        TEXT,
                        "desc"      TEXT,
                        price       TEXT,
                        link        TEXT,
                        categories  TEXT,
                        price_num   BIGINT,
                        first_seen  TEXT,
                        last_seen   TEXT
                    )
                    """
                )
                cur.execute(
                    "CREATE TABLE IF NOT EXISTS runs ("
                    " id BIGSERIAL PRIMARY KEY,"
                    " ts TEXT, sources TEXT, query TEXT, pages TEXT,"
                    " found INTEGER, new INTEGER)"
                )
                cur.execute(
                    "CREATE INDEX IF NOT EXISTS idx_last_seen ON orders(last_seen)")
            else:
                cur.execute(
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
                cur.execute(
                    "CREATE TABLE IF NOT EXISTS runs ("
                    " id INTEGER PRIMARY KEY AUTOINCREMENT,"
                    " ts TEXT, sources TEXT, query TEXT, pages TEXT,"
                    " found INTEGER, new INTEGER)"
                )
                cur.execute(
                    "CREATE INDEX IF NOT EXISTS idx_last_seen ON orders(last_seen)")
            conn.commit()
        finally:
            conn.close()


def _row_key(r):
    return r.get("link") or f"{r.get('source','')}:{r.get('name','')}"


def save_run(rows, meta):
    """
    Сохраняет заказы. Возвращает список ключей, встретившихся ВПЕРВЫЕ
    (для подсветки «новых» заказов).
    """
    now = datetime.utcnow().isoformat(timespec="seconds")
    new_keys = []
    with _LOCK:
        conn = _connect()
        try:
            cur = conn.cursor()
            upsert = _q(_sql_upsert())
            for r in rows:
                key = _row_key(r)
                if not key:
                    continue
                cats = ",".join(r.get("categories", []))
                cur.execute(_q("SELECT 1 FROM orders WHERE key=?"), (key,))
                exists = cur.fetchone() is not None
                if not exists:
                    new_keys.append(key)
                cur.execute(upsert, (
                    key, r.get("source", ""), r.get("name", ""), r.get("desc", ""),
                    r.get("price", ""), r.get("link", ""), cats,
                    r.get("price_num"), now, now,
                ))
            cur.execute(_q(_sql_run_insert()), (
                now, ",".join(meta.get("sources", [])), meta.get("query", ""),
                str(meta.get("pages", "")), len(rows), len(new_keys),
            ))
            conn.commit()
        finally:
            conn.close()
    return new_keys


def known_keys():
    """Множество всех ключей заказов, когда-либо сохранённых."""
    with _LOCK:
        conn = _connect()
        try:
            cur = conn.cursor()
            cur.execute("SELECT key FROM orders")
            return {row[0] for row in cur.fetchall()}
        finally:
            conn.close()


def history(limit=500, query=""):
    """Последние сохранённые заказы (для восстановления сессии)."""
    with _LOCK:
        conn = _connect()
        try:
            cur = _dict_cursor_pg(conn) if USE_PG else conn.cursor()
            desc_col = '"desc"' if USE_PG else "desc"
            if query:
                like = f"%{query}%"
                cur.execute(
                    _q(f"SELECT * FROM orders WHERE name LIKE ? OR {desc_col} LIKE ?"
                       " ORDER BY last_seen DESC LIMIT ?"),
                    (like, like, limit),
                )
            else:
                cur.execute(
                    _q("SELECT * FROM orders ORDER BY last_seen DESC LIMIT ?"),
                    (limit,),
                )
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
        finally:
            conn.close()


def stats():
    with _LOCK:
        conn = _connect()
        try:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM orders")
            total = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM runs")
            runs = cur.fetchone()[0]
            cur.execute("SELECT MAX(last_seen) FROM orders")
            last = cur.fetchone()[0]
            return {"total": total, "runs": runs,
                    "last_seen": last, "backend": backend_name()}
        finally:
            conn.close()


def clear():
    with _LOCK:
        conn = _connect()
        try:
            cur = conn.cursor()
            cur.execute("DELETE FROM orders")
            cur.execute("DELETE FROM runs")
            conn.commit()
        finally:
            conn.close()
