import os
import psycopg2
from psycopg2.extras import RealDictCursor

def _db_url() -> str:
    url = (os.environ.get("DATABASE_URL") or "").strip()
    if not url:
        raise RuntimeError("DATABASE_URL não definida. Crie um .env com DATABASE_URL=...")

    # garante sslmode=require
    if "sslmode=" not in url:
        url += ("&" if "?" in url else "?") + "sslmode=require"

    return url

def get_conn():
    return psycopg2.connect(
        _db_url(),
        cursor_factory=RealDictCursor,
        connect_timeout=10,
    )

def fetch_all(sql: str, params=None):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params or ())
            return cur.fetchall()

def execute(sql: str, params=None):
    with get_conn() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params or ())
        conn.commit()