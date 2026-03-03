import os
import psycopg2
from psycopg2.extras import RealDictCursor

def get_database_url() -> str:
    url = os.environ.get("DATABASE_URL", "").strip()
    if not url:
        raise RuntimeError("DATABASE_URL não definida. Crie um .env na raiz com DATABASE_URL=...")

    # garante sslmode=require
    if "sslmode=" not in url:
        url += ("&" if "?" in url else "?") + "sslmode=require"

    return url

def get_conn():
    """
    Retorna conexão psycopg2 com SSL requerido (Render).
    Usa RealDictCursor para retornar dict (coluna->valor).
    """
    return psycopg2.connect(
        get_database_url(),
        cursor_factory=RealDictCursor,
        connect_timeout=10,
    )