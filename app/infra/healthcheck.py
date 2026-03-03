from infra.db import fetch_all

def ping_db() -> str:
    row = fetch_all("SELECT version() AS v;")[0]
    return row["v"]