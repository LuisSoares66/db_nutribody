import pandas as pd
from datetime import date

def _fmt_date(d: date | None) -> str:
    """Converte date para dd/mm/aaaa (string)."""
    if not d:
        return ""
    return d.strftime("%d/%m/%Y")

def exportar_hospitais_para_excel(hospitais, caminho_xlsx: str):
    """
    hospitais: lista de objetos Hospital (SQLAlchemy)
    caminho_xlsx: caminho do arquivo .xlsx (ex: data/hospitais.xlsx)
    """
    linhas = []
    for h in hospitais:
        linhas.append({
            "id_hospital": h.id,
            "nome_hospital": getattr(h, "nome_hospital", "") or "",
            "cidade": getattr(h, "cidade", "") or "",
            "estado": getattr(h, "estado", "") or "",

            # ✅ NOVAS COLUNAS
            "data_visita": _fmt_date(getattr(h, "data_visita", None)),
            "data_retorno": _fmt_date(getattr(h, "data_retorno", None)),
        })

    df = pd.DataFrame(linhas)

    # garante ordem das colunas (ajuste conforme seu Excel real)
    colunas = [
        "id_hospital", "nome_hospital", "cidade", "estado",
        "data_visita", "data_retorno"
    ]
    for c in colunas:
        if c not in df.columns:
            df[c] = ""
    df = df[colunas]

    with pd.ExcelWriter(caminho_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="hospitais")