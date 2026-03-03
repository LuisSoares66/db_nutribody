# app/product_catalog.py
import os
import re
import pandas as pd

import sys, os

def _base_dir():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

def _data_dir():
    data_dir = os.path.join(_base_dir(), "data")
    os.makedirs(data_dir, exist_ok=True)
    return data_dir

def catalog_path() -> str:
    return os.path.join(_data_dir(), "produtos.xlsx")

def _norm_col(c: str) -> str:
    # normaliza cabeçalho: remove espaços extras, acentos básicos e caracteres chatos
    s = str(c).strip().lower()
    s = re.sub(r"\s+", " ", s)  # colapsa espaços
    s = s.replace(" ", "_")
    s = s.replace(".", "")
    s = s.replace("(", "").replace(")", "")
    s = s.replace("/", "_")
    # remove caracteres não alfanuméricos/underscore
    s = re.sub(r"[^a-z0-9_]", "", s)
    return s

def _rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_norm_col(c) for c in df.columns]

    # alguns Excels acabam virando "vitb12_mcg" ou "vitb12mcg"
    # e "ptn_g" ou "ptng"
    rename_map = {
        "produto": "produto",
        "embalagem": "embalagem",
        "referencia": "referencia",
        "kcal": "kcal",

        "ptn_g": "ptn",
        "ptng": "ptn",
        "ptn": "ptn",

        "lip_g": "lip",
        "lipg": "lip",
        "lip": "lip",

        "fibras_g": "fibras",
        "fibrasg": "fibras",
        "fibras": "fibras",

        "sodio_mg": "sodio",
        "sodiomg": "sodio",
        "sodio": "sodio",

        "ferro_mg": "ferro",
        "ferromg": "ferro",
        "ferro": "ferro",

        "potassio_mg": "potassio",
        "potassiomg": "potassio",
        "potassio": "potassio",

        "vitb12_mcg": "vit_b12",
        "vitb12mcg": "vit_b12",
        "vit_b12": "vit_b12",

        "gordura_saturada_g": "gordura_saturada",
        "gordurasaturada_g": "gordura_saturada",
        "gordura_saturada": "gordura_saturada",
    }

    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # se por algum motivo a coluna produto não ficou exatamente "produto", tenta achar similar
    if "produto" not in df.columns:
        for c in df.columns:
            if c.startswith("produto"):
                df = df.rename(columns={c: "produto"})
                break

    return df

def load_produtos_catalog() -> pd.DataFrame:
    path = catalog_path()
    if not os.path.exists(path):
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    # Lê todas as abas
    xls = pd.ExcelFile(path)
    frames = []

    for sheet in xls.sheet_names:
        df_raw = pd.read_excel(path, sheet_name=sheet)
        if df_raw is None or df_raw.empty:
            continue

        df = _rename_columns(df_raw)

        # precisa ter coluna produto
        if "produto" not in df.columns:
            continue

        df["fabricante"] = str(sheet).strip()

        cols_final = [
            "fabricante", "produto", "embalagem", "referencia", "kcal",
            "ptn", "lip", "fibras", "sodio", "ferro", "potassio",
            "vit_b12", "gordura_saturada"
        ]
        for c in cols_final:
            if c not in df.columns:
                df[c] = ""

        # limpa
        for c in cols_final:
            df[c] = df[c].astype(str).fillna("").str.strip()

        df = df[df["produto"] != ""]
        frames.append(df[cols_final])

    if not frames:
        return pd.DataFrame(columns=[
            "fabricante","produto","embalagem","referencia","kcal",
            "ptn","lip","fibras","sodio","ferro","potassio","vit_b12","gordura_saturada"
        ])

    out = pd.concat(frames, ignore_index=True)
    out = out.drop_duplicates(subset=["fabricante", "produto"], keep="first")
    out = out.sort_values(["fabricante", "produto"], na_position="last")
    return out