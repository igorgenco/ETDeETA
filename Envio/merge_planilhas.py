from pathlib import Path
import pandas as pd
import ast
import re

BASE = Path(r"C:\GencoServer\Genco IT\Planilhas para testes")
DOWN = BASE / "downloads"
OUT  = BASE / "output"

ODOO_FILE = DOWN / "gnc_export.xlsx"
APP_FILE  = DOWN / "docs.xlsx"

def normalize_gnc(x):
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = re.sub(r"^GNC", "", s)
    s = re.sub(r"\D", "", s)
    return s.lstrip("0")

def filled(x):
    if x is None:
        return False
    s = str(x).strip()
    if s == "" or s.upper() == "FALSE":
        return False
    return True

def extract_status_name(v):
    s = (v or "").strip()
    if not s:
        return ""
    try:
        parsed = ast.literal_eval(s)
        if isinstance(parsed, (list, tuple)) and len(parsed) >= 2:
            return str(parsed[1]).strip()
    except Exception:
        pass
    return s

MAP_STATUS = {
    "Waiting for Deposit": "Aguardando depósito",
    "Details": "Produção",
    "Production": "Produção",
    "Soon": "Produção",
    "Ready for Shipping": "Produção",
    "Loaded": "Carregado",
    "Shipped": "Carregado",
    "Done": "Concluído",
    "Archived": "Concluído",
}

def map_status_odoo_to_app(v):
    name = extract_status_name(v)
    return MAP_STATUS.get(name, "Interno")

def merge():
    df_odoo = pd.read_excel(ODOO_FILE, dtype=str)
    df_app  = pd.read_excel(APP_FILE, dtype=str)

    df_odoo.columns = [str(c).strip() for c in df_odoo.columns]
    df_app.columns  = [str(c).strip() for c in df_app.columns]

    df_odoo["GNC_norm"] = df_odoo["GNC"].map(normalize_gnc)
    df_app["gnc_norm"]  = df_app["gnc"].map(normalize_gnc)

    src = df_odoo[["GNC_norm", "ETA", "ETD", "ATA", "ATD", "Status", "Lead Time"]].copy()
    src = src[
        src["ETA"].map(filled) |
        src["ETD"].map(filled) |
        src["ATA"].map(filled) |
        src["ATD"].map(filled) |
        src["Status"].map(filled) |
        src["Lead Time"].map(filled)
    ]

    src = src.drop_duplicates(subset=["GNC_norm"], keep="last").set_index("GNC_norm")

    updated = 0
    for i, row in df_app.iterrows():
        g = row["gnc_norm"]
        if g in src.index:
            eta = src.loc[g, "ETA"]
            etd = src.loc[g, "ETD"]
            ata = src.loc[g, "ATA"]
            atd = src.loc[g, "ATD"]
            st  = src.loc[g, "Status"]
            lt  = src.loc[g, "Lead Time"]

            changed = False
            if filled(eta): df_app.at[i, "eta"] = eta; changed = True
            if filled(etd): df_app.at[i, "etd"] = etd; changed = True
            if filled(ata): df_app.at[i, "ata"] = ata; changed = True
            if filled(atd): df_app.at[i, "atd"] = atd; changed = True

            # Lead Time (Odoo) -> previsao (docs_upload.xlsx)
            if filled(lt): df_app.at[i, "previsao"] = lt; changed = True

            if filled(st):
                df_app.at[i, "status"] = map_status_odoo_to_app(st)
                changed = True

            if changed:
                updated += 1

    df_app = df_app.drop(columns=["gnc_norm"], errors="ignore")

    # manter só as colunas que sobem pro app (e na ordem)
    COLS_UPLOAD = ["id", "gnc", "previsao", "etd", "eta", "atd", "ata", "status", "comments", "archived"]
    for c in COLS_UPLOAD:
        if c not in df_app.columns:
            df_app[c] = ""
    df_app = df_app[COLS_UPLOAD]

    OUT.mkdir(parents=True, exist_ok=True)
    out_file = OUT / "docs_upload.xlsx"
    out_file.unlink(missing_ok=True)
    df_app.to_excel(out_file, index=False)

    print("OK. Linhas atualizadas:", updated)
    print("Arquivo gerado:", out_file)

if __name__ == "__main__":
    merge()