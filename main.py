import os
import base64
import xmlrpc.client
import pandas as pd
import requests

# Odoo (Railway > Variables)
ODOO_URL = os.environ["ODOO_URL"].rstrip("/")
ODOO_DB = os.environ["ODOO_DB"]
ODOO_USER = os.environ["ODOO_USER"]
ODOO_API_KEY = os.environ["ODOO_API_KEY"]

# Resend (Railway > Variables)
RESEND_API_KEY = os.environ["RESEND_API_KEY"]
MAIL_FROM = os.environ["MAIL_FROM"]   # ex: export@genco-ie.com
MAIL_TO = os.environ["MAIL_TO"]       # ex: gencoit@genco-ie.com

# Salvar no Volume (persistente)
OUT_PATH = os.getenv("OUT_PATH", "/data/gnc_export.xlsx")

MODEL = "x_gnc"
FIELDS_MAP = {
    "x_name": "GNC",
    "x_studio_eta": "ETA",
    "x_studio_vessel_date": "ETD",
    "x_studio_ata": "ATA",
    "x_studio_atd": "ATD"
    "x_studio_status": "Status",
}

def send_email_resend(filepath: str):
    with open(filepath, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode("utf-8")

    payload = {
        "from": MAIL_FROM,
        "to": [MAIL_TO],
        "subject": "Export Odoo - x_gnc",
        "text": "Segue a planilha em anexo.",
        "attachments": [
            {"filename": os.path.basename(filepath), "content": content_b64}
        ],
    }

    r = requests.post(
        "https://api.resend.com/emails",
        headers={"Authorization": f"Bearer {RESEND_API_KEY}", "Content-Type": "application/json"},
        json=payload,
        timeout=30,
    )
    r.raise_for_status()
    print("OK: email enviado")

def main():
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_API_KEY, {})
    if not uid:
        raise RuntimeError("Falha no login do Odoo (DB/USER/API_KEY).")

    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

    rows = models.execute_kw(
        ODOO_DB, uid, ODOO_API_KEY,
        MODEL, "search_read",
        [[]],
        {"fields": list(FIELDS_MAP.keys()), "limit": 0},
    )

    df = pd.DataFrame(rows).rename(columns=FIELDS_MAP)
    df = df[["GNC", "ETA", "ETD", "ATA", "ATD", "Status"]]

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    df.to_excel(OUT_PATH, index=False)
    print(f"OK: gerado {OUT_PATH} com {len(df)} linhas")

    send_email_resend(OUT_PATH)

if __name__ == "__main__":
    main()
