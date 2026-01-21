import os
import xmlrpc.client
import pandas as pd

# Variáveis de ambiente (Railway > Variables)
ODOO_URL = os.environ["ODOO_URL"].rstrip("/")
ODOO_DB = os.environ["ODOO_DB"]
ODOO_USER = os.environ["ODOO_USER"]
ODOO_API_KEY = os.environ["ODOO_API_KEY"]

# Salvar no Volume do Railway (montado em /data)
OUT_PATH = os.getenv("OUT_PATH", "/data/gnc_export.xlsx")

MODEL = "x_gnc"

# campos técnicos -> nomes bonitos no Excel
FIELDS_MAP = {
    "x_name": "GNC",
    "x_studio_eta": "ETA",
    "x_studio_vessel_date": "ETD",
    "x_studio_status": "Status",
}

def main():
    # login
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_API_KEY, {})
    if not uid:
        raise RuntimeError("Falha no login do Odoo (DB/USER/API_KEY).")

    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

    # filtros opcionais (deixa vazio pra puxar tudo)
    domain = []
    fields = list(FIELDS_MAP.keys())

    rows = models.execute_kw(
        ODOO_DB, uid, ODOO_API_KEY,
        MODEL, "search_read",
        [domain],
        {"fields": fields, "limit": 0}
    )

    df = pd.DataFrame(rows)

    # renomear colunas para o Excel
    df = df.rename(columns=FIELDS_MAP)

    # manter só as colunas desejadas e na ordem certa
    df = df[["GNC", "ETA", "ETD", "Status"]]

    # garantir que a pasta do arquivo existe
    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)

    df.to_excel(OUT_PATH, index=False)
    print(f"OK: gerado {OUT_PATH} com {len(df)} linhas")

if __name__ == "__main__":
    main()
