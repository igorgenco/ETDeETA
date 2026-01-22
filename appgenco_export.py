import os
import re
import base64
import requests

# Site (Railway > Variables)
APPGENCO_URL = os.environ["APPGENCO_URL"].rstrip("/")
APPGENCO_USER = os.environ["APPGENCO_USER"]
APPGENCO_PASS = os.environ["APPGENCO_PASS"]

# Resend (Railway > Variables)
RESEND_API_KEY = os.environ["RESEND_API_KEY"]
MAIL_FROM = os.environ["MAIL_FROM"]
MAIL_TO = os.environ["MAIL_TO"]

# Arquivo no Volume
OUT_PATH = os.getenv("OUT_PATH", "/data/docs.xlsx")

LOGIN_PATH = "/admin/login/?next=/admin/orders/order/"
EXPORT_URL = f"{APPGENCO_URL}/admin/orders/order/"  # POST aqui com action=export_order


def send_email_resend(filepath: str):
    with open(filepath, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode("utf-8")

    payload = {
        "from": MAIL_FROM,
        "to": [MAIL_TO],
        "subject": "Export - Orders (XLSX)",
        "text": "Segue a planilha em anexo.",
        "attachments": [{"filename": os.path.basename(filepath), "content": content_b64}],
    }

    r = requests.post(
        "https://api.resend.com/emails",
        headers={
            "Authorization": f"Bearer {RESEND_API_KEY}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )
    r.raise_for_status()
    print("OK: email enviado")


def get_csrf_from_cookie(session: requests.Session) -> str:
    token = session.cookies.get("csrftoken")
    if not token:
        raise RuntimeError("Não achei cookie csrftoken.")
    return token


def main():
    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    s = requests.Session()

    # 1) abre login pra pegar CSRF e cookie
    login_url = f"{APPGENCO_URL}{LOGIN_PATH}"
    r = s.get(login_url, timeout=30)
    r.raise_for_status()
    csrf = get_csrf_from_cookie(s)

    # 2) faz login (Django admin)
    r = s.post(
        login_url,
        data={
            "csrfmiddlewaretoken": csrf,
            "username": APPGENCO_USER,
            "password": APPGENCO_PASS,
            "next": "/admin/orders/order/",
        },
        headers={"Referer": login_url},
        timeout=30,
        allow_redirects=True,
    )
    r.raise_for_status()

    # 3) abre a página de orders pra pegar CSRF válido pra action
    r = s.get(EXPORT_URL, timeout=30)
    r.raise_for_status()
    csrf2 = extract_csrf(r.text)

    # 4) dispara export (select_across=1 = todos)
    r = s.post(
        EXPORT_URL,
        data={
            "csrfmiddlewaretoken": csrf2,
            "action": "export_order",
            "select_across": "1",
            "index": "0",
        },
        headers={"Referer": EXPORT_URL},
        timeout=120,
    )
    r.raise_for_status()

    # 5) salva arquivo (confere se veio XLSX)
    ctype = (r.headers.get("Content-Type") or "").lower()
    if "spreadsheetml" not in ctype and "octet-stream" not in ctype:
        # provavelmente voltou HTML (erro/permissão)
        raise RuntimeError(f"Resposta não é XLSX. Content-Type: {r.headers.get('Content-Type')}")

    with open(OUT_PATH, "wb") as f:
        f.write(r.content)

    print(f"OK: baixado {OUT_PATH} ({len(r.content)} bytes)")

    # 6) envia por email
    send_email_resend(OUT_PATH)


if __name__ == "__main__":
    main()
