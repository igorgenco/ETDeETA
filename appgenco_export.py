```python
import os
import re
import base64
import requests

# ====== ENV (Railway > Variables) ======
APPGENCO_URL = os.environ["APPGENCO_URL"].rstrip("/")
APPGENCO_USER = os.environ["APPGENCO_USER"]
APPGENCO_PASS = os.environ["APPGENCO_PASS"]

RESEND_API_KEY = os.environ["RESEND_API_KEY"]
MAIL_FROM = os.environ["MAIL_FROM"]
MAIL_TO = os.environ["MAIL_TO"]

OUT_PATH = os.getenv("OUT_PATH", "/data/docs.xlsx")

LOGIN_PATH = "/admin/login/?next=/admin/orders/order/"
ORDERS_PATH = "/admin/orders/order/"

# Filtro do admin: ?q=&archived=all
FILTER_PARAMS = {"q": "", "archived": "all"}


def send_email_resend(filepath: str) -> None:
    with open(filepath, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode("utf-8")

    payload = {
        "from": MAIL_FROM,
        "to": [MAIL_TO],
        "subject": "Export - Orders (XLSX)",
        "text": "Segue a planilha em anexo.",
        "attachments": [
            {"filename": os.path.basename(filepath), "content": content_b64}
        ],
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


def extract_csrf(html: str, session: requests.Session) -> str:
    m = re.search(r'name="csrfmiddlewaretoken"\s+value="([^"]+)"', html)
    if m:
        return m.group(1)

    cookie_token = session.cookies.get("csrftoken")
    if cookie_token:
        return cookie_token

    raise RuntimeError("Não achei csrfmiddlewaretoken (nem no HTML, nem no cookie csrftoken).")


def extract_one_selected_action_id(html: str) -> str:
    m = re.search(r'name="_selected_action"\s+value="(\d+)"', html)
    if not m:
        raise RuntimeError(
            "Não achei nenhum _selected_action na página. "
            "Verifique se a lista tem pelo menos 1 order visível."
        )
    return m.group(1)


def looks_like_login_page(html: str) -> bool:
    return ('name="username"' in html) and ('name="password"' in html)


def is_xlsx_response(r: requests.Response) -> bool:
    ctype = (r.headers.get("Content-Type") or "").lower()
    cdisp = (r.headers.get("Content-Disposition") or "").lower()

    if "spreadsheetml" in ctype:
        return True
    if "attachment" in cdisp and ".xlsx" in cdisp:
        return True
    return False


def main():
    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)

    s = requests.Session()

    login_url = f"{APPGENCO_URL}{LOGIN_PATH}"
    orders_url = f"{APPGENCO_URL}{ORDERS_PATH}"

    # 1) GET login (pegar CSRF + cookie)
    r = s.get(login_url, timeout=30)
    r.raise_for_status()
    csrf_login = extract_csrf(r.text, s)

    # 2) POST login
    r = s.post(
        login_url,
        data={
            "csrfmiddlewaretoken": csrf_login,
            "username": APPGENCO_USER,
            "password": APPGENCO_PASS,
            "next": ORDERS_PATH,
        },
        headers={"Referer": login_url, "X-CSRFToken": csrf_login},
        timeout=30,
        allow_redirects=True,
    )
    r.raise_for_status()

    # 3) GET orders page COM FILTRO (?q=&archived=all)
    r = s.get(orders_url, params=FILTER_PARAMS, timeout=30)
    r.raise_for_status()

    if looks_like_login_page(r.text):
        raise RuntimeError("Login não foi mantido (voltou para a página de login). Verifique usuário/senha.")

    csrf_orders = extract_csrf(r.text, s)
    one_id = extract_one_selected_action_id(r.text)

    # 4) POST export no MESMO filtro
    r = s.post(
        orders_url,
        params=FILTER_PARAMS,
        data={
            "csrfmiddlewaretoken": csrf_orders,
            "action": "export_order",
            "select_across": "1",
            "index": "0",
            "_selected_action": one_id,
        },
        headers={"Referer": str(r.url), "X-CSRFToken": csrf_orders},
        timeout=120,
        allow_redirects=True,
    )
    r.raise_for_status()

    if not is_xlsx_response(r):
        ctype = r.headers.get("Content-Type")
        snippet = (r.text or "")[:300].replace("\n", " ")
        raise RuntimeError(f"Resposta não é XLSX. Content-Type: {ctype}. Trecho: {snippet}")

    # 5) salva arquivo
    with open(OUT_PATH, "wb") as f:
        f.write(r.content)

    print(f"OK: baixado {OUT_PATH} ({len(r.content)} bytes)")

    # 6) envia por email
    send_email_resend(OUT_PATH)


if __name__ == "__main__":
    main()
```
