print("SCRIPT INICIOU", flush=True)

from pathlib import Path
import time
import datetime as dt
import win32com.client

BASE = Path(r"C:\GencoServer\Genco IT\Planilhas para testes")
DOWN = BASE / "downloads"
DOWN.mkdir(parents=True, exist_ok=True)

ACCOUNT = "gencoit@genco-ie.com"

ODOO_SUBJECT = "Export Odoo - x_gnc"
APP_SUBJECT  = "Export - Orders (XLSX)"

ODOO_SAVE_AS = "gnc_export.xlsx"
APP_SAVE_AS  = "docs.xlsx"

def get_store(outlook, smtp):
    for store in outlook.Folders:
        if str(store.Name).lower() == smtp.lower():
            return store
    return None

def get_inbox_folder(store):
    for name in ["Inbox", "Caixa de Entrada"]:
        try:
            return store.Folders[name]
        except Exception:
            pass
    raise RuntimeError("Não achei Inbox/Caixa de Entrada nessa conta.")

def save_latest_xlsx_today_from_inbox(store, subject_contains, save_as):
    inbox = get_inbox_folder(store)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    today = dt.date.today()

    for mail in items:
        try:
            if mail.Class != 43:
                continue

            received = mail.ReceivedTime
            if hasattr(received, "date"):
                if received.date() != today:
                    continue

            subj = str(mail.Subject or "")
            if subject_contains.lower() in subj.lower():
                for att in mail.Attachments:
                    name = str(att.FileName or "")
                    if name.lower().endswith(".xlsx"):
                        out_path = DOWN / save_as
                        out_path.unlink(missing_ok=True)   # sobrescreve
                        att.SaveAsFile(str(out_path))
                        print("Baixado:", save_as, "| Assunto:", subj, flush=True)
                        return True
        except:
            pass

    return False

def main():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    store = get_store(outlook, ACCOUNT)
    if not store:
        print("Não achei a conta:", ACCOUNT, flush=True)
        raise SystemExit(1)

    for attempt in range(1, 11):
        ok1 = save_latest_xlsx_today_from_inbox(store, ODOO_SUBJECT, ODOO_SAVE_AS)
        ok2 = save_latest_xlsx_today_from_inbox(store, APP_SUBJECT,  APP_SAVE_AS)

        print(f"Tentativa {attempt}/10 -> Odoo:{ok1} App:{ok2}", flush=True)

        if ok1 and ok2:
            print("Pasta:", DOWN, flush=True)
            return

        time.sleep(60)

    print("Falhou: não encontrou e-mails de HOJE com os anexos.", flush=True)
    raise SystemExit(1)

if __name__ == "__main__":
    main()