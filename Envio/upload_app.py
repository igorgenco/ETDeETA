from pathlib import Path
from playwright.sync_api import sync_playwright

APP_URL = "http://genco.saltambiental.com.br/admin/orders/order/"

APP_USER = "Daniel_Marques"
APP_PASS = "Genco2025"

FILE_TO_UPLOAD = Path(__file__).resolve().parent / "output" / "docs_upload.xlsx"

SEL_USER = "input#id_username"
SEL_PASS = "input#id_password"
SEL_LOGIN_BTN = "input[type='submit'][value='Log in']"

SEL_FILE_INPUT = "input[type='file']"
SEL_UPLOAD_BTN = "input[type='submit'][value='upload']"

def main():
    if not FILE_TO_UPLOAD.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {FILE_TO_UPLOAD}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        page.goto(APP_URL)

        # se cair no login, faz login
        if page.locator(SEL_USER).count() > 0:
            page.fill(SEL_USER, APP_USER)
            page.fill(SEL_PASS, APP_PASS)
            page.click(SEL_LOGIN_BTN)
            page.wait_for_load_state("networkidle")

        # upload
        page.set_input_files(SEL_FILE_INPUT, str(FILE_TO_UPLOAD))
        page.click(SEL_UPLOAD_BTN)

        page.wait_for_timeout(8000)
        browser.close()

if __name__ == "__main__":
    main()
