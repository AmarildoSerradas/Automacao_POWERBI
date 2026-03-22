from playwright.sync_api import sync_playwright
import time
from pathlib import Path
from PIL import Image
from datetime import datetime, timedelta
import io
import win32clipboard

# ================= FUNÇÃO CLIPBOARD =================

def copiar_imagem_para_clipboard(caminho_img):
    img = Image.open(caminho_img).convert("RGB")

    output = io.BytesIO()
    img.save(output, "BMP")
    data = output.getvalue()[14:]  # remove header BMP
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

# ================= DATAS =================

hoje = datetime.today()
ontem = (hoje - timedelta(days=1)).strftime("%d/%m/%Y")
antes_de_ontem = (hoje - timedelta(days=2)).strftime("%d/%m/%Y")
ano = str(datetime.now().year)
meses = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
]


mes_atual = meses[datetime.now().month - 1]
# ================= LOJAS =================

lojas = ["LOJA 1", "LOJA 2", "LOJA 3", "LOJA 5", "LOJA 6", "LOJA 7"]

# ================= PASTAS =================

BASE_DIR = Path(__file__).resolve().parent
PASTA = BASE_DIR / "cfc"
PASTA2 = BASE_DIR / "P.A.R"
PASTA3 = BASE_DIR / "Água"

PASTA.mkdir(exist_ok=True)
PASTA2.mkdir(exist_ok=True)
PASTA3.mkdir(exist_ok=True)

# ================= PLAYWRIGHT =================

with sync_playwright() as pw:
    navegador = pw.chromium.launch_persistent_context(
        user_data_dir="Amarildomtfd",
        headless=False
    )

    # ================= POWER BI - CFC =================

    pagina = navegador.new_page()
    pagina.goto("https://app.powerbi.com/view?r=eyJrIjoiYzc5ZTJkOTctYTE2ZS00ZTcyLWI0YmUtZDgyZDQyOGFmODcxIiwidCI6ImQ3ZTVjZGNiLWE0YTYtNDcyZC1hZWFhLTU4NDZiYTZiYmVhZSJ9")
    time.sleep(2)
    try:
        pagina.get_by_test_id("thumbnail-image").screenshot(path=PASTA / "CAPA.png")
        pagina.get_by_test_id("thumbnail-image").click()
    except:
        pass
    time.sleep(2)

    pagina.get_by_role("region", name="Relatório do Power BI").screenshot(path=PASTA / "geral.png")

    for loja in lojas:
        pagina.get_by_role("option", name=loja).click()
        time.sleep(3)
        pagina.get_by_role("region", name="Relatório do Power BI").screenshot(path=PASTA / f"{loja}.png")

    # ================= CONVERSÃO CFC ================= 

    for caminho in PASTA.glob("*.*"):
        if caminho.suffix.lower() not in [".png", ".jpg", ".jpeg"]:
            continue

        with Image.open(caminho) as img:
            img.convert("RGB").save(caminho.with_suffix(".jpg"), "JPEG", quality=95)

        if caminho.suffix.lower() != ".jpg":
            caminho.unlink()

    # ================= WHATSAPP - CFC =================

    pagina2 = navegador.new_page()
    pagina2.goto("https://web.whatsapp.com/")
    time.sleep(10)

    pagina2.get_by_role("tab", name="Grupos").click()
    pagina2.get_by_role("paragraph").fill("CFC DE")
    pagina2.get_by_text("CFC DEV").click()
    time.sleep(3)

    imagens_cfc = sorted(PASTA.glob("*.jpg"))

    for img in imagens_cfc:
        copiar_imagem_para_clipboard(img)
        time.sleep(0.5)

        pagina2.keyboard.press("Control+V")
        time.sleep(1)

        pagina2.keyboard.press("Enter")
        time.sleep(1.5)

    # ================= POWER BI - PAR =================

    pagina3 = navegador.new_page()
    pagina3.goto("https://app.powerbi.com/view?r=eyJrIjoiZjQ4YjkyMDMtMjg1NC00MjUxLWE5YTEtNjNlODNkNjQwNGRlIiwidCI6ImQ3ZTVjZGNiLWE0YTYtNDcyZC1hZWFhLTU4NDZiYTZiYmVhZSJ9")
    time.sleep(3)
    try:
        pagina3.get_by_test_id("thumbnail-image").screenshot(path=PASTA2 / "CAPA.png")
        pagina3.get_by_test_id("thumbnail-image").click()
    except:
        pass
    time.sleep(3)

    pagina3.get_by_role("textbox", name="Data de término. Intervalo de").click()
    pagina3.keyboard.press("Control+A")
    pagina3.keyboard.press("Backspace")
    pagina3.get_by_role("textbox", name="Data de término. Intervalo de").fill(ontem)

    pagina3.get_by_role("textbox", name="Data de início. Intervalo de").click()
    pagina3.keyboard.press("Control+A")
    pagina3.keyboard.press("Backspace")
    pagina3.get_by_role("textbox", name="Data de início. Intervalo de").fill(antes_de_ontem)

    pagina3.get_by_text("PERDA").click()
    time.sleep(2)
    pagina3.get_by_text("PERDA").click()


    for loja in lojas:
        pagina3.get_by_title(loja).click()
        time.sleep(4)
        pagina3.get_by_role("region", name="Relatório do Power BI").screenshot(path=PASTA2 / f"{loja}.png")

    # ================= CONVERSÃO PAR =================

    for caminho in PASTA2.glob("*.*"):
        if caminho.suffix.lower() not in [".png", ".jpg", ".jpeg"]:
            continue

        with Image.open(caminho) as img:
            img.convert("RGB").save(caminho.with_suffix(".jpg"), "JPEG", quality=95)

        if caminho.suffix.lower() != ".jpg":
            caminho.unlink()

    # ================= WHATSAPP - PAR =================

    pagina2.goto("https://web.whatsapp.com/")
    time.sleep(5)

    pagina2.get_by_role("tab", name="Grupos").click()
    pagina2.get_by_role("paragraph").fill("PAR DE")
    pagina2.get_by_text("PAR DEV").click()
    time.sleep(3)

    imagens_par = sorted(PASTA2.glob("*.jpg"))

    for img in imagens_par:
        copiar_imagem_para_clipboard(img)
        time.sleep(0.5)

        pagina2.keyboard.press("Control+V")
        time.sleep(1)

        pagina2.keyboard.press("Enter")
        time.sleep(1.5)

    pagina4 = navegador.new_page()
    pagina4.goto("https://app.powerbi.com/view?r=eyJrIjoiOTVlNDMwMDEtNjFmMi00NWVkLTgyYzEtNmFhMzUwNzgwZGY0IiwidCI6ImQ3ZTVjZGNiLWE0YTYtNDcyZC1hZWFhLTU4NDZiYTZiYmVhZSJ9")
    time.sleep(2)
    try:
        pagina4.get_by_test_id("thumbnail-image").screenshot(path=PASTA / "CAPA.png")
        pagina4.get_by_test_id("thumbnail-image").click()
    except:
        pass

    pagina4.get_by_text(ano).click()
    pagina4.get_by_role("combobox", name="Mês").click()
    pagina4.get_by_text(mes_atual).click()
    pagina4.get_by_role("columnheader", name="Dia").click()
    time.sleep(2)
    pagina4.get_by_role("region", name="Relatório do Power BI").screenshot(path=PASTA3 / f"Loja.png")
    time.sleep(3)
    # ================= FINAL =================

    for lojaA in lojas:
        try:
            pagina4.get_by_role("combobox", name="LOJA").click()
            try:
                pagina4.get_by_text(lojaA).click()
                pagina4.get_by_text("LOJA").nth(1).click()
                time.sleep(2)
                pagina4.get_by_role("region", name="Relatório do Power BI").screenshot(path=PASTA3 / f"{lojaA}.png")
            except:
                pagina4.get_by_text("LOJA").nth(1).click()
                pass
            
        except:
            pass

    for caminho in PASTA3.glob("*.*"):
        if caminho.suffix.lower() not in [".png", ".jpg", ".jpeg"]:
            continue

        with Image.open(caminho) as img:
            img.convert("RGB").save(caminho.with_suffix(".jpg"), "JPEG", quality=95)

        if caminho.suffix.lower() != ".jpg":
            caminho.unlink()

    # ================= WHATSAPP - Água =================

    pagina2.goto("https://web.whatsapp.com/")
    time.sleep(5)

    pagina2.get_by_role("tab", name="Grupos").click()
    pagina2.get_by_role("paragraph").fill("PAR")
    pagina2.get_by_text("PAR DEV").click()
    time.sleep(3)

    imagens_água = sorted(PASTA3.glob("*.jpg"))

    for img in imagens_água:
        copiar_imagem_para_clipboard(img)
        time.sleep(0.5)

        pagina2.keyboard.press("Control+V")
        time.sleep(1)

        pagina2.keyboard.press("Enter")
        time.sleep(1.5)






    time.sleep(30)
