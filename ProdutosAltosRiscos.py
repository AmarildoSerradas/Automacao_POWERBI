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
    data = output.getvalue()[14:]
    output.close()

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

# ================= FUNÇÃO WHATSAPP =================

def enviar_whatsapp(pagina, nome_grupo, pasta):

    pagina.goto("https://web.whatsapp.com/")

    # Espera interface real carregar (não só rede)
    pagina.wait_for_selector("#side", timeout=60000)

    # DEBUG inicial
    pagina.screenshot(path="debug_inicio.png")

    # BUSCA (escopo correto: sidebar)
    busca = pagina.locator("#side").get_by_role("textbox")
    busca.click()
    busca.fill(nome_grupo)

    # Espera grupo aparecer de forma confiável
    try:
        pagina.wait_for_selector(f"span[title='{nome_grupo}']", timeout=10000)
        pagina.click(f"span[title='{nome_grupo}']")
    except:
        pagina.screenshot(path="erro_grupo.png")
        raise Exception(f"Grupo '{nome_grupo}' não encontrado")

    # CAIXA DE MENSAGEM (escopo correto: footer)
    caixa = pagina.locator("footer").get_by_role("textbox")
    caixa.wait_for(timeout=60000)

    # DEBUG após abrir chat
    pagina.screenshot(path="debug_chat.png")

    imagens = sorted(pasta.glob("*.jpg"))

    if not imagens:
        print(f"Nenhuma imagem encontrada em: {pasta}")
        return

    for img in imagens:
        try:
            copiar_imagem_para_clipboard(img)

            pagina.wait_for_timeout(1000)

            caixa.click()
            pagina.keyboard.press("Control+V")

            # Espera preview aparecer (evita enviar vazio)
            pagina.get_by_role("button", name="Enviar").click()

            pagina.wait_for_timeout(2000)

        except Exception as e:
            print(f"Erro ao enviar {img.name}: {e}")
            pagina.screenshot(path=f"erro_envio_{img.stem}.png")

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

def limpar(pasta):
    for arquivo in pasta.glob("*.jpg"):
        if arquivo.name != "CAPA.jpg":
            try:
                arquivo.unlink()
            except:
                pass

limpar(PASTA)
limpar(PASTA2)
limpar(PASTA3)

def converter(pasta):
    for caminho in pasta.glob("*.*"):
        if caminho.suffix.lower() not in [".png", ".jpg", ".jpeg"]:
            continue

        with Image.open(caminho) as img:
            img.convert("RGB").save(caminho.with_suffix(".jpg"), "JPEG", quality=95)

        if caminho.suffix.lower() != ".jpg":
            caminho.unlink()

# ================= PLAYWRIGHT =================

with sync_playwright() as pw:
    navegador = pw.chromium.launch_persistent_context(
        user_data_dir="Amarildomtfd",
        headless=False
    )

    # ================= POWER BI - PAR =================

    pagina3 = navegador.new_page()
    pagina3.goto("https://app.powerbi.com/view?r=eyJrIjoiZjQ4YjkyMDMtMjg1NC00MjUxLWE5YTEtNjNlODNkNjQwNGRlIiwidCI6ImQ3ZTVjZGNiLWE0YTYtNDcyZC1hZWFhLTU4NDZiYTZiYmVhZSJ9")
    pagina3.wait_for_timeout(5000)

    try:
        pagina3.get_by_test_id("thumbnail-image").click()
    except:
        pass
    
    pagina3.get_by_role("textbox", name="Data de término. Intervalo de").fill(ontem)
    pagina3.get_by_role("textbox", name="Data de início. Intervalo de").fill(antes_de_ontem)

    pagina3.get_by_text("PERDA").click()
    time.sleep(3)
    pagina3.locator("div:nth-child(7) > .powervisuals-glyph").click()

    pagina3.wait_for_timeout(2000)

    for loja in lojas:
        try:
            pagina3.get_by_title(loja).click()
            pagina3.wait_for_timeout(3000)
            pagina3.get_by_role("region", name="Relatório do Power BI").screenshot(
                path=PASTA2 / f"{loja}.png"
            )
        except:
            pass

    converter(PASTA2)

    # ================= POWER BI - ÁGUA =================

    pagina4 = navegador.new_page()
    pagina4.goto("https://app.powerbi.com/view?r=eyJrIjoiOTVlNDMwMDEtNjFmMi00NWVkLTgyYzEtNmFhMzUwNzgwZGY0IiwidCI6ImQ3ZTVjZGNiLWE0YTYtNDcyZC1hZWFhLTU4NDZiYTZiYmVhZSJ9")
    pagina4.wait_for_timeout(5000)

    try:
        pagina4.get_by_text(ano).click()
    except:
        pass

    try:
        pagina4.get_by_role("combobox", name="Mês").click()
        pagina4.get_by_text(mes_atual).click()
    except:
        pass

    pagina4.get_by_role("columnheader", name="Dia").click()
    pagina4.wait_for_timeout(2000)

    pagina4.get_by_role("region", name="Relatório do Power BI").screenshot(
        path=PASTA3 / "Loja.png"
    )

    for loja in lojas:
        try:
            pagina4.get_by_role("combobox", name="LOJA").click()
            pagina4.get_by_text(loja).click()
            pagina4.wait_for_timeout(2000)

            pagina4.get_by_role("region", name="Relatório do Power BI").screenshot(
                path=PASTA3 / f"{loja}.png"
            )
        except:
            pass

    converter(PASTA3)

    # ================= WHATSAPP =================

    pagina2 = navegador.new_page()

    enviar_whatsapp(pagina2, "Segurança P.A.R.", PASTA2)
    enviar_whatsapp(pagina2, "Varandas grupo lojas", PASTA3)
    # enviar_whatsapp(pagina2, "Frente de Caixa - Varandas", PASTA)

    time.sleep(10)