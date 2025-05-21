from tqdm import tqdm
import asyncio
import pandas as pd
from playwright.async_api import async_playwright
import os
from dotenv import load_dotenv
import logging


load_dotenv()

EXCEL_FILE = 'empresas.xlsx'
ODOO_URL = 'https://tributojusto.odoo.com'
USERNAME = os.getenv("ODOO_EMAIL") 
PASSWORD = os.getenv("ODOO_SENHA")


logging.basicConfig(
    filename='automacao_empresas.log',
    filemode='a',
    format='%(asctime)s - %(levelname)s - %(message)s',
    level=logging.INFO
)


async def login(page):
    await page.goto(f"{ODOO_URL}/web/login")
    await page.fill('input[name="login"]', USERNAME)
    await page.fill('input[name="password"]', PASSWORD)
    await page.click('button[type="submit"]')
    await page.wait_for_timeout(3000)

    try:
        await page.wait_for_selector('#result_app_6 > div:nth-child(2)', state='visible', timeout=40000)
        logging.info("✅ Login realizado com sucesso!")
    except Exception as e:
        logging.info("❌ Falha no login. Verifique:")
        logging.info(f"1. Credenciais estão corretas?")
        logging.info(f"2. O elemento #result_app_6 existe na página?")
        logging.info(f"Erro detalhado: {str(e)}")
        raise


async def process_company(page, company_name, index, total):
    try:
        progresso = f"[{index + 1}/{total}]"
        logging.info(f"\n🔍 {progresso} Processando: {company_name}")

        await page.goto(f"{ODOO_URL}/odoo/contacts?debug=1&view_type=list", wait_until='domcontentloaded')
        await page.wait_for_timeout(2000)

        search_input = page.locator('.o_searchview_input')
        await search_input.wait_for(state='visible', timeout=40000)
        await search_input.click()
        await search_input.fill(company_name)
        await page.wait_for_timeout(1000)
        await page.keyboard.press('Enter')
        await page.wait_for_timeout(3000)

        await page.locator('th.o_list_record_selector input[type="checkbox"]').click()

        await page.locator(
            'body div.o_action_manager div.o_control_panel_actions div.o_cp_action_menus button'
        ).click(timeout=40000)
        await page.wait_for_timeout(1500)

        await page.locator('span.o-dropdown-item:has-text("Mesclar")').click(timeout=5000)
        await page.wait_for_timeout(1500)

        await page.get_by_role("button", name="Mesclar contatos").click(timeout=10000)

        await page.locator(
            'h2:has-text("Não há mais contatos para mesclar para esta solicitação")'
        ).wait_for(state="visible", timeout=40000)

        await page.get_by_role("button", name="Fechar").click(timeout=5000)

        await search_input.click()
        await page.keyboard.press('Backspace')
        await page.wait_for_timeout(2500)

        percent = int(((index + 1) / total) * 100)
        print(f"✅ Concluído: {company_name} — Progresso: {percent}%")

    except Exception as e:
        logging.info(f"❌ Erro ao processar {company_name}: {str(e)}")
        await page.screenshot(path=f'error_{company_name[:10]}.png')  # Debug visual


async def main():
    df = pd.read_excel(EXCEL_FILE)
    companies = df['NomeEmpresa'].tolist()
    total = len(companies)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context()
        page = await context.new_page()

        await login(page)

        # Usando tqdm para barra de progresso
        for index, company in enumerate(tqdm(companies, desc="🔄 Processando empresas", unit="empresa")):
            await process_company(page, company, index, total)
            await page.wait_for_timeout(2000)

        await browser.close()
        logging.info("\n🏁 Processamento finalizado para todas as empresas!")

if __name__ == '__main__':
    asyncio.run(main())
