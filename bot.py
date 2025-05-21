import asyncio
import pandas as pd
from playwright.async_api import async_playwright
import os
from dotenv import load_dotenv

load_dotenv()

EXCEL_FILE = 'empresas.xlsx'
ODOO_URL = 'https://teste-tj.odoo.com'
USERNAME = os.getenv("ODOO_EMAIL") 
PASSWORD = os.getenv("ODOO_SENHA")

async def login(page):
    await page.goto(f"{ODOO_URL}/web/login")
    await page.fill('input[name="login"]', USERNAME)
    await page.fill('input[name="password"]', PASSWORD)
    await page.click('button[type="submit"]')
    await page.wait_for_timeout(3000)
    

    try:
        # Aguarda o elemento específico da dashboard
        await page.wait_for_selector('#result_app_6 > div:nth-child(2)', state='visible', timeout=40000)
        print("Login realizado com sucesso!")
    except Exception as e:
        print("Falha no login. Verifique:")
        print(f"1. Credenciais estão corretas?")
        print(f"2. O elemento #result_app_6 existe na página?")
        print(f"Erro detalhado: {str(e)}")
        raise

async def process_company(page, company_name):
    try:
        # Acessar URL de contatos
        await page.goto('https://teste-tj.odoo.com/odoo/contacts?debug=1&view_type=list', wait_until='domcontentloaded')
        await page.wait_for_timeout(2000) 


        search_input = page.locator('.o_searchview_input')
        await search_input.wait_for(state='visible', timeout=40000)
        await search_input.click()
        
        # Preencher nome da empresa
        await search_input.fill(company_name)
        await page.wait_for_timeout(1000)  # Delay para digitação
        await page.keyboard.press('Enter')
        
        # Restante do fluxo com timeouts ajustados
        await page.wait_for_timeout(3000)  # Aguardar resultados
        
        # Selecionar todos
        #await page.pause()
        await page.locator('th.o_list_record_selector input[type="checkbox"]').click()

        
        # Abrir menu de ações
        await page.locator('body > div.o_action_manager > div > div > div > div.o_control_panel_actions.d-empty-none.d-flex.align-items-center.justify-content-start.justify-content-lg-around.order-2.order-lg-1.w-100.mw-100.w-lg-auto > div > div.o_cp_action_menus.d-flex.pe-2.gap-1 > div > button').click(timeout=40000)
        await page.wait_for_timeout(1500)
        
        # Clicar em mesclar
        await page.locator('span.o-dropdown-item:has-text("Mesclar")').click(timeout=5000)
        await page.wait_for_timeout(1500)
        
        # Confirmar mesclagem
        await page.get_by_role("button", name="Mesclar contatos").click(timeout=10000)

        # Aguardar processamento com timeout maior
        await page.locator('h2:has-text("Não há mais contatos para mesclar para esta solicitação")').wait_for(state="visible", timeout=10000)

        # Fechar diálogo
        await page.get_by_role("button", name="Fechar").click(timeout=5000)

        # Limpar pesquisa
        await search_input.click()
        await page.keyboard.press('Backspace')
        await page.wait_for_timeout(2500)

    except Exception as e:
        print(f"Erro ao processar {company_name}: {str(e)}")
        await page.screenshot(path=f'error_{company_name[:10]}.png')  # Debug visual

async def main():
    # Ler planilha
    df = pd.read_excel(EXCEL_FILE)
    companies = df['NomeEmpresa'].tolist()  # Alterar para o nome da coluna correto

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)  # Alterar para True em produção
        context = await browser.new_context()
        page = await context.new_page()
        
        await login(page)
        
        for company in companies:
            print(f"Processando: {company}")
            await process_company(page, company)
            await page.wait_for_timeout(3000)  # Intervalo entre empresas
        
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())