import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
from sqlalchemy import create_engine

def relatorio(df,path_d,aba,codigocol = 'COD_CLIENTE',urlcol = 'URL',tam = 0,tempowait_download = 3000):
    tempowait = 300
    aba.get("https://hub.xpi.com.br/new/dashboard/#/performance")
    if tam == 0:
        tam = len(df)
    aba.execute_script("document.body.style.zoom='100%'")
    clierro = []
    clifaltando = []
    Clientes_df = df.loc[:tam-1].copy()
    if 'Patrimônio XP' in Clientes_df.columns:
        Clientes_df.sort_values(by = 'Patrimônio XP', ascending = False, inplace = True)
    for cliente in Clientes_df['COD_CLIENTE'].unique().tolist():
        achou = 0
        filename_loc = f"XPerformance - {cliente}"
        for filename in os.listdir(path_d):
            if filename.startswith(filename_loc):
                achou = 1
        if achou == 0:
            clifaltando.append(cliente)
        cont2 = 0
    tot = len(clifaltando)
    if len(clifaltando)>0:
        print(f'Iniciando! Quantidade de clientes: {len(clifaltando)}')
        start_time = time.time()
        for codigo in clifaltando:
            cont2 += 1
            url = Clientes_df.loc[Clientes_df[codigocol] == codigo,urlcol].values[0]
            site = f"https://hub.xpi.com.br/new/posicao-consolidada/#/{url}"
            aba.get(site)
            achoucod = 0
            travou = 1
            for t in range(0,tempowait):
                codigo_lis = aba.find_elements(By.XPATH,'//*[@id="root"]/arsenal-loader/div/div/div[1]/div[2]/soma-grid/soma-grid-row[2]/soma-grid-col/soma-coachmark/soma-coachmark-highlight/soma-card/div/div[3]/div[1]/div[1]/div/div/p')
                if len(codigo_lis)>0:
                    travou = 0
                    codigo_int = codigo_lis[0].text
                    break
                else:
                    time.sleep(0.1)
            if travou == 0:
                if codigo_int != str(codigo):
                    print('Esquisito')
                    print(codigo)
                    print(codigo_int)
                    travou = 1
                else:
                    try:
                        botaojs(aba,"#root > arsenal-loader > div > div > div.sc-eWHaVC.cPtGvU > soma-grid > soma-grid-row:nth-child(5) > soma-grid-col > soma-coachmark > soma-coachmark-highlight > soma-card > div.sc-iXzfSG.eRBiOJ > div.sc-iapWAC.dmTHUX > soma-button.sc-dPZUQH.kOsIJG.soma-button.primary.sm.hydrated').shadowRoot.querySelector('button')")
                        botaojs(aba,"#root > arsenal-loader > div > div > div.sc-eWHaVC.cPtGvU > soma-grid > soma-grid-row:nth-child(5) > soma-grid-col > soma-coachmark > soma-coachmark-highlight > soma-card > div.sc-iXzfSG.eRBiOJ > div.sc-iapWAC.dmTHUX > div.sc-eyvILC.SVORC > soma-modal > div > div.sc-hHOBiw.jKtwMT > soma-button').shadowRoot.querySelector('button')")
                        botaojs(aba,"#root > arsenal-loader > div > div > div.sc-eWHaVC.cPtGvU > soma-grid > soma-grid-row:nth-child(5) > soma-grid-col > soma-coachmark > soma-coachmark-highlight > soma-card > div.sc-iXzfSG.eRBiOJ > div.sc-iapWAC.dmTHUX > div.sc-eyvILC.SVORC > soma-modal > div > div.sc-hHOBiw.jKtwMT > div.sc-kzqdkY.iJmwrk > soma-button.soma-button.primary.md.hydrated').shadowRoot.querySelector('button')")
                    except:
                        print('Erro botões')
                        travou = 1
            if travou == 0:
                travou = 1
                filename_loc = f"XPerformance - {codigo_int}"
                for t in range(0,tempowait_download):
                    travou = 1
                    files = [f for f in os.listdir(path_d) if f.startswith(filename_loc)]
                    if len(files)>0:
                        travou = 0
                        break
                    else:
                        time.sleep(0.1)
            elapsed_time = time.time() - start_time
            avg_time_per_client = elapsed_time / (cont2)
            remaining_clients = tot - cont2
            estimated_time_remaining = avg_time_per_client * remaining_clients
            if travou == 1:
                print(f"Erro em {codigo}! {cont2} / {tot} ({round((cont2/len(clifaltando))*100,1)}%)). Estimativa de tempo restante: {estimated_time_remaining:.0f} segundos.")
                clierro.append(codigo)
            else:
                print(f"Download {codigo} completo! {cont2} / {tot} ({round((cont2/len(clifaltando))*100,1)}%)). Estimativa de tempo restante: {estimated_time_remaining:.0f} segundos.")
    return clierro
def navegacao(d_path = None):
    service = Service(ChromeDriverManager().install())
    if d_path is None:
        d = webdriver.Chrome(service=service)
    else:
        chrome_options = Options()
        prefs = {'download.default_directory': d_path}
        chrome_options.add_experimental_option('prefs', prefs)
        d = webdriver.Chrome(service=service, options=chrome_options)
    return d
def hub(aba,login,senha):
    url = 'https://hub.xpi.com.br/dashboard/#/performance'
    aba.get(url)
    element = aba.find_element("xpath", "//*[@id='username']")
    element.send_keys(login)
    elementPassword = aba.find_element("xpath", "//*[@id='twoDigitKeyboard']")
    elementPassword.send_keys(senha)
def botaojs(driver,pathjs):
    botao_js = f"return document.querySelector('{pathjs}"
    botao_element = driver.execute_script(botao_js)
    driver.execute_script("arguments[0].scrollIntoView(true);", botao_element)
    driver.execute_script("arguments[0].click();", botao_element)

def anbima(d_path, o_path,driver):
    driver.get("https://www.anbima.com.br/pt_br/informar/taxas-de-titulos-publicos.htm")
    driver.maximize_window()  # Maximize the browser window
    driver.execute_script("document.body.style.zoom='100%'")
    time.sleep(2)
    click_button()
    time.sleep(2)
    driver.find_element(By.XPATH, '/html/body/table/tbody/tr/td/div/table[5]/tbody/tr[1]/td/a').click()
    downloaded_file_path = wait_for_download(d_path)
    converted_file_path = convert_and_move_xls_to_xlsx(downloaded_file_path, o_path)
    print(f"File saved at {converted_file_path}")
def wait_for_download(d_path, timeout=300):
    """Wait until a file starting with m23 or m24 is downloaded."""
    seconds = 0
    downloaded_file = None
    while seconds < timeout:
        for file_name in os.listdir(d_path):
            if file_name.startswith(("m23", "m24")) and file_name.endswith(".xls"):  # Verificando o prefixo e a extensão
                downloaded_file = file_name
                return os.path.join(d_path, downloaded_file)
        time.sleep(1)
        seconds += 1
    raise Exception("Download did not complete after {} seconds".format(timeout))


def convert_and_move_xls_to_xlsx(file_path, destination_dir):
    all_sheets = pd.read_excel(file_path, engine='xlrd', sheet_name=None)
    base_name = os.path.basename(file_path).replace('.xls', '.xlsx')
    new_path = os.path.join(destination_dir, base_name)
    with pd.ExcelWriter(new_path, engine='xlsxwriter') as writer:
        for sheet_name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    time.sleep(2)
    try:
        os.remove(file_path)
    except PermissionError:
        print(f"Unable to delete {file_path}. It might be still in use.")
    return new_path


def click_button(driver):
    # Verifica se está dentro de um iframe
    iframes = driver.find_elements(By.TAG_NAME, 'iframe')
    for iframe in iframes:
        try:
            driver.switch_to.frame(iframe)
            button = driver.find_element(By.XPATH, '//*[@id="cinza50"]/form/div[2]/table/tbody/tr/td/img')

            # Scroll explícito até o botão
            driver.execute_script("arguments[0].scrollIntoView(true);", button)
            time.sleep(2)

            # Tente clicar com JavaScript
            driver.execute_script("arguments[0].click();", button)
            return
        except NoSuchElementException:
            driver.switch_to.default_content()
def baixarposicao(aba):
    aba.get('https://hub.xpi.com.br/new/renda-fixa/#/gerencial/exportacoes')
    tempo_max = 200
    try:
        # Esperar e encontrar o elemento que contém o shadow-root
        shadow_host = WebDriverWait(aba, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#root > arsenal-loader > soma-modal"))
        )

        # Acessar o shadow-root
        shadow_root = aba.execute_script('return arguments[0].shadowRoot', shadow_host)

        # Encontrar o botão dentro do shadow-root e clicar
        close_button = shadow_root.find_element(By.CSS_SELECTOR, "div > div > div > soma-icon > div > svg")
        close_button.click()
        print('Aviso fechado.')
    except:
        print(f'Erro ao fechar o aviso')
    clicou = 0
    loc_export_posicao = '//*[@id="root"]/arsenal-loader/soma-container[2]/arsenal-loader/main/div/div/div[3]/div/soma-card/soma-button'
    posicao_lis = aba.find_elements(By.XPATH,loc_export_posicao)
    for t in range(tempo_max):
        if len(posicao_lis) == 1 and clicou == 0:
            posicao = posicao_lis[0]
            if posicao.text == 'Exportar posição':
                try:
                    posicao.click()
                    clicou = 1
                except:
                    time.sleep(0.1)
            else:
                print('Texto estranho '+posicao.text)
        else:
            time.sleep(0.1)
        if clicou == 1:
            loc_gerar_excel = '//*[@id="root"]/arsenal-loader/soma-container[2]/arsenal-loader/main/div/div/div[3]/soma-drawer/soma-drawer-content/div/div[4]/div[2]/soma-button[2]'
            excel_lis = aba.find_elements(By.XPATH,loc_gerar_excel)
            if len(excel_lis)>0:
                excel = excel_lis[0]
                if excel.text == 'Gerar excel':
                    try:
                        excel.click()
                        break
                    except:
                        aba.get('https://hub.xpi.com.br/new/renda-fixa/#/gerencial/exportacoes')
                        time.sleep(0.1)
                else:
                    print('Estranho')
                    time.sleep(0.1)
            else:
                time.sleep(0.1)
    deucerto = 0
    download = 0
    loc_sucesso = '//*[@id="root"]/arsenal-loader/soma-container[2]/arsenal-loader/main/div/div/div[3]/soma-drawer/soma-drawer-content/div[2]/h2'
    for t in range(tempo_max):
        deu_lis = aba.find_elements(By.XPATH, loc_sucesso)
        if len(deu_lis) == 1 and deu_lis[0].text == 'Posição Geral e Renda Fixa exportado':
            print('Download concluído!')
            download = 1
            break
        else:
            time.sleep(0.1)
    if download == 1:
        try:
            shadow_host = WebDriverWait(aba, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#root > arsenal-loader > soma-container.sc-jxOSlx.nhvIP.soma-container.hydrated > arsenal-loader > main > div > div > div:nth-child(3) > soma-drawer > soma-drawer-content > div:nth-child(1) > soma-button"))
            )

            # Acessar o shadow-root
            shadow_root = aba.execute_script('return arguments[0].shadowRoot', shadow_host)
            button = shadow_root.find_element(By.CSS_SELECTOR, "button")
            button.click()
            print('Botão clicado.')
        except Exception as e:
            print(f'Erro ao clicar no botão: {e}')
def proc_pos_rf(d_path):
    latest_time = 0
    latest_file = None
    prefix = 'posicao-geral-renda-fixa'
    suffix = '.csv'
    files = [f for f in os.listdir(d_path) if f.startswith(prefix) and f.endswith(suffix)]
    for file_name in files:
        if file_name.startswith('posicao-geral-renda-fixa') and file_name.endswith(".csv"):
            file_path = os.path.join(d_path, file_name)
            file_time = os.path.getmtime(file_path)
            if file_time > latest_time:
                latest_time = file_time
                latest_file = file_path
    if latest_file is None:
        print('Arquivo não econtrado')
    else:
        print(f'Arquivo encontrado: {os.path.basename(latest_file)}')
    temp_df = pd.read_csv(os.path.join(d_path,latest_file), delimiter=';', thousands='"', decimal=',', encoding='utf-8')
    df = temp_df.copy()
    colunastempo = ['Dt_Aplicacao','Vencimento','Carencia','Juros','Amortizacao']
    for col in colunastempo:
        df[col] = df[col] = pd.to_datetime(df[col], format='%d/%m/%Y').dt.date
    df = df.fillna('-').replace('-',None)
    df.columns = df.columns.str.lower()
    return df

def subindo_rf_pos(df, DATABASE_URL):
    engine = create_engine(DATABASE_URL, echo=False)
    today = datetime.today().date()
    df.to_excel("I:\Guelt AAI\Middle & Operacional\Mesa RF\Oportunidades\Backup\EstoqueRF.xlsx",index = False)
    # Inicia uma transação
    with engine.begin() as connection:
        exists_today = connection.execute("""
            SELECT EXISTS (
                SELECT 1 FROM guelt_main.rf_posicao WHERE data_upload = %(today)s
            )
        """, {'today': today}).scalar()
        if exists_today:
            connection.execute("""
                DELETE FROM guelt_main.rf_posicao WHERE data_upload = %(today)s
            """, {'today': today})
        df['data_upload'] = today
        df['qtd'] = df['qtd'].astype('int64')
        df.to_sql('rf_posicao', con=connection, schema='guelt_main', if_exists='append', index=False, method='multi', chunksize=1000)


def check_and_click_button(driver, js_selector, t=0):
    try:
        botao = driver.execute_script(f"return {js_selector}")
        if botao:
            botao.click()
            time.sleep(0.5)
    except Exception as e:
        if t == 0:
            print(f"Erro ao tentar clicar no botão")


def extratcs(row, end):
    return row.find_element(By.CSS_SELECTOR, end).text


def scraping_produtos(ativo, tipo, driver):
    button = driver.execute_script("""
    return document.querySelector('hub-table-actions').shadowRoot
                   .querySelector('soma-button').shadowRoot
                   .querySelector('button[aria-label="Limpar filtro"]');
    """)
    if button:
        js_selector = "document.querySelector('#root > arsenal-loader > soma-container.sc-jxOSlx.nhvIP.soma-container.hydrated > arsenal-loader > div:nth-child(2) > hub-table-actions').shadowRoot.querySelector('div:nth-child(1) > soma-button').shadowRoot.querySelector('button')"
        check_and_click_button(driver, js_selector, 1)
    campo_busca = driver.execute_script(
        'return document.querySelector("#root > arsenal-loader > soma-container.sc-jxOSlx.nhvIP.soma-container.hydrated > arsenal-loader > div:nth-child(2) > hub-table-actions > hub-table-actions-search > soma-search").shadowRoot.querySelector("div > input[type=search]")')
    campo_busca.click()
    dados = None

    linhas_da_tabela_ant = driver.find_elements(By.CSS_SELECTOR,
                                                "#root > arsenal-loader > soma-container.sc-jxOSlx.nhvIP.soma-container.hydrated > arsenal-loader > soma-table > soma-table-body > div > div > soma-table-row"
                                                )
    tamog = (len(linhas_da_tabela_ant))
    campo_busca.send_keys(ativo)
    for t in range(0, 50):
        linhas_da_tabela = driver.find_elements(By.CSS_SELECTOR,
                                                "#root > arsenal-loader > soma-container.sc-jxOSlx.nhvIP.soma-container.hydrated > arsenal-loader > soma-table > soma-table-body > div > div > soma-table-row"
                                                )
        if tamog == len(linhas_da_tabela):
            time.sleep(0.1)
        else:
            break

    dados = []
    for linha in linhas_da_tabela:
        texto = extratcs(linha, "soma-table-cell.sc-hknOHE.dXWFBw.soma-table-cell.hydrated > div > div > soma-tooltip")
        if ativo in texto:
            vencimento = extratcs(linha, "soma-table-cell.sc-hknOHE.eVxouQ.soma-table-cell.hydrated")
            taxa_min = extratcs(linha, "soma-subtitle.sc-iMTnTL.jSZCll.soma-subtitle.hydrated")
            taxa_max = extratcs(linha, "soma-subtitle.sc-krNlru.bAAtCs.soma-subtitle.hydrated")
            incentivada = verificar_icone(linha, "[icon='encouraged-badge']")
            qualificado = verificar_icone(linha, "[icon='investor-qualified']")
            profissional = verificar_icone(linha, "[icon='investor-professional']")
            if tipo == 'CP':
                carencia = extratcs(linha, "soma-table-cell.sc-hknOHE.dXWFDX.soma-table-cell.hydrated")
                gross_min = extratcs(linha, "soma-subtitle.sc-jaXxmE.fKAYpw.soma-subtitle.hydrated")
                gross_max = extratcs(linha, "soma-subtitle.sc-ibQAlb.bPBieR.soma-subtitle.hydrated")
                rating = extratcs(linha, "soma-table-cell:nth-child(9)")
                qtdisp = extratcs(linha, "soma-table-cell:nth-child(8)")
                roa = extratcs(linha, "soma-table-cell:nth-child(10)")
                juros = None
            elif tipo == 'EB':
                carencia = extratcs(linha, "soma-table-cell.sc-hknOHE.dXWFDX.soma-table-cell.hydrated")
                gross_min = None
                gross_max = None
                rating = extratcs(linha, "soma-table-cell:nth-child(8)")
                qtdisp = extratcs(linha, "soma-table-cell:nth-child(7)")
                roa = extratcs(linha, "soma-table-cell:nth-child(9)")
                juros = None
            elif tipo == 'LF':
                rating = extratcs(linha, "soma-table-cell:nth-child(8)")
                roa = extratcs(linha, "soma-table-cell:nth-child(9)")
                qtdisp = extratcs(linha, "soma-table-cell:nth-child(7)")
                gross_min = extratcs(linha, "soma-subtitle.sc-jaXxmE.fKAYpw.soma-subtitle.hydrated")
                gross_max = extratcs(linha, "soma-subtitle.sc-ibQAlb.bPBieR.soma-subtitle.hydrated")
                carencia = None
                juros = extratcs(linha, "soma-table-cell:nth-child(10)")
            else:
                print(tipo)
            pu = extratcs(linha, "soma-table-cell.sc-hknOHE.eVxoBy.soma-table-cell.hydrated")
            qtdmin = extratcs(linha, "soma-table-cell.sc-hknOHE.kvJvjp.soma-table-cell.hydrated")
            risco = extratcs(linha,
                             "soma-table-cell.sc-hknOHE.krJJND.soma-table-cell.hydrated > soma-popover > soma-badge")
            dados.append({
                "Nome": texto,
                "Vencimento": vencimento,
                "Carencia": carencia,
                "Taxa Min": taxa_min,
                "Taxa Max": taxa_max,
                "Gross Up: Tx Min": gross_min,
                "Gross Up: Tx Maxn": gross_max,
                "PU": pu,
                "Qtd. Min": qtdmin,
                "Qtd. Disp": qtdisp,
                "Rating": rating,
                "Roa": roa,
                "Risco": risco,
                "Juros": juros,
                "Incentivada": incentivada,
                "Qualificado": qualificado,
                "Profissional": profissional
            })

    df = pd.DataFrame(dados)
    if df.empty:
        return None
    else:
        return df


def scrapingtp(driver):
    linhas_da_tabela = driver.find_elements(By.CSS_SELECTOR,
                                            "#root > arsenal-loader > soma-container.sc-jxOSlx.nhvIP.soma-container.hydrated > arsenal-loader > soma-table > soma-table-body > div > div > soma-table-row"
                                            )
    if len(linhas_da_tabela) == 0:
        time.sleep(0.5)
        linhas_da_tabela = driver.find_elements(By.CSS_SELECTOR,
                                                "#root > arsenal-loader > soma-container.sc-jxOSlx.nhvIP.soma-container.hydrated > arsenal-loader > soma-table > soma-table-body > div > div > soma-table-row"
                                                )
    else:
        dados = []
        for linha in linhas_da_tabela:
            texto = extratcs(linha,
                             "soma-table-cell.sc-hknOHE.dXWFBw.soma-table-cell.hydrated > div > div > soma-tooltip")
            vencimento = extratcs(linha, "soma-table-cell.sc-hknOHE.eVxouQ.soma-table-cell.hydrated")
            taxa_min = extratcs(linha, "soma-subtitle.sc-iMTnTL.jSZCll.soma-subtitle.hydrated")
            taxa_max = extratcs(linha, "soma-subtitle.sc-krNlru.bAAtCs.soma-subtitle.hydrated")
            pu = extratcs(linha, "soma-table-cell.sc-hknOHE.eVxoBy.soma-table-cell.hydrated")
            qtdmin = extratcs(linha, "soma-table-cell.sc-hknOHE.kvJvjp.soma-table-cell.hydrated")
            qtdisp = extratcs(linha, "soma-table-cell:nth-child(7)")
            roa = extratcs(linha, "soma-table-cell:nth-child(8)")
            risco = extratcs(linha,
                             "soma-table-cell.sc-hknOHE.krJJND.soma-table-cell.hydrated > soma-popover > soma-badge")
            carencia = None
            gross_min = None
            gross_max = None
            rating = None
            juros = None
            dados.append({
                "Nome": texto,
                "Vencimento": vencimento,
                "Carencia": carencia,
                "Taxa Min": taxa_min,
                "Taxa Max": taxa_max,
                "Gross Up: Tx Min": gross_min,
                "Gross Up: Tx Maxn": gross_max,
                "PU": pu,
                "Qtd. Min": qtdmin,
                "Qtd. Disp": qtdisp,
                "Rating": rating,
                "Roa": roa,
                "Risco": risco,
                "Juros": juros
            })
        return pd.DataFrame(dados)


def scraping_prateleira(aba, depara):
    aba.get("https://hub.xpi.com.br/new/renda-fixa/#/operacional")
    papeis = depara.copy()
    papeis.sort_values(by=['Tipo', 'Emissor'], ascending=[True, True], inplace=True)
    button_map = {
        'EB': 'OPTION-EMISSAOBANCARIA',
        'CP': 'OPTION-CREDITOPRIVADO',
        'TP': 'OPTION-TITULOSPUBLICOS',
        'LF': 'OPTION-LETRAFINANCEIRAS'
    }
    tabelao = pd.DataFrame()
    for tipo in papeis['Tipo'].unique():
        tab_tipo = pd.DataFrame()
        if tipo in button_map:
            button_id = button_map[tipo]
            button = aba.find_element(By.ID, button_id)
            if 'selected' not in button.get_attribute('class'):
                button.click()
            for emissor in papeis.loc[papeis['Tipo'] == tipo, 'Emissor'].unique():
                dfe = scraping_produtos(emissor, tipo, aba)
                if not dfe is None:
                    dfe['Emissor'] = emissor
                    tab_tipo = pd.concat([tab_tipo, dfe], ignore_index=True)
            tab_tipo['Tipo'] = tipo
            tabelao = pd.concat([tabelao, tab_tipo], ignore_index=True)
        else:
            print(f"Abreviação '{tipo}' não reconhecida.")
    button_id = button_map["TP"]
    button = aba.find_element(By.ID, button_id)
    if 'selected' not in button.get_attribute('class'):
        button.click()
    time.sleep(0.5)
    tp_df = scrapingtp(aba)
    if tp_df is not None and not tp_df.empty:
        tp_df['Tipo'] = 'TP'
    else:
        time.sleep(0.5)
        tp_df = scrapingtp(aba)
        tp_df['Tipo'] = 'TP'
    tabelao = pd.concat([tabelao, tp_df], ignore_index=True)
    hoje = datetime.today()
    tabelao['Atualização'] = hoje
    return tabelao


def imprimir_ativos(titulo, ativos):
    if len(ativos) > 0:
        ativos_formatados = "\n".join(sorted(ativos))
        print(f"{titulo}:\n{ativos_formatados}\n")


def imprimindo(df_tb_og):
    df_tb = df_tb_og.copy()
    maxhj = df_tb.loc[df_tb['Atualização'] < df_tb['Atualização'].max(), 'Atualização'].max()
    ultimo = df_tb.loc[df_tb['Atualização'] == maxhj].reset_index(drop=True).copy()
    df_tb['Dia'] = pd.to_datetime(pd.to_datetime(df_tb['Atualização'])).dt.date
    ontem = df_tb.loc[df_tb['Dia'] < datetime.today().date(), 'Dia'].max()
    ontem_df = df_tb.loc[df_tb['Dia'] == ontem].reset_index(drop=True).copy()
    minontem = ontem_df['Atualização'].max()
    atual = df_tb.loc[df_tb['Atualização'] == df_tb['Atualização'].max()]

    ativos_novos = set(atual.loc[~atual['Nome'].isin(df_tb['Nome'].unique()), 'Nome'])
    ativos_ontem = set(ontem_df['Nome'].unique())
    ativos_ant = set(ultimo['Nome'].unique())
    ativos_agr = set(atual['Nome'].unique())

    ativos_add_agr = ativos_agr - ativos_ant
    ativos_removidos_agr = ativos_ant - ativos_agr
    ativos_add_hj = ativos_agr - ativos_ontem
    ativos_removidos_hj = ativos_ontem - ativos_agr

    imprimir_ativos("Ativos adicionados agora", ativos_add_agr)
    imprimir_ativos("Ativos removidos agora", ativos_removidos_agr)
    imprimir_ativos("Ativos adicionados hoje", ativos_add_hj)
    imprimir_ativos("Ativos removidos hoje", ativos_removidos_hj)
    imprimir_ativos("Ativos novos", ativos_novos)


def verificar_icone(linha, seletor_icone):
    try:
        icone = linha.find_element(By.CSS_SELECTOR, seletor_icone)
        return True if icone else False
    except:
        return False

