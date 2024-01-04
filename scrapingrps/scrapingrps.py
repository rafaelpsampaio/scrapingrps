import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

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