from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, NoSuchWindowException
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import time

from utils import *


while True:

    try:

        # API GOOGLE PLANILHAS
        scope = ['https://www.googleapis.com/auth/spreadsheets',
                    "https://www.googleapis.com/auth/drive"]

        filename = "service_account_cemag.json"
        credentials = ServiceAccountCredentials.from_json_keyfile_name(filename, scope)
        client = gspread.authorize(credentials)
        sa = gspread.service_account(filename)

        sheet = 'ajuste inventario 2024'
        worksheet = 'transferência'

        sh = sa.open(sheet)
        wks = sh.worksheet(worksheet)

        df = wks.get()

        tabela_completa = pd.DataFrame(df)

        tabela_completa.columns = tabela_completa.iloc[0]

        tabela_completa = tabela_completa[1:]
        tabela_completa.reset_index(inplace=True)

        tabela_completa['STATUS'] = tabela_completa['STATUS'].fillna('')

        tabela = tabela_completa[tabela_completa['STATUS'] != 'OK']

        tabela.reset_index(drop=True,inplace=True)

        if tabela.empty:
            print('❌ Sem dados na tabela ❌')
            time.sleep(20)
            continue

        print(tabela)

        # acessando site
        link = "https://hcemag.innovaro.com.br/sistema"
        nav = webdriver.Chrome()
        nav.maximize_window()
        nav.get(link)

        # login e senha
        login(nav)

        # menu innovaro
        menu_transferencia(nav)

        # Conectar com a planilha google
        dep_origem = 'Almox central'
        dep_destino = 'mat fora uso'
        rec = tabela['CÓDIGO'].to_list() 
        linha = tabela['index'].to_list() 
        # ['028844','034953','110012','028844','034953','110012','028844','034953','110012','028844','034953','110012']

        qtd=tabela['DIFERENÇA'].to_list() 

        for index,recurso in enumerate(rec): 
            print(f"#### Iniciando Linha {linha[index]+1} ####")
            status = transferindo(nav,dep_origem,dep_destino,recurso,qtd[index])
            if status == 'Erro ao clicar no menu':
                print(status)
                break
            print(f"#### Concluindo Linha {linha[index]+1} ####")
            wks.update('C' + str(linha[index]+1), status)

    except NoSuchWindowException as e:
        print(f"A janela do navegador foi fechada inesperadamente.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")