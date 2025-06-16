import time
import sqlite3
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchWindowException
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from utils import *

import psycopg2
from psycopg2.extras import DictCursor  # Para retornar resultados como dicionários

def leitura_google_planilhas(worksheet, key_sheets):

    scope = ['https://www.googleapis.com/auth/spreadsheets',
                    "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name("service_account_cemag.json", scope)
    client = gspread.authorize(credentials)
    # Conectando com google sheets e acessando Análise Previsão de Consumo (CMM / NTP ) DEE

    sh = client.open_by_key(key_sheets)
    wks = sh.worksheet(worksheet)
    
    wks.update_acell('L2295', 'teste')

    df = wks.get()

    dados = pd.DataFrame(df)

    dados.columns = dados.iloc[4]  # linha 4 (index 3) vira o cabeçalho
    dados = dados.drop(index=range(0, 5)).reset_index(drop=True)  # remove as 4 primeiras linhas e reseta o índice
    
    dados = dados[((dados['Status'] == '') | (dados['Status'] == 'ERRO')) & (dados['Data'] != '')]

    return dados, wks

def preencher_google_planilhas(worksheet, key_sheets, index, status):

    scope = ['https://www.googleapis.com/auth/spreadsheets',
                    "https://www.googleapis.com/auth/drive"]

    credentials = ServiceAccountCredentials.from_json_keyfile_name("service_account_cemag.json", scope)
    client = gspread.authorize(credentials)
    # Conectando com google sheets e acessando Análise Previsão de Consumo (CMM / NTP ) DEE

    sh = client.open_by_key(key_sheets)
    wks = sh.worksheet(worksheet)
    
    wks.update_acell('L'+str(index), status)

def processar_transferencias(rows):
    if not rows:
        return

    try:
        # Configuração do Selenium e navegação
        try:
            nav = webdriver.Chrome()
        except:
            chrome_driver_path = verificar_chrome_driver()
            nav = webdriver.Chrome(chrome_driver_path)
        
        nav.maximize_window()
        # nav.get("https://hcemag.innovaro.com.br/sistema/")
        # nav.get("http://192.168.3.141/")
        # nav.get("http://192.168.3.140/")
        nav.get("http://127.0.0.1/sistema")
        # nav.get("https://hcemag.innovaro.com.br/sistema") # base de teste

        # Login e navegação no sistema
        login(nav)
        time.sleep(5)
        menu_transferencia(nav)

        for index, row in rows.iterrows():

            try:
                rec = row['Código Chapa'] # recurso
                qtd = row['Peso'] # quantidade
                observacao_text = None
                dep_destino = 'Almox corte e estamparia' # deposito destino
                linha_planilha = index + 6

                status = transferindo(nav, 'almox central',dep_destino , rec, qtd, observacao_text)

                preencher_google_planilhas("RQ PCP-003-000 (Transferencia Corte)","1t7Q_gwGVAEwNlwgWpLRVy-QbQo7kQ_l6QTjFjBrbWxE",linha_planilha,status)

                # Fechar aba no navegador
                try:
                    print('Clicando em fechar aba')
                    time.sleep(2)
                    WebDriverWait(nav, 1).until(EC.element_to_be_clickable((
                        By.XPATH, "//span[contains(@onclick, 'Environment.getInstance().closeTab')]/div"))).click()
                    time.sleep(1)
                except TimeoutException:
                    print('Erro ao fechar aba')
                time.sleep(0.5)

            except Exception as e:
                print(f"Erro ao processar a linha ID {index}: {e}")
                continue  # Continua para a próxima linha

    except NoSuchWindowException:
        print("A janela do navegador foi fechada inesperadamente.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
    finally:
        if nav:
            try:
                nav.quit()
            except Exception as e:
                print(f"Erro ao fechar o navegador: {e}")

def main():
    while True:
        rows, wks = leitura_google_planilhas("RQ PCP-003-000 (Transferencia Corte)","1t7Q_gwGVAEwNlwgWpLRVy-QbQo7kQ_l6QTjFjBrbWxE")
        
        if rows:
            print(f"Encontradas {len(rows)} transferencias a serem processadas.")
            processar_transferencias(rows)
        else:
            print("Nenhuma transferencia pendente encontrada.")

        # Aguarda um intervalo antes de verificar novamente
        time.sleep(300)  # Espera 5 minutos

if __name__ == "__main__":
    main()