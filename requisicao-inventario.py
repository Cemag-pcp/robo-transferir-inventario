import time
import sqlite3
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchWindowException
from utils_copy import *

import psycopg2
from psycopg2.extras import DictCursor  # Para retornar resultados como dicionários

import sys
sys.path.append(r'C:\Users\TI DEV\transferencia_automatica\robo-transferir-inventario')

def verificar_requisicoes():
    
    dados = pd.read_csv("dados.csv")
    
    return dados    

def processar_requisicoes(dados):

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

        # Login e navegação no sistema
        login(nav)
        time.sleep(5)
        menu_requisicao(nav)

        for index,row in dados.iterrows():
            if row['status'] != 'OK' or pd.isna(row['status']):

                try:
                    # Processar cada linha
                    rec = row['codigo']
                    qtd = row['qtd']
                    tipo_requisicao = 'Req p inventario'
                    requisitante_matricula = '4054'
                    ccusto_text = '2000'
        
                    status = requisitando(nav, rec, qtd, tipo_requisicao, requisitante_matricula, ccusto_text) 
                    
                    # Atualizar o banco de dados
                    if status != 'OK':
                        
                        dados.at[index,'status'] = status
                        dados.to_csv("dados.csv", index=False)

                        # Fechar aba e continuar
                        try:
                            print('Clicando em fechar aba')
                            time.sleep(2)
                            WebDriverWait(nav, 1).until(EC.element_to_be_clickable((
                                By.XPATH, "//span[contains(@onclick, 'Environment.getInstance().closeTab')]/div"))).click()
                            time.sleep(1)
                        except TimeoutException:
                            print('Erro ao fechar aba')
                        time.sleep(0.5)

                        continue  # Segue para a próxima linha se houver erro

                    # Atualizar a linha na tabela
                    dados.at[index,'status'] = status
                    dados.to_csv("dados.csv", index=False)

                except Exception as e:
                    print(f"Erro ao processar: {e}")
                    continue  # Segue para a próxima linha em caso de erro

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

# def main():
#     while True:
#         rows = verificar_requisicoes()
        
#         if rows:
#             print(f"Encontradas {len(rows)} requisições a serem processadas.")
#             processar_requisicoes(rows)
#         else:
#             print("Nenhuma requisição pendente encontrada.")

#         # Aguarda um intervalo antes de verificar novamente
#         time.sleep(300)  # Espera 5 minutos

# if __name__ == "__main__":
#     main()  
    

while True:
    
    dados = verificar_requisicoes()
    processar_requisicoes(dados)

