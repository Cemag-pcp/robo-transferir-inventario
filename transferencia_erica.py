import time
import sqlite3
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchWindowException

from utils import *

import psycopg2
from psycopg2.extras import DictCursor  # Para retornar resultados como dicionários

def verificar_transferencias():
    try:
        # Conectar ao banco de dados PostgreSQL
        conn = psycopg2.connect(
            dbname='postgres',  
            user='postgres',      
            password='15512332',    
            host='database-2.cdcogkfzajf0.us-east-1.rds.amazonaws.com',        
            port='5432'              
        )

        cursor = conn.cursor(cursor_factory=DictCursor)  # Usa DictCursor para obter resultados como dicionários

        # Executar a consulta com junção
        query = """
        SELECT
            st.id,
            st.quantidade,
            st.obs,
            st.data_solicitacao,
            f.matricula AS funcionario_nome,
            i.codigo AS item_nome,
            d.nome AS deposito_destino,
            st.data_entrega,
            st.rpa
        FROM
            almoxarifado_v2.solicitacao_solicitacaotransferencia st
        LEFT JOIN
            almoxarifado_v2.cadastro_funcionario f ON st.funcionario_id = f.id
        LEFT JOIN
            almoxarifado_v2.cadastro_itenstransferencia i ON st.item_id = i.id
        LEFT JOIN
            almoxarifado_v2.cadastro_depositodestino d ON st.deposito_destino_id = d.id
        WHERE st.data_entrega IS NOT NULL AND (st.rpa IS NULL OR st.rpa != 'OK') and (obs = 'Almox Corte e Estamparia' or obs = 'Almox Usinagem' or obs = 'Almox Serra')
        """

        cursor.execute(query)
        rows = cursor.fetchall()
        return rows

    except Exception as e:
        print(f"Erro ao conectar ao banco de dados ou executar a consulta: {e}")
        return []

    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def processar_transferencias(rows):
    if not rows:
        return

    conn = None
    cursor = None

    try:
        # Conectar ao banco PostgreSQL
        conn = psycopg2.connect(
            dbname='postgres',  
            user='postgres',      
            password='15512332',    
            host='database-2.cdcogkfzajf0.us-east-1.rds.amazonaws.com',        
            port='5432'              
        )

        cursor = conn.cursor()

        # Configuração do Selenium e navegação
        try:
            nav = webdriver.Chrome()
        except:
            chrome_driver_path = verificar_chrome_driver()
            nav = webdriver.Chrome(chrome_driver_path)
        
        nav.maximize_window()
        # nav.get("https://hcemag.innovaro.com.br/sistema/")
        nav.get("http://192.168.3.141/")

        # Login e navegação no sistema
        login(nav)
        time.sleep(5)
        menu_transferencia(nav)

        for row in rows:
            id_ = row[0]  # ID da linha atual

            try:
                rec = row[5]
                qtd = row[1]
                observacao_text = row[8]
                dep_destino = row[6]
                dep_origem = row[2]

                status = transferindo(nav, dep_origem,dep_destino , rec, qtd, observacao_text)

                # Atualizando o status no banco
                query_update = """UPDATE almoxarifado_v2.solicitacao_solicitacaotransferencia 
                                  SET rpa = %s WHERE id = %s"""
                cursor.execute(query_update, (status, id_))
                conn.commit()

                # Fechar aba no navegador
                try:
                    print('Clicando em fechar aba')
                    time.sleep(2)
                    fechar_aba = WebDriverWait(nav, 10).until(
                        EC.element_to_be_clickable(
                            (By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]')
                        )
                    )
                    fechar_aba.click()
                    time.sleep(1)
                except TimeoutException:
                    print('Erro ao fechar aba')
                time.sleep(0.5)

            except Exception as e:
                print(f"Erro ao processar a linha ID {id_}: {e}")
                continue  # Continua para a próxima linha

    except NoSuchWindowException:
        print("A janela do navegador foi fechada inesperadamente.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")
    finally:
        # Fechar cursor e conexão com o banco
        if cursor:
            cursor.close()
        if conn:
            conn.close()
        if nav:
            try:
                nav.quit()
            except Exception as e:
                print(f"Erro ao fechar o navegador: {e}")

def main():
    while True:
        rows = verificar_transferencias()
        
        if rows:
            print(f"Encontradas {len(rows)} transferencias a serem processadas.")
            processar_transferencias(rows)
        else:
            print("Nenhuma transferencia pendente encontrada.")

        # Aguarda um intervalo antes de verificar novamente
        time.sleep(300)  # Espera 5 minutos

if __name__ == "__main__":
    main()