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


def verificar_requisicoes():
    
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
                sr.id,
                sr.quantidade,
                sr.obs,
                sr.data_solicitacao,
                cac.nome as classe_requisicao,
                cc1.codigo AS cc_nome,
                f.matricula AS funcionario_nome,
                i.codigo AS item_nome,
                sr.data_entrega,
                sr.rpa
            FROM
                almoxarifado_v2.solicitacao_solicitacaorequisicao sr
            JOIN
                almoxarifado_v2.cadastro_cc cc1 ON sr.cc_id = cc1.id
            LEFT JOIN
                almoxarifado_v2.cadastro_classerequisicao cac ON sr.classe_requisicao_id = cac.id    
            LEFT JOIN
                almoxarifado_v2.cadastro_funcionario f ON sr.funcionario_id = f.id
            LEFT JOIN
                almoxarifado_v2.cadastro_itenssolicitacao i ON sr.item_id = i.id
            WHERE
                sr.data_entrega IS NOT NULL 
                AND (sr.rpa IS NULL OR sr.rpa != 'OK');
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

def processar_requisicoes(rows):
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
        chrome_driver_path = verificar_chrome_driver()
        nav = webdriver.Chrome(chrome_driver_path)
        nav.maximize_window()
        # nav.get("https://hcemag.innovaro.com.br/sistema/")
        nav.get("http://192.168.3.141/")

        # Login e navegação no sistema
        login(nav)
        time.sleep(5)
        menu_requisicao(nav)

        for row in rows:
            id_ = row[0]  # ID da linha atual

            try:
                # Processar cada linha
                rec = row[7]
                qtd = row[1]
                tipo_requisicao = row[4]
                requisitante_matricula = row[6]
                ccusto_text = row[5]
                observacao_text = row[2]

                status = requisitando(nav, rec, qtd, tipo_requisicao, requisitante_matricula, ccusto_text, observacao_text) 
                
                # Atualizar o banco de dados
                if status != 'OK':
                    query_update = f"""UPDATE almoxarifado_v2.solicitacao_solicitacaorequisicao SET rpa = '{status}' WHERE id = {id_}"""
                    cursor.execute(query_update)
                    conn.commit()

                    # Fechar aba e continuar
                    try:
                        print('Clicando em fechar aba')
                        time.sleep(2)
                        fechar_aba = WebDriverWait(nav, 10).until(
                            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]'))
                        )
                        fechar_aba.click()
                        time.sleep(1)
                    except TimeoutException:
                        print('Erro ao fechar aba')
                    time.sleep(0.5)

                    continue  # Segue para a próxima linha se houver erro

                # Atualizar a linha na tabela
                query_update = f"""UPDATE almoxarifado_v2.solicitacao_solicitacaorequisicao SET rpa = '{status}' WHERE id = {id_}"""
                cursor.execute(query_update)
                conn.commit()  # Confirma a transação

            except Exception as e:
                print(f"Erro ao processar a linha ID {id_}: {e}")
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

def main():
    while True:
        rows = verificar_requisicoes()
        
        if rows:
            print(f"Encontradas {len(rows)} requisições a serem processadas.")
            processar_requisicoes(rows)
        else:
            print("Nenhuma requisição pendente encontrada.")

        # Aguarda um intervalo antes de verificar novamente
        time.sleep(300)  # Espera 5 minutos

if __name__ == "__main__":
    main()  