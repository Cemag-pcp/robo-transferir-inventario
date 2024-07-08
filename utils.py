from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException,StaleElementReferenceException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

import time
import pandas as pd

def login(nav):

    try:
        # logando
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="username"]'))).send_keys("Ti.prod")
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="password"]'))).send_keys("Cem@@1600")
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="password"]'))).send_keys(Keys.ENTER)

        print("Acessou a página de login")

    except Exception as e:
        print(f"Ocorreu um erro durante o login: {e}")

def menu_transferencia(nav):
    
    nav.switch_to.default_content()
    
    #menu
    try:
        menu=WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.CLASS_NAME, 'menuBar-button-label')))
        time.sleep(2.5)
        menu.click()
        time.sleep(2.5)
        print('Menu aberto')
    except TimeoutException:
        print('Erro ao clicar no menu')
        return
    time.sleep(.5)
    
    #Clicando em estoque
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(0.5)
    click_producao = test_list.loc[test_list[0] == 'Estoque'].reset_index(drop=True)['index'][0]
    lista_menu[click_producao].click()
    time.sleep(.5)
        
    #Clicando em Transferência
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(0.5)
    click_producao = test_list.loc[test_list[0] == 'Transferência'].reset_index(drop=True)['index'][0]
    lista_menu[click_producao].click()
    time.sleep(.5)
    
    #menu
    try:
        menu=WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.CLASS_NAME, 'menuBar-button-label')))
        time.sleep(2.5)
        menu.click()
        time.sleep(2.5)
    except TimeoutException:
        print('Erro ao clicar no menu')
        return
    time.sleep(.5)

def transferindo(nav,dep_origem,dep_destino,rec,qtd):
        
    #menu
    try:
        menu=WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.CLASS_NAME, 'menuBar-button-label')))
        time.sleep(2.5)
        menu.click()
        time.sleep(2.5)
        print('Menu aberto')
    except TimeoutException:
        print('Erro ao clicar no menu')
        return 'Erro ao clicar no menu'
    time.sleep(.5)
    
    #Clicando em solicitação de transferencia entre depositos
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(0.5)
    click_producao = test_list.loc[test_list[0] == 'Solicitação de transferência entre depósitos'].reset_index(drop=True)['index'][0]
    lista_menu[click_producao].click()
    
    # Carregando ao clicar no MENU
    while True:
        elements = nav.find_elements(By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span[2]')
        if len(elements) >= 1:
            break
        else:
            print('...Carregando')
            time.sleep(1)  # Esperar 1 segundo antes de verificar novamente 
    
    #Mudando de iframe
    iframes(nav)
    
    #Clicando em Mudar visualização
    try:
        print('clicando em mudar visualização')
        mudar_visualizacao=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]')))
        mudar_visualizacao.click()
    except TimeoutException:
        print('Erro ao mudar visualização')
        return 'Erro ao mudar visualização'
    time.sleep(.5)
    
    #Clicando em insert
    try:
        print('clicando em insert')
        insert=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]')))
        insert.click()
    except TimeoutException:
        print('Erro ao da insert')
        return 'Erro ao da insert'
    time.sleep(.5)
    
    #inputando deposito origem 
    try:
        print('Depósito origem')
        deposito_origem=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[9]/td[2]/table/tbody/tr/td[1]')))
        deposito_origem.click()
        print('Clicou no Depósito origem')
        deposito_origem_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[9]/td[2]/table/tbody/tr/td[1]/input')))
        deposito_origem_input.clear()
        deposito_origem_input.send_keys(dep_origem)
        deposito_origem_input.send_keys(Keys.TAB)
        
    except TimeoutException:
        print(f'Erro ao inputar Depósito origem: {dep_origem}')
        return f'Erro ao inputar Depósito origem: {dep_origem}'
    time.sleep(.5)
    
    #inputando deposito destino
    try:
        print('Depósito destino')
        deposito_destino=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[11]/td[2]/table/tbody/tr/td[1]')))
        deposito_destino.click()
        
        deposito_destino_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[11]/td[2]/table/tbody/tr/td[1]/input')))
        deposito_destino_input.clear()
        deposito_destino_input.send_keys(dep_destino)
        deposito_destino_input.send_keys(Keys.TAB)
        
    except TimeoutException:
        print(f'Erro ao inputar Depósito destino: {dep_destino}')
        return f'Erro ao inputar Depósito destino: {dep_destino}'
    time.sleep(.5)
    
    #inputando recurso
    try:
        print('inputando Recurso')
        recurso=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[13]/td[2]/table/tbody/tr/td[1]')))
        recurso.click()
        
        recurso_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[13]/td[2]/table/tbody/tr/td[1]/input')))
        recurso_input.clear()
        recurso_input.send_keys(rec)
        recurso_input.send_keys(Keys.TAB)
        
    except TimeoutException:
        print(f'Erro ao inputar recurso: {rec}')
        return f'Erro ao inputar recurso: {rec}'
    time.sleep(.5)
    
    #inputando quantidade
    try:
        print('inputando quantidade')
        quantidade=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[23]/td[2]/table/tbody/tr/td[1]')))
        quantidade.click()
        
        quantidade_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[23]/td[2]/table/tbody/tr/td[1]/input')))
        quantidade_input.clear()
        quantidade_input.send_keys(qtd)
        quantidade_input.send_keys(Keys.TAB)
        
    except TimeoutException:
        print(f'Erro ao inputar quantidade: {qtd}')
        return f'Erro ao inputar quantidade: {qtd}'
    time.sleep(.5)
        
    #Clicando em confirmar (insert)
    try:
        print('clicando em confirmar')
        confirmar=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="solicitacoes"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]')))
        confirmar.click()
    except TimeoutException:
        print('Erro ao confirmar')
        return 'Erro ao confirmar'
    time.sleep(.5)
    
    #Clicando em aprovar
    try:
        print('clicando em aprovar')
        # aprovar=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="buttonsBar_solicitacoes"]/td[1]')))
        # aprovar.click()
        
        button = WebDriverWait(nav, 1).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="buttonsBar_solicitacoes"]/td[1]'))
        )

        # Usar JavaScript para clicar no botão
        nav.execute_script("arguments[0].click();", button)

        nav.switch_to.default_content()

        # Carregando até aparecer o MODAL para CONFIRMAR
        while True:
            elements = nav.find_elements(By.XPATH, '//*[@id="confirm"]')
            if len(elements) >= 1:
                break
            else:
                print('...Carregando')
                time.sleep(1)  # Esperar 1 segundo antes de verificar novamente

        confirmar=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="confirm"]')))
        confirmar.click()
        iframes(nav)
    except TimeoutException:
        print('Erro ao aprovar')
        return 'Erro ao aprovar'
    time.sleep(1.5)
    
    #Clicando em baixar
    try:
        print('clicando em baixar')
        baixar=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="buttonsBar_solicitacoes"]/td[3]')))
        baixar.click()
        
        data_baixa=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="informaçõesDaBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]')))
        data_baixa.click()
        
        # Carregando até aparecer o campo de DATA
        while True:
            elements = nav.find_elements(By.XPATH, '//*[@id="informaçõesDaBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input')
            if len(elements) >= 1:
                break
            else:
                print('...Carregando')
                time.sleep(1)  # Esperar 1 segundo antes de verificar novamente

        data_baixa_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="informaçõesDaBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input')))
        data_baixa_input.clear()
        time.sleep(1)
        data_baixa_input.send_keys('01/07/2024')
        time.sleep(1)
        
        nav.switch_to.default_content()

        confirmar_baixa=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td')))
        confirmar_baixa.click()
        mensagem_erro = None

        # Carregando até terminar o LOADING após clicar em CONFIRMAR BAIXA 
        while True:
            elements = nav.find_elements(By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span[2]')
            if len(elements) >= 1:
                break
            else:
                if len(nav.find_elements(By.XPATH, '//*[@id="confirm"]')) >= 1:
                    time.sleep(3)
                    mensagem_erro = nav.find_elements(By.CLASS_NAME, 'message_errorToHtml')
                    texto_erro = mensagem_erro[0].text
                    confirm = WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="confirm"]')))
                    confirm.click()
                    break
                try:
                    confirmar_baixa.click()
                except StaleElementReferenceException as e:
                    pass
                print('...Carregando')
                time.sleep(3)  # Esperar 3 segundos antes de verificar novamente 
        # Pegando a mensagem de erro caso o SALDO seja insuficiente
        if mensagem_erro:
            try:
                print('clicando em fechar aba')
                fechar_aba=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]')))
                fechar_aba.click()
                time.sleep(1)
                print(texto_erro)
                return texto_erro
            except TimeoutException:
                print('Erro ao fechar aba')
                return 'Erro ao fechar aba'
            # Registrar no Google Sheets 
            # Fechar Aba
            # Recomeçar o processo
    except TimeoutException:
        print('Erro ao baixar')
        return 'Erro ao baixar'
    time.sleep(1.5)
    #Clicando em aprovar
    try:
        print('clicando em gravar')
        nav.switch_to.default_content()
        gravar=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]')))
        gravar.click()
        iframes(nav)
        # Até aparecer a nova página após clicar em GRAVAR
        while True:
            elements = nav.find_elements(By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')
            if len(elements) >= 1:
                break
            else:
                print('...Carregando')
                time.sleep(1)  # Esperar 1 segundo antes de verificar novamente
        
    except TimeoutException:
        print('Erro ao gravar')
        return 'Erro ao gravar'
    time.sleep(1)
    nav.switch_to.default_content()
    
    #fechar aba
    try:
        print('clicando em fechar aba')
        fechar_aba=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]')))
        fechar_aba.click()
        time.sleep(1)
                
    except TimeoutException:
        print('Erro ao fechar aba')
        return 'Erro ao fechar aba'
    time.sleep(.5)    

    return 'OK'
    
# nav.switch_to.default_content()
    
def iframes(nav):
    
    iframe_list = nav.find_elements(By.CLASS_NAME, 'tab-frame')

    for iframe in range(len(iframe_list)):
        time.sleep(1)
        try:
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass

def listar(nav, classe):

    try:

        lista_menu = nav.find_elements(By.CLASS_NAME, classe)

        elementos_menu = []

        for x in range(len(lista_menu)):
            a = lista_menu[x].text
            elementos_menu.append(a)

        test_lista = pd.DataFrame(elementos_menu)
        test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

        print("listou as opções do menu")

    except Exception as e:
        print(f"Ocorreu um erro durante a listagem de opções: {e}")

    return (lista_menu, test_lista)