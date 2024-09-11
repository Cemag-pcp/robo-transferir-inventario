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
import datetime
import requests
import zipfile
import subprocess
import os

def verificar_chrome_driver():
    # URL que você deseja acessar
    url = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"

    # Faz uma solicitação GET à URL
    response = requests.get(url)

    if response.status_code == 200:
        data = response.json()

        # Versão que você está procurando
        target_version = get_chrome_version()

        def version_key(version):
            # Converte a versão em uma lista de inteiros para comparação
            return list(map(int, version.split('.')))

        # Filtra as versões disponíveis
        versions = [v['version'] for v in data['versions']]

        # Encontra a versão mais próxima
        closest_version = None
        min_diff = float('inf')

        for version in versions:
            # Calcula a diferença em cada nível de versão
            current_diff = sum(
                abs(a - b) for a, b in zip(version_key(version), version_key(target_version)))

            if current_diff < min_diff:
                min_diff = current_diff
                closest_version = version

        # Filtra as versões disponíveis
        versions_info = data['versions']

        # Armazena versões que correspondem à versão desejada
        matching_versions = []

        for version_info in versions_info:
            version = version_info['version']

            # Verifica se a versão corresponde à versão desejada
            if version == closest_version:
                matching_versions.append(version_info)

        # Exibe todos os registros encontrados
        if matching_versions:
            print("Registros encontrados:")
            for match in matching_versions:
                download_chrome_driver = match['downloads']['chromedriver']
        else:
            print("Nenhum registro encontrado.")

        def get_url_by_platform(data, platform):
            for item in data:
                if item['platform'] == platform:
                    return item['url']
            return None

        # Exemplo de uso:
        url = get_url_by_platform(download_chrome_driver, 'win32')
        filename = 'chromedriver-win32.zip'
        download_file(url, filename)

        extract_to = 'chromedriver_extracted'
        unzip_file(filename, extract_to)

        chromedriver_path = find_chromedriver(extract_to)
        if chromedriver_path:
            print(f"chromedriver.exe encontrado em: {chromedriver_path}")
        else:
            print("chromedriver.exe não encontrado.")

        return chromedriver_path

    else:
        print(f"Falha ao acessar a URL: {response.status_code}")

def download_file(url, filename):
    response = requests.get(url)
    if response.status_code == 200:
        with open(filename, 'wb') as file:
            file.write(response.content)
        print(f"Download concluído: {filename}")
    else:
        print(f"Falha no download: {response.status_code}")

def unzip_file(zip_path, extract_to='.'):
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)
        print(f"Arquivos extraídos para: {extract_to}")

def find_chromedriver(extract_to='.'):
    for root, dirs, files in os.walk(extract_to):
        if 'chromedriver.exe' in files:
            return os.path.join(root, 'chromedriver.exe')
    return None

def get_chrome_version():
    try:
        # Executa o comando para obter a versão do Chrome
        version = subprocess.check_output(
            ['reg', 'query', r'HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon',
                '/v', 'version'],
            stderr=subprocess.STDOUT
        )
        # Processa a saída para extrair a versão
        version = version.decode().strip().split()[-1]
        return version
    except subprocess.CalledProcessError:
        return "Chrome não está instalado."

def login(nav):

    try:
        # logando
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="username"]'))).send_keys("ti.cemag")
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="password"]'))).send_keys("cem@#1501")
        WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="password"]'))).send_keys(Keys.ENTER)

        print("Acessou a página de login")

    except Exception as e:
        print(f"Ocorreu um erro durante o login: {e}")

def menu_innovaro_2(nav):
    
    """
    Função para abrir ou fechar menu no innovaro do tipo 2
    :nav: webdriver
    """
    
    #abrindo menu

    try:
        nav.switch_to.default_content()
    except:
        pass

    WebDriverWait(nav, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="bt_1898143037"]/table/tbody/tr/td[2]'))).click()

    time.sleep(2)

def menu_innovaro_1(nav):
    
    """
    Função para abrir ou fechar menu no innovaro do tipo 1
    :nav: webdriver
    """
    
    #abrindo menu

    try:
        nav.switch_to.default_content()
    except:
        pass

    menu=WebDriverWait(nav, 10).until(EC.element_to_be_clickable(
            (By.CLASS_NAME, 'menuBar-button-label')))
    time.sleep(2.5)
    menu.click()
    time.sleep(2.5)

def menu_transferencia(nav):
    
    nav.switch_to.default_content()
    
    #menu
    try:
        menu_innovaro_1(nav)
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
        menu_innovaro_1(nav)
        print('Menu fechado')
    except TimeoutException:
        print('Erro ao clicar no menu')
        return
    time.sleep(.5)

def menu_requisicao(nav):
    
    nav.switch_to.default_content()
    
    #menu
    try:
        menu_innovaro_1(nav)
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
    click_producao = test_list.loc[test_list[0] == 'Requisição'].reset_index(drop=True)['index'][0]
    lista_menu[click_producao].click()
    time.sleep(.5)
    
    #menu
    try:
        menu_innovaro_1(nav)
        print('Menu fechado')
    except TimeoutException:
        print('Erro ao clicar no menu')
        return
    time.sleep(.5)

def transferindo(nav,dep_origem,dep_destino,rec,qtd,observacao_text):
        
    nav.switch_to.default_content()
        
    #menu
    try:
        menu_innovaro_1(nav)
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

    WebDriverWait(nav,60).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]/span[2]')))
    
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
        time.sleep(0.5)
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
        time.sleep(0.5)
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
        time.sleep(0.5)
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
        time.sleep(0.5)
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

        confirmar=WebDriverWait(nav,60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="confirm"]')))
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
        
        # WebDriverWait(nav,60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="informaçõesDaBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input')))

        data_baixa_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="informaçõesDaBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input')))
        data_baixa_input.send_keys(Keys.CONTROL + 'A')
        time.sleep(2)
        data_baixa_input.send_keys(datetime.datetime.now().date().strftime("%d/%m/%Y"))
        data_baixa_input.send_keys(Keys.TAB)
        time.sleep(2)
        
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
        # Até aparecer a nova página após clicar em GRAVAR
        mensagem_erro = None
        try:
            WebDriverWait(nav,15).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="confirm"]')))
        except:
            print("Sem nenhum erro contábil")
        if len(nav.find_elements(By.XPATH, '//*[@id="confirm"]')) >= 1:
            print("Erro, clicando em confirmar")
            mensagem_erro = nav.find_elements(By.CLASS_NAME, 'message_errorToHtml')
            texto_erro = mensagem_erro[0].text
            confirm = WebDriverWait(nav,60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="confirm"]')))
            time.sleep(2)
            confirm.click()
            
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
    
def requisitando(nav,rec,qtd,tipo_requisicao,requisitante_matricula,ccusto_text,observacao_text):

    nav.switch_to.default_content()
  
    #menu
    try:
        menu_innovaro_1(nav)
        print('Menu aberto')
    except TimeoutException:
        print('Erro ao clicar no menu')
        return 'Erro ao clicar no menu'
    time.sleep(3)
    
    #Clicando em solicitação de transferencia entre depositos
    lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
    time.sleep(0.5)
    click_producao = test_list.loc[test_list[0] == 'Requisições'].reset_index(drop=True)['index'][0]
    lista_menu[click_producao].click()
    time.sleep(3)
    
    #Mudando de iframe
    iframes(nav)
    
    #Clicando em Mudar visualização
    try:
        print('clicando em mudar visualização')
        mudar_visualizacao=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]')))
        mudar_visualizacao.click()
    except TimeoutException:
        print('Erro ao mudar visualização')
        return 'Erro ao mudar visualização'
    time.sleep(.5)
    
    #Clicando em insert
    try:
        print('clicando em insert')
        insert=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]')))
        insert.click()
    except TimeoutException:
        print('Erro ao da insert')
        return 'Erro ao da insert'
    time.sleep(.5)
    
    #inputando classe
    try:
        print('Classe de recurso')
        classe_recurso=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[4]/table/tbody/tr/td[1]')))
        classe_recurso.click()
        
        classe_recurso_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[1]/td[4]/table/tbody/tr/td[1]/input')))
        classe_recurso_input.clear()
        time.sleep(2)
        classe_recurso_input.send_keys(tipo_requisicao)
        classe_recurso_input.send_keys(Keys.TAB)
    except TimeoutException:
        print(f'Erro ao inputar classe de recurso')
        return 'Erro ao inputar classe de recurso'
    time.sleep(.5)
    
    #inputando requisitante
    try:
        print('Requisitante')
        requisitante=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]')))
        requisitante.click()
        
        requisitante_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
        requisitante_input.clear()
        time.sleep(2)
        requisitante_input.send_keys(requisitante_matricula)
        requisitante_input.send_keys(Keys.TAB)  
    except TimeoutException:
        print(f'Erro ao inputar Requisitante')
        return 'Erro ao inputar Requisitante'
    time.sleep(.5)
    
    #inputando ccusto
    try:
        print('inputando ccusto')
        ccusto=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td[1]')))
        ccusto.click()
        
        ccusto_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[3]/td[4]/table/tbody/tr/td[1]/input')))
        ccusto_input.clear()
        time.sleep(2)
        ccusto_input.send_keys(ccusto_text)
        ccusto_input.send_keys(Keys.TAB)       
    except TimeoutException:
        print(f'erro ao inputar ccusto')
        return 'Erro ao inputar ccusto'
    time.sleep(.5)
    
    #inputando recurso
    try:
        print('inputando recurso')
        recurso=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]')))
        recurso.click()
        
        recurso_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')))
        recurso_input.clear()
        time.sleep(2)
        recurso_input.send_keys(rec)
        recurso_input.send_keys(Keys.TAB)       
    except TimeoutException:
        print(f'Erro ao inputar recurso')
        return 'Erro ao inputar recurso'
    time.sleep(.5)
    
    #inputando quantidade
    try:
        print('inputando quantidade')
        quantidade=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[7]/td[3]/table/tbody/tr/td[1]')))
        quantidade.click()
        
        quantidade_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[7]/td[3]/table/tbody/tr/td[1]/input')))
        quantidade_input.clear()
        quantidade_input.send_keys(qtd)
        quantidade_input.send_keys(Keys.TAB)       
    except TimeoutException:
        print(f'Erro ao inputar quantidade')
        return 'Erro ao inputar quantidade'
    time.sleep(.5)

    #inputando quantidade
    try:
        print('inputando Observação')
        observacao=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td/table/tbody/tr[9]/td[2]/table/tbody/tr/td[1]/textarea')))
        observacao.click()
        observacao.clear()
        observacao.send_keys(observacao_text)
        observacao.send_keys(Keys.TAB)         
    except TimeoutException:
        print(f'Erro ao inputar quantidade')
        return 'Erro ao inputar quantidade'
    time.sleep(.5)
        
    #Clicando em confirmar (insert)
    try:
        print('clicando em confirmar')
        confirmar=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[4]')))
        confirmar.click()
    except TimeoutException:
        print('Erro ao confirmar')
        return 'Erro ao confirmar'
    time.sleep(.5)
    
    #Clicando em Mudar visualização
    try:
        print('clicando em mudar visualização')
        mudar_visualizacao=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]')))
        mudar_visualizacao.click()
    except TimeoutException:
        print('Erro ao mudar visualização')
        return 'Erro ao mudar visualização'
    time.sleep(.5)
    
    #selecionando checkbox
    try:
        print('clicando em selecionar checkbox')
        checkbox=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/table/tbody/tr[1]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td[1]')))
        checkbox.click()
    except TimeoutException:
        print('Erro ao selecionar checkbox')
        return 'Erro ao selecionar checkbox'
    time.sleep(.5)
    
    #Clicando em aprovar
    # Clicando em aprovar
    try:
        print('Clicando em aprovar')

        button = WebDriverWait(nav, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="buttonsBar_grdRequisicoes"]/td[1]'))
        )

        # Usar JavaScript para clicar no botão
        nav.execute_script("arguments[0].click();", button)

        nav.switch_to.default_content()

        # Esperar pela mensagem de sucesso
        try:
            element = WebDriverWait(nav, 10).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Requisição aprovada com sucesso!')]"))
            )
            if element:
                print("Requisição aprovada com sucesso!")
        except Exception as e:
            print(f"Erro ao buscar o texto: {e}")
            return e

        # Clicar no botão de confirmação "Ok"
        confirmar = WebDriverWait(nav, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))
        )
        confirmar.click()
        iframes(nav)

    except TimeoutException:
        print('Erro ao aprovar')
        return 'Erro ao aprovar'

    time.sleep(1.5)
    
    #Clicando em baixar
    try:
        print('clicando em baixar')
        baixar=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="buttonsBar_grdRequisicoes"]/td[3]')))
        baixar.click()
        
        try:
            print('preenchendo classe movimentação de depósito')
        
            #classe movimentação de depósito        
            classe_mov_deposito=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdInfoBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]')))
            classe_mov_deposito.click()

            classe_mov_deposito_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdInfoBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input')))
            classe_mov_deposito_input.send_keys(Keys.CONTROL + 'A')
            time.sleep(2)
            classe_mov_deposito_input.send_keys('Movimentação de depósitos')
            classe_mov_deposito_input.send_keys(Keys.TAB)
            
            checkbox_classe_movimentacao=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="1"]')))
            checkbox_classe_movimentacao.click()
            
            ok_classe_movimentacao=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="buttonsBar_grLookup"]/td[1]')))
            ok_classe_movimentacao.click()
            time.sleep(1)
            
        except TimeoutException:
            print('Erro ao preencher classe movimentação de depósito')
            return 'Erro ao preencher classe movimentação de depósito'
        
        try:
            print('Preenchendo depósito')
            deposito=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdInfoBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]')))
            deposito.click()

            deposito_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdInfoBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')))
            deposito_input.send_keys(Keys.CONTROL + 'A')
            time.sleep(2)
            deposito_input.send_keys('central')
            deposito_input.send_keys(Keys.TAB)
            
        except TimeoutException:
            print('Erro ao preencher depósito')
            return 'Erro ao preencher depósito'
        
        try:
            print('Preenchendo data de movimentação')
            data_movimentacao=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdInfoBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]')))
            data_movimentacao.click()

            data_movimentacao_input=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdInfoBaixa"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')))
            data_movimentacao_input.send_keys(Keys.CONTROL + 'A')
            time.sleep(2)
            # data_movimentacao_input.send_keys(datetime.datetime.now().date().strftime('%d/%m/%Y'))
            data_movimentacao_input.send_keys(datetime.datetime.now().date().strftime("%d/%m/%Y"))
            data_movimentacao_input.send_keys(Keys.TAB)
            time.sleep(2)
        except TimeoutException:
            print('Erro ao preencher data de movimentação')
            return 'Erro ao preencher data de movimentação'
        
        nav.switch_to.default_content()
        confirmar_baixa=WebDriverWait(nav,10).until(EC.element_to_be_clickable((By.XPATH,'/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]')))
        confirmar_baixa.click()
        iframes(nav)
        # Carregando ao clicar no MENU
        WebDriverWait(nav,60).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="grdRequisicoes"]/tbody/tr[1]/td[1]/table/tbody/tr[10]/td[7]/div')))
            
    except TimeoutException:
        print('Erro ao baixar')
        return 'Erro ao baixar'
    time.sleep(1.5)
    
    # Clicando em gravar
    try:
        print('clicando em gravar')
        nav.switch_to.default_content()
        gravar = WebDriverWait(nav, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[2]'))
        )
        gravar.click()
        time.sleep(2)

        confirmar_gravacao = WebDriverWait(nav, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="answers_0"]'))
        )
        confirmar_gravacao.click()

        # Tentativa de captura da mensagem de erro, se existir
        try:
            error_element = WebDriverWait(nav, 5).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'message_errorToHtml')]"))
            )
            if error_element:
                error_message = error_element.text
                print(f"Erro encontrado: {error_message}")

                # Tentar fechar a mensagem de erro
                try:
                    print("tentando")
                    confirm_button = WebDriverWait(nav, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))
                    )
                    confirm_button.click()
                    print("Mensagem de erro fechada.")
                except Exception as e:
                    print(f"Erro ao tentar fechar a mensagem de erro: {e}")

                return error_message  # Retorna a mensagem de erro e interrompe o processo

        except TimeoutException:
            print("Nenhuma mensagem de erro encontrada, prosseguindo...")
            confirm_button = WebDriverWait(nav, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="confirm"]'))
                )
            confirm_button.click()

        # Prosseguir se não houver erro
        print("Gravação concluída com sucesso.")

    except TimeoutException:
        print('Erro ao gravar')
        return 'Erro ao gravar'
    time.sleep(1)

    # Fechar aba
    try:
        print('clicando em fechar aba')
        time.sleep(2)
        fechar_aba = WebDriverWait(nav, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/table/tbody/tr/td[1]/table/tbody/tr/td[4]'))
        )
        fechar_aba.click()
        time.sleep(1)
    except TimeoutException:
        print('Erro ao fechar aba')
        return 'Erro ao fechar aba'
    time.sleep(.5)

    return 'OK'
        
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