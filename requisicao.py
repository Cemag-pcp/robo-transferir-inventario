from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import time

from utils import *

while True:
    
    # acessando site
    link = "https://hcemag.innovaro.com.br/sistema"
    nav = webdriver.Chrome()
    nav.maximize_window()
    nav.get(link)

    # login e senha
    login(nav)

    time.sleep(5)
    # menu innovaro
    menu_requisicao(nav)

    # Conectar com a planilha google
    dep_origem = 'Almox central'
    dep_destino = 'mat fora uso'
    rec = '028844'
    # 034953
    # 110012

    qtd=1  

    #LER A PLANILHA

    for i in range(100): 
        try:
            requisitando(nav)   
        except:
            pass
        