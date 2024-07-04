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
rec = '028844'
# 034953
# 110012

qtd=1  

for i in range(100): 
    transferindo(nav,dep_origem,dep_destino,rec,qtd)