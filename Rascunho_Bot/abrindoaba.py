from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import requests
from pathlib import Path
import glob
import pandas as pd
import time


NumChamado = '13926'


# DEF para chamado
def leiaInt(msg):
	ok = False
	valor = 0 
	while True:
		NumChamado = str(input(msg))
		if NumChamado.isnumeric():
			valor = int(NumChamado)
			ok = True
		else:
			print('Erro! Digite um chamado valido!\n Digite novamente')
		if ok:
			break
	return valor



driver = webdriver.Chrome(r"C:\Users\thiago.asouza\AppData\Local\Programs\Python\Python311\chromedriver.exe")

driver.get("https://moriah.qualitorsoftware.com/login.php")


#Variavel com o css selector do usuario
xpath_usu = "#cdusuario"
campo_usuario = driver.find_element("css selector", xpath_usu)
#Preenchendo o login
campo_usuario.send_keys('thiago.asouza')

time.sleep(3)

#Criar campo para clicar fora, tem bloquei se não digitar a senha.
#Variavel com o css xpath da senha
css_selector_senha = "/html/body/div[3]/form/div[1]/div[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[2]/table/tbody/tr[4]/td/span/input"
campo_senha = driver.find_element("xpath", css_selector_senha)
campo_senha.click()
campo_senha.click()
campo_senha.click()
#Preenchendo a senha
campo_senha.send_keys('Thiagolm34.')


time.sleep(4)

#Clicando no botao de entrar
#variavel do botao entrar
bt_entrar = "/html/body/div[3]/form/div[1]/div[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[2]/table/tbody/tr[9]/td/button"
#definindo elemento
clicar_entrar =  driver.find_element("xpath", bt_entrar)
#Click 
clicar_entrar.click()


driver.execute_script("window.open()")

time.sleep(5)

#ps = driver.find_element("xpath","/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/textarea").click()

#time.sleep(5)

driver.switch_to.window(driver.window_handles[1])

# Carrega a nova aba
driver.get(f'https://moriah.qualitorsoftware.com/html/hd/hdchamado/cadastro_chamado.php?cdchamado={NumChamado}')

time.sleep(20)

# Fecha a aba
#driver.find_element("tag name",'body').send_keys(Keys.CONTROL + 'w') 


driver.close()