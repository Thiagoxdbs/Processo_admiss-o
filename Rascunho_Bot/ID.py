from selenium import webdriver
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import time


#Login = str(input('Digite seu login: '))
#Senha = str(input('Digite sua senha: '))



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


# Identificar o chamado
NumChamado = leiaInt('Digite o numero do chamado: ')

#caminho para o chromedriver
driver = webdriver.Chrome(r"C:\Users\thiago.asouza\AppData\Local\Programs\Python\Python311\chromedriver.exe")
#abrindo navegador
driver.get("https://moriah.qualitorsoftware.com/login.php")

#Tempo para processar a pagina
time.sleep(4)

#Variavel com o css selector do usuario
xpath_usu = "#cdusuario"
campo_usuario = driver.find_element("css selector", xpath_usu)
#Preenchendo o login
campo_usuario.send_keys('thiago.asouza')

time.sleep(15)

#Criar campo para clicar fora, tem bloquei se não digitar a senha.
#Variavel com o css xpath da senha
css_selector_senha = "/html/body/div[3]/form/div[1]/div[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[2]/table/tbody/tr[4]/td/span/input"
campo_senha = driver.find_element("xpath", css_selector_senha)
campo_senha.click()
campo_senha.click()
campo_senha.click()
#Preenchendo a senha
campo_senha.send_keys('Rantarolm34.')

time.sleep(5)

#Clicando no botao de entrar
#variavel do botao entrar
bt_entrar = "/html/body/div[3]/form/div[1]/div[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[2]/table/tbody/tr[9]/td/button"
#definindo elemento
clicar_entrar =  driver.find_element("xpath", bt_entrar)
#Click 
clicar_entrar.click()

time.sleep(40)

# Abre uma nova aba
driver.find_element("tag name",'body').send_keys(Keys.CONTROL + 't') 

# Carrega a nova aba
driver.get(f'https://moriah.qualitorsoftware.com/html/hd/hdchamado/cadastro_chamado.php?cdchamado=13548{NumChamado}')

# Pesquisar o chamado
#variavel da pesquisa
#br_pesquisa = "/html/body/div[1]/table/tbody/tr[1]/td/table/tbody/tr[1]/td/div/div[2]/div/table/tbody/tr/td/div[1]/input"
#variavel do botao pesquisa
#bt_pesquisa = "/html/body/div[1]/table/tbody/tr[1]/td/table/tbody/tr[1]/td/div/div[2]/div/table/tbody/tr/td/div[2]/i"
#definindo elemento da pequisa
#definindo_pesquisar =  driver.find_element("xpath", br_pesquisa)
#preenchendo chamado
#definindo_pesquisar.send_keys(NumChamado)
#definindo elemento do botao de pequisar
#clicar_pesquisar =  driver.find_element("xpath", bt_pesquisa)
#Click 
#clicar_pesquisar.click()

#time.sleep(10)

# Clicar em exibir chamado
#Variavel do exibir chamado
#ck_exibir = "#gridChamadosRow_0 > td:nth-child(2)"
#definindo elemento do exibir chamado
#exibir_chamado = driver.find_element("css selector", ck_exibir)
#Click 
#exibir_chamado.click()

 


time.sleep(20)



  # Clicar CRIAÇÃO DE ACESSOS
  # Buscar data de inicio guardar em um id
  # Buscar departamento guardar em um id 
  # Buscar cpf guardar em um id 
  # Close navegador



