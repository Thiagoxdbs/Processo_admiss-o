from selenium import webdriver
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.keys import Keys
import time


#login_usu = str(input('Digite seu login: '))
#senha_usu = str(input('Digite sua senha: '))

driver = webdriver.Chrome("C:/chromedriver.exe")
driver.get("https://moriah.qualitorsoftware.com/login.php")

time.sleep(10)


#Login
#xpath
usuario_path = '/html/body/div[3]/form/div[1]/div[2]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/div[2]/table/tbody/tr[2]/td/span/input'
#identifica e retorna os elementos
usuario = driver.find_element("xpath", usuario_path)
usuario.send_keys("Rantarolm34.2")


#Senha
#xpath
senha_path = '#cdsenha'
#identifica e retorna os elementos
input_senha = driver.find_element('css selector', senha_path)
input_senha.send_keys("Rantarolm34.")


time.sleep(10)
driver.quit()

