from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import keyboard
import pyautogui
import FuncoesAuxiliares

class Skyione:

    def acessoSkyinoneConsinco(self):

        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico)
        navegador.get('https://arcomix.autosky.cloud/')
        navegador.maximize_window()
        varFuncao = FuncoesAuxiliares.Funcao_Apoio()

        #Clicando na caixa de texto login
        text_usuario = WebDriverWait(navegador, 3).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/div/div[3]/div[1]/div/div[1]/div/div/form/div/div[1]/div/div/input'))
        )
        text_usuario.send_keys("vicente.silva@arcomix.com.br")


        #Clicando na caixa de texto senha
        text_senha = WebDriverWait(navegador, 3).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="root"]/div/div/div[3]/div[1]/div/div[1]/div/div/form/div/div[2]/div/div/input'))
        )
        text_senha.send_keys("Vic190710")


        # Aguarde até que o botão que você deseja clicar esteja presente
        botao_login = WebDriverWait(navegador, 15).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div/div[3]/div[1]/div/div[1]/div/div/form/div/div[5]/button/span[1]'))  # Substitua pelo ID correto do botão
        )
        # Execute o clique
        botao_login.click()

        # Aguarde até que o botão que você deseja clicar esteja presente
        botao_login = WebDriverWait(navegador, 3).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="root"]/div/div/div[3]/div[1]/div/div[1]/div/div/form/div/div[5]/button/span[1]'))  # Substitua pelo ID correto do botão
        )
        # Execute o clique
        botao_login.click()

       # varFuncao.aguardar_janela_por_imagem()


        pyautogui.sleep(10)  