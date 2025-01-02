from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import keyboard
import pyautogui
import FuncoesAuxiliares


class NDD:

    def capturaNDD(self):
         #MODULO WEB -- INICIO
              
        #download_folder="C:\\Projetos_Python\\NF_ESCRITURADAS\\arquivos\\NDD\\" #Defina o caminho desejado
        varFuncao = FuncoesAuxiliares.Funcao_Apoio()

        # download.default_directory
        chrome_options = Options()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": "C:\\Projetos_Python\\NF_ESCRITURADAS\\arquivos\\NDD",
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
            
        })

        #Chamando o browser e redirecionando para a página
        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico, options=chrome_options)
        navegador.get("https://nddconnect.e-datacenter.nddigital.com.br/Auth/login")
        navegador.maximize_window()

        
        #Clicando na caixa de texto login
        text_usuario = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="login__input--login"]'))
        )
        text_usuario.send_keys("arcomix")

        pyautogui.sleep(1)

        #Clicando na caixa de texto senha
        text_senha = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH,'//*[@id="login__input--password"]'))
        )
        text_senha.send_keys("arcomix")


        # Aguarde até que o botão que você deseja clicar esteja presente
        botao_login = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/ui-view/login/div/div/div[2]/div[2]/div/login-form/div[1]/form/fieldset/div[3]/div/button'))  # Substitua pelo ID correto do botão
        )
        # Execute o clique no botão de login
        botao_login.click()
        
        #pyautogui.sleep(15)

        try:
            elemento = WebDriverWait(navegador, 30).until(
                EC.presence_of_element_located((By.ID, "iframe"))
            )
            print("Elemento encontrado.")
        except TimeoutException:
            print("Elemento não foi encontrado dentro do tempo limite.")

                       
        #Trocando Para o iframe correto
        navegador.switch_to.frame("iframe")

        #Clicando Resumo de incidentes
        botao_resumo_incidente = WebDriverWait(navegador, 20).until(
            EC.element_to_be_clickable((By.XPATH,'//*[@id="cmdNavigation"]/div')) 
        )
        botao_resumo_incidente.click()


        #Clicando Gestao de entrada
        botao_gestao_entrada = WebDriverWait(navegador, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="cmdNavigation"]/div/ul/li[2]/span'))  
        )
        botao_gestao_entrada.click()


        #Clicando confronto contábil
        botao_confronto_contabil = WebDriverWait(navegador, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="cmdNavigation"]/div/ul/li[2]/ul/li[2]/ul/li[2]/a'))  
        )
        botao_confronto_contabil.click()

        #botao download
        botao_download = WebDriverWait(navegador, 20).until(
        
        EC.element_to_be_clickable((By.XPATH,'//*[@id="wgButtonDownload"]/button/ul/li[1]/span'))  #cmdNavigation
        )
        botao_download.click()

        varFuncao.monitorar_pasta()
               
        #keyboard.wait('space')
        navegador.close()
       

    