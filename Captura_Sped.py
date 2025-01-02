from subprocess import Popen
from pywinauto import Application
from datetime import datetime
import RepositorioDAO
import pyautogui
import FuncoesAuxiliares
import pyperclip
import os
import cv2
import time
import shutil

class Sped_Emp:

    def TranferefeArquivo_Sped(self, path_modulo_consico, varPathPastas_sped, numero_nome):
        varFuncao = FuncoesAuxiliares.Funcao_Apoio()
        sucesso = "0" #Zero para processamento ok

        #Acessando caminho das pastas do explorer no servidor da skyone e colar o arquivo sped no cominho quardado no path_modulo_consico.
        pyautogui.press("ctrl")
        pyautogui.hotkey("win","r")
        pyautogui.sleep(1)
        pyautogui.typewrite("Documents \n",interval=0.015)
        pyautogui.sleep(1)
        pyautogui.press("ctrl")
        pyautogui.sleep(1)
        pyautogui.hotkey("ctrl","l")
        pyautogui.typewrite(path_modulo_consico,interval=0.015)
        pyautogui.sleep(2)
       
        try:
         # Realiza a cópia
         shutil.copy(varPathPastas_sped, "\\\\10.102.227.2\\consinco2\\importacao\\sped")
         pyautogui.sleep(1)
        
         # Verifica se o arquivo foi copiado com sucesso
         arquivo_copiado = os.path.join("\\\\10.102.227.2\\consinco2\\importacao\\sped", os.path.basename(varPathPastas_sped))

         if not os.path.exists(arquivo_copiado):
                
                sucesso ="1" #Falha ao encontrar o arquivo
                varFuncao.GeraLogsInfo(f"LOJA : {numero_nome} Falha ao copiar arquivo.")
                return sucesso
         
        except Exception as e:

            sucesso ="1" #Falha ao encontrar o arquivo
            print(f"Ocorreu um erro ao copiar o arquivo: {e}")
            varFuncao.GeraLogsInfo(f"LOJA : {numero_nome} Erro ao capturar arquivo !!!")
            return sucesso
        
        pyautogui.sleep(10)

        pyautogui.press("ctrl")
        pyautogui.press("alt")
        pyautogui.sleep(1)
        pyautogui.press("f")
        pyautogui.sleep(1)
        pyautogui.press("f")

        print("############################## Arquivo Sped, copiado para a área do servidor com sucesso ############################################################")
































    """

    ano_atual =  datetime.now().strftime("%Y")
    #mes_ano = datetime.now().strftime("%m-%Y")
    mes_ano ="09-2024" 
    path_modulo_consico = "\\\\10.102.227.2\consinco2\importacao\sped \n"
    varExecuteDAO = RepositorioDAO.DAO()
    
      
    varQueryLojas ="select nroempresa, empresa from CONSINCO.dim_empresa where nroempresa not in (99, 800, 986, 987, 989, 999, 10, 13, 20, 25, 29)  order by NROEMPRESA"
    
        
    varQueryLojas = varExecuteDAO.executaQuery(varQueryLojas)
    for row in varQueryLojas:
        nrlj=row[0]
        nome_nm_lj = row[1]
        
        numeroLoja = str(nrlj) #Numero da loja string
        numero_nome = str(nome_nm_lj) #Nome e numero da loja string
        
        #varPathPastas_sped = f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\LOJAS\\Loja {numero_nome}\\Federal\\{ano_atual}\\{mes_ano}\\nome_do_arquivo.extencao"
        varPathPastas_sped =f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA{numeroLoja}\\{mes_ano}\\SPED_EMP{numeroLoja}.txt"
                           
        #Acessando caminho das pastas do explorer no servidor da skyone e colar o arquivo sped no cominho quardado no path_modulo_consico.
        pyautogui.hotkey("win","r")
        pyautogui.sleep(1)
        pyautogui.typewrite("Documents \n",interval=0.015)
        pyautogui.sleep(1)
        pyautogui.press("ctrl")
        pyautogui.sleep(1)
        pyautogui.hotkey("ctrl","l")
        pyautogui.typewrite(path_modulo_consico,interval=0.015)
        pyautogui.sleep(2)

        #NESTE PONTO VOU TER QUE COLAR O ARQUIVO QUE ESTA NO CAMINHO FEDERAL PARA O  CAMINHO : \\10.102.227.2\consinco2\importacao\sped
                    #Origem             Destino         
        shutil.copy(varPathPastas_sped, "\\\\10.102.227.2\consinco2\importacao\sped")
        
    """

        
        

    


