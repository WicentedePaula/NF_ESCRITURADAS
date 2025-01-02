from pywinauto import Application
import RepositorioDAO
import pyautogui
import FuncoesAuxiliares
import os
from datetime import datetime
import time
import Captura_Sped
import Ndd_Web
import Entrada


if __name__  == '__main__':
    print('EXECUTANDO')
    
    #pyautogui.PAUSE=1.5
    pyautogui.FAILSAFE=False

    varExecuteDAO = RepositorioDAO.DAO()
    varFuncao = FuncoesAuxiliares.Funcao_Apoio() 
    sped = Captura_Sped.Sped_Emp()
    ndd = Ndd_Web.NDD()
    varEntrada = Entrada.Entrada()

    #Variaveis tela de login C5
    varImgLoginC5 ="C:\\Projetos_Python\\NF_ESCRITURADAS\\Img\\Janelas\\jn_login_consinco.png"
    varModuloFiscal ="C:\\C5Client\\Fiscal\\Fiscal.exe \n"
    varJanelaFiscal = "C:\\Projetos_Python\\NF_ESCRITURADAS\\img\\Janelas\\jn_JanelaFiscal.png"
    varJanelaEntrada ="C:\\Projetos_Python\\NF_ESCRITURADAS\\img\\Janelas\\jn_entrada.png"
    
    #Variaveis de atenticacao c5    
    varNome="automacao2"
    varSenha= os.getenv('pw_automacao')

    ########################## Acessando a consinco ###############################################################
     
    pyautogui.hotkey("win","r")
    pyautogui.sleep(2)
    pyautogui.typewrite(varModuloFiscal,interval=0.015)
    
    varFuncao.aguardar_janela_por_imagem(varImgLoginC5,"Janela Login")

    time.sleep(2)
    #iniciando aplicativo da consinco
    app = Application().connect(class_name="Gupta:Dialog")
    pyautogui.sleep(2)
    #O formulário em questão não tem título, sendo assim foi identificado desta forma.
    dlg = app.window() 

    #Digitando os dados no formulário de login
    dlg['Edit4'].click_input() # Caixa de texto usuário
    dlg['Edit4'].type_keys(varNome) #Escrevendo na caixa de texto usuário
    pyautogui.sleep(1)
    dlg['Edit5'].click_input() # Click dentro da caixa de texto senha
    dlg['Edit5'].type_keys(varSenha) #Escrevendo na caixa de texto senhaouraria.exe 
    dlg['Button0'].click_input() # Click no butão entrar

    varFuncao.aguardar_janela_por_imagem(varJanelaFiscal,"Janela Fiscal")

    ######################### Gera o Relatório de entrada pela consinco, e chama os demais módulos #############################################
    varEntrada.RelatorioEntrada()
                                                                                                                      
     
    print("CONCLUIDO")
    
    

   
   
  