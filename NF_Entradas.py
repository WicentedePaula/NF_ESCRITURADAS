from subprocess import Popen
from pywinauto import Application
import RepositorioDAO
import pyautogui
import FuncoesAuxiliares
import pyperclip
import os
import cv2
import time
from decimal import Decimal


class Digitacao:
   
    
    def Entradas(self,dta1, dta2):
        pyautogui.PAUSE=0
        
        varExecuteDAO = RepositorioDAO.DAO()
        varQueryLojas ="select nroempresa, empresa from CONSINCO.dim_empresa where nroempresa not in (99, 800, 986, 987, 989, 999, 10, 13, 20, 25, 29)  order by NROEMPRESA"
        #varQueryLojas ="select nroempresa, empresa from CONSINCO.dim_empresa where nroempresa not in (99, 800, 986, 987, 989, 999, 13, 20, 25,29) AND NROEMPRESA > 21 order by NROEMPRESA"
       
        varFuncao = FuncoesAuxiliares.Funcao_Apoio()
        varJn_acerto_operador="C:\\Projetos_Python\\TESOURARIA\\img\\janelas\\jn_acerto_operador.png"
            
     
        vlrOutros_Sistema = Decimal(0.0)
        vlrOutros_Tabela_String ="0"
        srtDiferenca= Decimal(0.0)

        #Sequencia de atalhos que abre a tela Acerto de Operador.
        pyautogui.press("ctrl")
        pyautogui.sleep(1)
        pyautogui.press("alt")
        pyautogui.sleep(1)
        pyautogui.press("r")
        pyautogui.sleep(1)

        pyautogui.press("enter")
        pyautogui.sleep(1)
        pyautogui.press("enter")
                                                                                                                            
        varFuncao.aguardar_janela_por_imagem(varJn_acerto_operador, "Acerto de Operador")
                         
        varLojas=varExecuteDAO.executaQuery(varQueryLojas)
        #LOOP DAS LOJAS
        for row in varLojas:
            lj=row[0]
            strLoja = str(lj)
            nm_lj = row[1]

            validacaoLojas =f"C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj{strLoja}.csv"
            validacaoxlsx =f'C:/Projetos_Python/TESOURARIA/arquivos/download/movimentolj{strLoja}.xlsx'
