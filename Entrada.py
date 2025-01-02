from pywinauto import Application
import RepositorioDAO
import pyautogui
import FuncoesAuxiliares
import os
from datetime import datetime
import time
import Captura_Sped
import Ndd_Web
import csv
import pandas as pd
from dateutil.relativedelta import relativedelta


class Entrada:

    def RelatorioEntrada(self):

        varFuncao = FuncoesAuxiliares.Funcao_Apoio()
        varExecuteDAO = RepositorioDAO.DAO()
        sped = Captura_Sped.Sped_Emp()
        ndd = Ndd_Web.NDD()

        varJanelaFiscal = "C:\\Projetos_Python\\NF_ESCRITURADAS\\img\\Janelas\\jn_JanelaFiscal.png"
        path_modulo_consico = "\\\\10.102.227.2\\consinco2\\importacao\\sped \n"
      
        ano_atual =  datetime.now().strftime("%Y")
        mes_ano = (datetime.now() - relativedelta(months=1)).strftime("%m-%Y")
        #mes_ano ="11-2024" 
      
        #Abrindo Janela de emissão de relatório de Entrada C5
        pyautogui.press("ctrl")
        pyautogui.press("alt")
        pyautogui.sleep(1)
        pyautogui.press("r")
        pyautogui.sleep(1)
        pyautogui.press("enter")
        pyautogui.sleep(1)
        pyautogui.press("enter")
        

        varFuncao.aguardar_janela_por_imagem(varJanelaFiscal,"varJanelaEntrada")


        #varQueryLojas ="select nroempresa, empresa from CONSINCO.dim_empresa where nroempresa not in (99, 800, 986, 987, 989, 999, 10, 13, 20, 25, 29) order by NROEMPRESA"
        #varQueryLojas ="select nroempresa, empresa from CONSINCO.dim_empresa where nroempresa in (1,2,3,5,24,30) order by NROEMPRESA"
        varQueryLojas ="select nroempresa, empresa from CONSINCO.dim_empresa where nroempresa in (8,14,17,18,27) order by NROEMPRESA"
        varQueryLojas = varExecuteDAO.executaQuery(varQueryLojas)
        for row in varQueryLojas:
            nrlj=row[0]
            nome_nm_lj = row[1]
            
            numeroLoja = str(nrlj) #Numero da loja string
            numero_nome = str(nome_nm_lj) #Nome e numero da loja string
            varArquivoEntrada =f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA{numeroLoja}\\{mes_ano}\\entradaLoja{numeroLoja}.txt"
            varArquivoNDD = f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA{numeroLoja}\\{mes_ano}\\arquivoNDD_Loja{numeroLoja}.csv"
            varPathPastas_sped =f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA{numeroLoja}\\{mes_ano}\\SPED_EMP{numeroLoja}.txt"
                                                                                                                                   
            pyautogui.press("ctrl")                                          
            pyautogui.hotkey("ctrl","shift","t") 
     
            varFuncao.AguardaAberturaJanela("Empresas")

            #Conectando em empresas administradoras
            appEmpre = Application().connect(title_re=".*Empresas.*", class_name="Gupta:Dialog") 
            dlgEmpresas = appEmpre.window(class_name="Gupta:Dialog")

            dlgEmpresas['Edit0'].click_input() # Clicando na data do campo numeroLoja
            pyautogui.sleep(1)
            pyautogui.keyDown("ctrl") #Selecionando a informacao
            pyautogui.keyDown("a")
            pyautogui.keyDown("a")
            pyautogui.keyDown("ctrl")
            dlgEmpresas['Edit0'].type_keys("{DELETE}")
            dlgEmpresas['Edit0'].type_keys(numeroLoja)
            dlgEmpresas['Button0'].click_input() # Clicando em ok para na janelas empresas.

            dtMesAnterior = varExecuteDAO.executaQuery("SELECT TO_CHAR(TRUNC(SYSDATE, 'MM') - INTERVAL '1' MONTH, 'DD/MM/YYYY') AS primeiro_dia_mes_anterior FROM dual") 
            dt_Atual= varExecuteDAO.executaQuery("SELECT TO_CHAR(SYSDATE, 'DD/MM/YYYY') AS data_dia_Atual FROM dual") 
            #dtMesAnterior ="01/09/2024" 
            #dt_Atual ="31/10/2024"
          
            ############################### Gerando Relatório de entrada C5 #################################################
            appEntrada = Application().connect(title_re=".*Fiscal.*")
            guptamdiframeEntrada = appEntrada[u'Gupta:MDIFrame']
            guptamdiframeEntrada.wait('ready')
            combobox_entrada = guptamdiframeEntrada.ComboBox
            combobox_entrada.select(u'Intervalo')  
            guptamdiframeEntrada[u'Edit0'].type_keys(dtMesAnterior)#[u'Edit', u'Edit0', u'Edit1'] 
            guptamdiframeEntrada[u'aEdit'].type_keys(dt_Atual)#[u'Edit2', u'aEdit'] 

            pyautogui.sleep(1) 
            guptamdiframeEntrada[u'Button10'].click_input()#Clicando no botão imprimindo em arquivo  [u'Button10', u'Imprimir em Arquivo', u'Imprimir em ArquivoButton']
            pyautogui.sleep(2)

            ##### Salvando o arquivo no local desejado
            appSalvarArquivo = Application().connect(title_re=".*Caminho e nome.*")
            windowSalvar = appSalvarArquivo.Dialog
            windowSalvar.Wait('ready')
            comboboxArquivo = windowSalvar.ComboBox2
            pyautogui.sleep(1)
            comboboxArquivo.ClickInput()
            pyautogui.sleep(1)
            windowSalvar[u'&Nome:Edit'].type_keys(varArquivoEntrada) 
            pyautogui.sleep(2)
            windowSalvar[u'Sa&lvar'].click_input() 
            pyautogui.sleep(1.5)

            #Se o arquivo ja existir o codigo abaixo sera executado
            jnErro = varFuncao.check_window_exists("Confirmar Salvar como")
            if jnErro == True:
                appDlgConf_Exclusao = Application().connect(title_re=".*Confirmar Salvar como.*")
                window = appDlgConf_Exclusao.Dialog
                window.Wait('ready')
                button = window[u'Button1']
                button.click_input()
            
            varFuncao.AguardaAberturaJanela("Atenção")

            appDlgConf_Exclusao = Application().connect(title_re=".*Atenção.*")
            window = appDlgConf_Exclusao.Dialog
            window.Wait('ready')
            button = window[u'&OKButton'] #[u'&OK', u'Button', u'&OKButton']
            button.click_input()

            #guptamdiframeEntrada.close()

             ########################## Iniciando transferencia de arquivo sped  ###################################################
            retorno = sped.TranferefeArquivo_Sped(path_modulo_consico, varPathPastas_sped, numero_nome)

            if retorno == 1:
                
                print("Retorno do arquivo sped com erro")
                continue
                                   
            
            ########################## Iniciando a Tranferencia do arquivo ndd #####################################################
            #MODULO WEB -- INICIO
            ndd.capturaNDD()
            

            ######################### Tranferido o arquivo do ndd para a pasta do confronto Sped
            caminho_arquivoNDD =f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA{numeroLoja}\\{mes_ano}\\"
            arquivoNDD =f"arquivoNDD_Loja{numeroLoja}.csv"
                               
            varFuncao.transferirArquivo(caminho_arquivoNDD,arquivoNDD)

            ######################### Cruzando relatórios de entrada com relatórios NDD ############################################
            resultado = varFuncao.confronto_NDD(varArquivoNDD, varArquivoEntrada)

            ######################### Cruzamento relatórios de retorno de cruzamento com o de Notas Pendentes no ato da Entrada ####
            pendenciaAtoEntrega = varExecuteDAO.NF_pendentes_Ato_da_Entrega(numeroLoja, dtMesAnterior, dt_Atual)

            for vlr in resultado:
             print(vlr)

           
            resultado_formatado = [linha.split(";") for linha in resultado]

            # Criar um DataFrame a partir da lista
            df = pd.DataFrame(resultado_formatado, columns=["FORNECEDOR", "NUMERO", "CNPJ", "CHAVE", "EMISSÃO", "VALOR CONECT", "CFOP", "STATUS"])

            # Definir o caminho do arquivo xlsx
            caminho_arquivo = f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA{numeroLoja}\\{mes_ano}\\resultado_confronto{numeroLoja}.xlsx"

            # Salvar o DataFrame em um arquivo Excel
            df.to_excel(caminho_arquivo, index=False, engine='openpyxl')

            print(f"Arquivo salvo com sucesso em {caminho_arquivo}")
                                                                                
            """
            resultado_formatado = [linha.split(";") for linha in resultado]
            #Salvar o DataFrame em um arquivo CSV
            caminho_arquivo = f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA{numeroLoja}\\{mes_ano}\\resultado_confronto{numeroLoja}.csv"
            # Gravando a lista no arquivo CSV
            with open(caminho_arquivo, mode='w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                
                #Escrever o cabeçalho no CSV
                writer.writerow(["FORNECEDOR ;NUMERO ; CNPJ; CHAVE; EMISSÃO; STATUS"])


                # Escrever uma linha no CSV (a lista como linha única)
                writer.writerows(resultado_formatado)

            print(f"Arquivo salvo com sucesso em {caminho_arquivo}")
            """
            
            