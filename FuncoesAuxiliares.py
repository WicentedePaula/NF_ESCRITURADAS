import pygetwindow as gw
import cv2
import pyautogui
import time
import datetime
import csv
from decimal import Decimal
import locale
import tkinter as tk
from tkinter import messagebox
import logging
from pywinauto import Application
import Calculos
import FuncoesAuxiliares
import pandas as pd
import RepositorioDAO
from openpyxl import load_workbook
import os
import shutil
import re



class Funcao_Apoio:

  
    #Retorna o título da Janela Ativa
    def GetScreenShot(self,caminhoArquivo, nmLoja):
        screenshot = pyautogui.screenshot()
        data_hora_atual = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        nome_arquivo = f'{nmLoja}_{data_hora_atual}.png'
        
        screenshot.save(f"{caminhoArquivo}{nome_arquivo}")


    def check_window_exists(self, window_name):
        time.sleep(2)
        # Obtém todas as janelas ativas
        windows = gw.getWindowsWithTitle(window_name)
        
        # Verifica se a lista de janelas não está vazia
        if windows:
            return True
        else:
            return False

    

    # Aguarda até a janela abrir.
    def AguardaAberturaJanela(self, janela_alvo):

     # Nome da janela que estou está esperando abrir
       
           # Loop para verificar continuamente se a janela alvo foi aberta
        while True:
            # Obtém todas as janelas ativas
            janelas = gw.getAllWindows()

            # Verifica se a janela alvo está entre as janelas ativas
            if any(janela_alvo in janela.title for janela in janelas):
                print("Janela alvo foi aberta!:")
                break  # Sai do loop quando a janela alvo for encontrada

            # Pausa por um curto período de tempo antes de verificar novamente
            time.sleep(1)  # Importe time se você ai

    def transferirArquivo(self, caminho_arquivoNDD, arquivoNDD):
        # Caminhos das pastas
        pasta_ndd = "C:\\Projetos_Python\\NF_ESCRITURADAS\\arquivos\\NDD\\"
       
        # Listar arquivos na pasta NDD
        arquivos = os.listdir(pasta_ndd)

        # Renomear e mover o arquivo
        if arquivos:
            arquivo_antigo = os.path.join(pasta_ndd, arquivos[0])  # Obter o único arquivo
            novo_nome = arquivoNDD  # Defina o novo nome desejado
            caminho_renomeado = os.path.join(caminho_arquivoNDD, novo_nome)

            # Mover e renomear o arquivo
            shutil.move(arquivo_antigo, caminho_renomeado)
            print(f"Arquivo renomeado para {novo_nome} e movido para a pasta de alteração.")
        else:
            Funcao_Apoio.GeraLogsInfo(f"Arquivo NDD {arquivoNDD}, Não foi encontrado " )
            print("Nenhum arquivo encontrado na pasta NDD.")

               
   
    # Testa se a janela existe baseado em imagens
    def aguardar_janela_por_imagem(self,imagem_janela, mensagem): # Passar o endereço da imagem salvo no computador
        
        while True:

            try:
                posicao = pyautogui.locateCenterOnScreen(imagem_janela, confidence=0.2)
                if posicao is not None:
                   return 0
          
            except Exception as e:
                print(f'Aguardando {mensagem}...', e)

    

    def show_popup(self,message):
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal
        messagebox.showinfo("Popup", message)
        root.destroy()

    
    def GeraLogsInfo(self, mensagem):

        caminho ="C:\\Projetos_Python\\NF_ESCRITURADAS\\arquivos\\logs\\Info\\"

        if not os.path.exists(caminho):
            os.makedirs(caminho)

        # Define o nome do arquivo com a data atual
        data_atual = datetime.datetime.now().strftime("%Y-%m-%d")
        nome_arquivo = f"info_processamento_{data_atual}.log"
        caminho_arquivo = os.path.join(caminho, nome_arquivo)

        # Abre o arquivo em modo append para adicionar logs
        with open(caminho_arquivo, "a") as arquivo_log:
            arquivo_log.write(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {mensagem}\n")



    def monitorar_pasta(self):
        caminho_pasta = "C:\\Projetos_Python\\NF_ESCRITURADAS\\arquivos\\NDD"
        """
        Monitora uma pasta até que ela contenha arquivos.
        Enquanto a pasta estiver vazia, a aplicação espera.
        
        """
        print(f"Monitorando a pasta: {caminho_pasta}")
        
        while True:
            # Lista os arquivos e diretórios na pasta
            conteudo_pasta = os.listdir(caminho_pasta)
            
            if conteudo_pasta:  # Se a pasta não estiver vazia
               
                break
            else:
                print("Pasta vazia, aguardando...")
                time.sleep(5)  # Aguarda 5 segundos antes de verificar novamente

    
    def confronto_NDD(self, arquivoNDD, ArquivoEntrada):

        arquivo_txt = ArquivoEntrada
        arquivo_csv = arquivoNDD

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)
        #pd. set_option ('display.max_colwidth ', None )

        # Ler o arquivo CSV com pandas
        df_csv = pd.read_csv(arquivo_csv, sep=';', encoding='ISO-8859-1')

        # Converter as colunas 12 e 14 para string
        coluna_csv_12 = df_csv.iloc[:, 12].astype(str).tolist()  # Razão Social
        try:
            coluna_csv_14 = df_csv.iloc[:, 14].astype(str).tolist()  # Número da Nota
        except:
            print("NRO_NOTA_VAZIO")
            coluna_csv_14 = []  # Evita erros se a coluna não existir

        coluna_csv_10 = df_csv.iloc[:, 10].astype(str).tolist()  # CNPJ
        coluna_csv_5 = df_csv.iloc[:, 5].astype(str).tolist()  # CHAVE
        coluna_csv_16 = df_csv.iloc[:, 16].astype(str).tolist()  # EMISSAO
        coluna_csv_18 = df_csv.iloc[:, 18].astype(str).tolist()  # Valor Conect
        coluna_csv_30 = df_csv.iloc[:, 30].astype(str).tolist()# cfop

        # Extrair as colunas situação (índice 4) e descrição (índice 9)
        coluna_situacao = df_csv.iloc[:, 4].astype(str).tolist() #Situacao
        coluna_descricao = df_csv.iloc[:, 7].astype(str).tolist()  # Descrição

        # Ler o arquivo TXT diretamente para um DataFrame
        df_txt = pd.read_csv(arquivo_txt, sep=';', encoding='ISO-8859-1', header=None)

        # Extrair a coluna 4 (como você fazia antes)
        lote_coluna_txt = df_txt.iloc[:, 4].tolist()  # Coluna 4 do arquivo TXT

        itens_presentes = set()
        nova_entrada = ""

        # Processar os dados
        for situacao, descricao, item_coluna_12, item_coluna_14, coluna_csv_10, coluna_csv_5, coluna_csv_16, coluna_csv_18, coluna_csv_30 in zip(coluna_situacao, coluna_descricao, coluna_csv_12, coluna_csv_14, coluna_csv_10, coluna_csv_5, coluna_csv_16, coluna_csv_18, coluna_csv_30):
            if situacao == 'Não Informado' and descricao == 'Autorizado':
                item_coluna_14_formatado = item_coluna_14.replace('.0', '')
                coluna_csv_30 = coluna_csv_30.replace('.0', '')

                # Verificar se o item da coluna 14 está presente no lote da coluna TXT
                if item_coluna_14_formatado.lstrip() in lote_coluna_txt:
                    nova_entrada = f"{item_coluna_12}; {item_coluna_14_formatado}; {coluna_csv_10}; {coluna_csv_5}; {coluna_csv_16}; {coluna_csv_18}; {coluna_csv_30} ; ESCRITURADA"
                    if nova_entrada not in itens_presentes:
                       #print("Número da nota incluída:" + item_coluna_14_formatado)
                        itens_presentes.add(nova_entrada)  # Adiciona ao conjunto
                else:
                    nova_entrada1 = f"{item_coluna_12}; {item_coluna_14_formatado}; {coluna_csv_10}; {coluna_csv_5}; {coluna_csv_16}; {coluna_csv_18}; {coluna_csv_30} ; NOTA COM PENDENCIA"
                    if nova_entrada1 not in itens_presentes:
                        itens_presentes.add(nova_entrada1)

        itens_presentes = list(itens_presentes)
        return itens_presentes


    def verificar_e_incluir_0001(self, resultado, nfPendente_entrada):

        listaResultante = []

        for item_principal in resultado[1:]:
            for item_secundaria in nfPendente_entrada:
                numeroPrincipal =str(item_principal[1]).lstrip()

                if numeroPrincipal == item_secundaria[11]:  # Compara os itens nos índices 1 da principal e 2 da secundária
                    listaResultante.append([item_principal[0], item_principal[1], item_principal[2], item_principal[3], item_principal[4], item_principal[5], item_principal[6], item_principal[7] ,item_secundaria[7]])  # Adiciona o item da coluna 7
                              
                                                                      

        return listaResultante


    def verificar_e_incluir(self, resultado, nfPendente_entrada):

        listaResultante = []

        for item_principal in resultado[1:]:
            for item_secundaria in nfPendente_entrada:
                numeroPrincipal =str(item_principal[1]).lstrip()

                if numeroPrincipal == item_secundaria[2]:  # Compara os itens nos índices 1 da principal e 2 da secundária
                    listaResultante.append([item_principal[0], item_principal[1], item_principal[2], item_principal[3], item_principal[4], item_principal[5], item_principal[6], item_principal[7] ,item_secundaria[7]])  # Adiciona o item da coluna 7
                                                                                   
                                                                      

        return listaResultante


    def confronto_NDD_bkp_atualizado(self, arquivoNDD, ArquivoEntrada):

        arquivo_txt = ArquivoEntrada
        arquivo_csv = arquivoNDD

        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)
        #pd. set_option ('display.max_colwidth ', None )

        # Ler o arquivo CSV com pandas
        df_csv = pd.read_csv(arquivo_csv, sep=';', encoding='ISO-8859-1')

        # Converter as colunas 12 e 14 para string
        coluna_csv_12 = df_csv.iloc[:, 12].astype(str).tolist()  # Razão Social
        try:
            coluna_csv_14 = df_csv.iloc[:, 14].astype(str).tolist()  # Número da Nota
        except:
            print("NRO_NOTA_VAZIO")
            coluna_csv_14 = []  # Evita erros se a coluna não existir

        coluna_csv_10 = df_csv.iloc[:, 10].astype(str).tolist()  # CNPJ
        coluna_csv_5 = df_csv.iloc[:, 5].astype(str).tolist()  # CHAVE
        coluna_csv_16 = df_csv.iloc[:, 16].astype(str).tolist()  # EMISSAO

        # Extrair as colunas situação (índice 4) e descrição (índice 9)
        coluna_situacao = df_csv.iloc[:, 4].astype(str).tolist() #Situacao
        coluna_descricao = df_csv.iloc[:, 7].astype(str).tolist()  # Descrição

        # Ler o arquivo TXT diretamente para um DataFrame
        df_txt = pd.read_csv(arquivo_txt, sep=';', encoding='ISO-8859-1', header=None)

        # Extrair a coluna 4 (como você fazia antes)
        lote_coluna_txt = df_txt.iloc[:, 4].tolist()  # Coluna 4 do arquivo TXT

        itens_presentes = set()
        nova_entrada = ""

        # Processar os dados
        for situacao, descricao, item_coluna_12, item_coluna_14, coluna_csv_10, coluna_csv_5, coluna_csv_16 in zip(coluna_situacao, coluna_descricao, coluna_csv_12, coluna_csv_14, coluna_csv_10, coluna_csv_5, coluna_csv_16):
            if situacao == 'Não Informado' and descricao == 'Autorizado':
                item_coluna_14_formatado = item_coluna_14.replace('.0', '')

                # Verificar se o item da coluna 14 está presente no lote da coluna TXT
                if item_coluna_14_formatado.lstrip() in lote_coluna_txt:
                    nova_entrada = f"{item_coluna_12}; {item_coluna_14_formatado}; {coluna_csv_10}; {coluna_csv_5}; {coluna_csv_16}; ESCRITURADA"
                    if nova_entrada not in itens_presentes:
                       #print("Número da nota incluída:" + item_coluna_14_formatado)
                        itens_presentes.add(nova_entrada)  # Adiciona ao conjunto
                else:
                    nova_entrada1 = f"{item_coluna_12}; {item_coluna_14_formatado}; {coluna_csv_10}; {coluna_csv_5}; {coluna_csv_16}; NOTA COM PENDENCIA"
                    if nova_entrada1 not in itens_presentes:
                        itens_presentes.add(nova_entrada1)

        itens_presentes = list(itens_presentes)
        return itens_presentes
                                                                          


    def confronto_NDD_principal2_desatualizado(self, arquivoNDD, ArquivoEntrada):
          
        arquivo_txt = ArquivoEntrada #f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA{nro_loja}\\10-2024\\entradaLoja{nro_loja}.txt"                                                                                     
        arquivo_csv = arquivoNDD #f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA{nro_loja}\\10-2024\\arquivoNDD_Loja{nro_loja}.csv"

        # Ler o arquivo CSV
        df_csv = pd.read_csv(arquivo_csv, sep=';', encoding='ISO-8859-1')

        # Converter as colunas 12 e 14 para string
        coluna_csv_12 = df_csv.iloc[:, 12].astype(str).tolist()  # Razão Social
        try:
            coluna_csv_14 = df_csv.iloc[:, 14].astype(str).tolist()  # Número da Nota / df_csv.iloc[:, 14]
        except:
            print("NRO_NOTA_VAZIO")
            coluna_csv_14 = []  # Evita erros se a coluna não existir

        # Extrair as colunas situação (índice 4) e descrição (índice 9)
        coluna_situacao = df_csv.iloc[:, 4].astype(str).tolist()
        coluna_descricao = df_csv.iloc[:, 7].astype(str).tolist() #df_csv.iloc[:, 9]

        lote_coluna_txt = []

        # Ler o arquivo TXT e extrair a coluna 4
        with open(arquivo_txt, 'r', encoding='ISO-8859-1') as file:
            # Pulando a primeira linha para livrar o cabeçalho
            next(file)
            linhas_txt = [linha.strip().split(';') for linha in file]

        # Dividir linhas_txt em lotes
        tamanho_lote = 500  # Ajuste o tamanho conforme necessário
        itens_presentes = set()
        nova_entrada = ""

        for i in range(0, len(linhas_txt), tamanho_lote):
            lote_linhas_txt = linhas_txt[i:i + tamanho_lote]
            # Extrair coluna 4 para o lote atual
            lote_coluna_txt = [linha[4] for linha in lote_linhas_txt if len(linha) > 4]

            # Processar os itens apenas no lote atual
            for situacao, descricao, item_coluna_12, item_coluna_14 in zip(coluna_situacao, coluna_descricao, coluna_csv_12, coluna_csv_14):
                
               # print(f"Situacao : {situacao} -- Descricao : {descricao}")
                if situacao =='Não Informado' and  descricao =='Autorizado': #Nï¿½o Informado / Autorizado
                    item_coluna_14_formatado = item_coluna_14.replace('.0', '')
                    
                    # Verificar se o item da coluna 14 está presente no lote da coluna TXT
                    if item_coluna_14_formatado.lstrip() in lote_coluna_txt:
                        nova_entrada = f"{item_coluna_12} ; {item_coluna_14_formatado} ;NOTA OK"

                        if nova_entrada not in itens_presentes:
                            print("Número da nota incluída:" + item_coluna_14_formatado)
                            itens_presentes.add(nova_entrada)  # Adiciona ao conjunto                                                  
                            continue
                    else:
                        nova_entrada1 = f"{item_coluna_12} ; {item_coluna_14_formatado} ;NOTA COM PENDENCIA"
                        if nova_entrada not in itens_presentes:
                           itens_presentes.add(nova_entrada1)
                                                        

        itens_presentes = list(itens_presentes)
        return itens_presentes      
                                          

    def converte_txt_csv(self, arquivo, nro_loja):
        # Caminho do arquivo .txt de entrada
        arquivo_txt = arquivo

        # Caminho do arquivo .csv de saída
        arquivo_csv = f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA1\\09-2024\\entradaLoja{nro_loja}.csv"

        # Defina o separador do arquivo .txt (por exemplo, ',' ou '\t' para tabulação)
        separador = ';'

        # Lê o arquivo .txt e grava no formato .csv
        with open(arquivo_txt, 'r', encoding='utf-8') as txt_file:
            with open(arquivo_csv, 'w', newline='', encoding='utf-8') as csv_file:
                leitor = csv.reader(txt_file, delimiter=separador)
                escritor = csv.writer(csv_file)
                for linha in leitor:
                    escritor.writerow(linha)

        print(f"Arquivo convertido com sucesso! Salvo em: {arquivo_csv}")

       
    def esperar_fechamento_janela(self, janela_alvo):
            
        # Loop para verificar continuamente se a janela alvo foi aberta
        while True:
            # Obtém todas as janelas ativas
            janelas = gw.getAllWindows()

            # Verifica se a janela alvo está entre as janelas ativas
            if any(janela_alvo in janela.title for janela in janelas):
                print("Janela alvo foi aberta!")
            else:
                print("Janela alvo não existe mais. Encerrando o script.")
                break  # Sai do loop quando a janela alvo não existe mais

            # Pausa por um curto período de tempo antes de verificar novamente
            time.sleep(1)  # Importe time se você ainda não o fez
 

    def SeJanelaExiste_porImagem(self, imagem_janela):
        time.sleep(3)
        try:
                posicao = pyautogui.locateCenterOnScreen(imagem_janela, confidence=0.2)
                if posicao is not None:
                   return True
          
        except Exception as e:
                return False

    

    def converter_para_decimal(self, valor_com_cifrao):
        valor_sem_cifrao=None
        #Removendo o ponto da casa de milhar se houver
        valor_com_cifrao = valor_com_cifrao.replace('.','').lstrip()

        # Verificar se o valor possui cifrão
        if 'R$' in valor_com_cifrao:
            # Remover o cifrão
            valor_sem_cifrao = valor_com_cifrao.replace('R$','').lstrip()
                           
            try:
                #Removendo a virgula da casa decimal e inserindo o ponto para que possa fazer a conversão
                valor_sem_cifrao = valor_sem_cifrao.replace(',','.')

                # Converter para decimal
                valor_decimal = Decimal(valor_sem_cifrao)

                return valor_decimal
            except ValueError:
                print("Erro: Valor não pôde ser convertido para decimal.")
                return None
        else:
            try: 
                valor_sem_cifrao = valor_com_cifrao
                #Removendo a virgula da casa decimal e inserindo o ponto para que possa fazer a conversão
                valor_sem_cifrao = valor_sem_cifrao.replace(',','.')                             
                #fazendo a conversao                
                valor_decimal = Decimal(valor_sem_cifrao)
                
                return valor_decimal
            
            except ValueError:
                print(f"Erro: Valor não pôde ser convertido para decimal AGORAs.",ValueError)
                return None

        
    def SelecionaConteudoCampo(self):
        pyautogui.keyDown('ctrl')
        pyautogui.keyDown('a')
        pyautogui.keyUp('a')
        pyautogui.keyUp('ctrl')


    def copiarCampo(self):
        pyautogui.keyDown('ctrl')
        pyautogui.keyDown('c')
        pyautogui.keyUp('c')
        pyautogui.keyUp('ctrl')
    
        
    def subtracao(self, vlr1, vlr2):
        if vlr1 =="":
            vlr1==0
        if vlr2 =="":
            vlr2==0
                    
        return vlr1 - vlr2


    def xlsx_to_csv(self, nro_lj):
                     
        xlsx_file= f'C:/Projetos_Python/TESOURARIA/arquivos/download/movimentolj{nro_lj}.xlsx'
        csv_file = f'C:/Projetos_Python/TESOURARIA/arquivos/digitacao/movimentolj{nro_lj}.csv'

        
        workbook = load_workbook(filename=xlsx_file, data_only=True)
        # Obter o nome da planilha ativa
        active_sheet_name = workbook.active.title

        # Ler o arquivo .xlsx
        df = pd.read_excel(xlsx_file, engine='openpyxl', sheet_name= active_sheet_name, header=0)

            
        # Iterar sobre todas as colunas para garantir que os números e datas sejam tratados corretamente
        for i, col in enumerate(df.columns):
            
            if i == 0 or i == 8 or i == 9 or i == 10:
                #Verificar se o valor pode ser convertido em float, ignorando strings inválidas
                df[col] = df[col].apply(
                lambda x: str(int(float(x)))  # Converte para int e depois para string, removendo as casas decimais
                if pd.notna(x) and isinstance(x, (int, float, str)) and str(x).replace('.', '', 1).isdigit()
                else str(x)  # Retorna o valor original como string se não for numérico
                )

            if i == 2:
                #Verificar se o valor pode ser convertido em float, ignorando strings inválidas
                df[col] = df[col].apply(lambda x:f'{int(float(x)):,}'.replace(',', 'X').replace('.', ',').replace('X', '.')  # Converte float para int e formata
                if pd.notna(x) and isinstance(x, (int, float, str)) and str(x).replace('.', '', 1).isdigit()  # Verifica se o valor é numérico
                else x  # Retorna o valor original se não puder ser convertido
                )
               
 
            elif i == 3:
               # print("")
                df[col] = df[col].apply(lambda x: 
                pd.to_datetime(x).strftime('%d/%m/%y')  # Converte a data para o formato DD-MM-YY
                if pd.notna(x) and pd.to_datetime(x, errors='coerce') is not pd.NaT  # Verifica se é uma data válida
                else str(x)  # Se não for uma data, converte para string
                )

            
            elif i == 4:
                df[col] = df[col].apply(
                lambda x: f'{float(x):,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
                if pd.notna(x) and isinstance(x, (int, float, str)) and str(x).replace('.', '', 1).isdigit()
                else x  # Retorna o valor original se não for numérico
                )
                  
        df.to_csv(csv_file, sep=';', index=False, encoding='ISO-8859-1', decimal=',', na_rep='', header=True)    


    def verifica_e_apaga_arquivo(self, caminho_arquivo):
        # Verifica se o arquivo existe
        if os.path.exists(caminho_arquivo):
            try:
                # Apaga o arquivo
                os.remove(caminho_arquivo)
                print(f"Arquivo '{caminho_arquivo}' apagado com sucesso.")
            except Exception as e:
                print(f"Erro ao tentar apagar o arquivo: {e}")
        else:
            print(f"O arquivo '{caminho_arquivo}' não existe.")

    
    def GeraSeqTurnoCSV(self, nroLoja, dtaMovimento):
        # Carregando o CSV e garantindo que todas as colunas sejam tratadas como strings
        df_csv = pd.read_csv(
            f'C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj{nroLoja}.csv',  
            sep=';', 
            encoding='ISO-8859-1', #ISO-8859-1
            dtype=str  # Força todas as colunas a serem lidas como strings
        )

        varExecute = RepositorioDAO.DAO()
        con = varExecute.getConection()

        print(df_csv.columns)  
        if "SEQTURNO" not in df_csv.columns:
            sql_query = f""" 
                SELECT h.seqturno, TO_CHAR(h.dtamovimento, 'DD/MM/YYYY') as Data, h.coo, 
                    h.nroempresa || h.nrocheckout || h.coo as KEY
                FROM consincomonitor.tb_docto h
                WHERE h.nroempresa = '{nroLoja}' 
                AND h.dtamovimento = TO_DATE('{dtaMovimento}', 'DD/MM/YYYY')
                """  
            
            df_sql = pd.read_sql_query(sql_query, con.conectar())

            #Garantindo que a coluna 'KEY' seja tratada como string em ambos os DataFrames
            df_csv['KEY'] = df_csv['KEY'].astype(str)
            df_sql['KEY'] = df_sql['KEY'].astype(str)
                        
            
            # Realizando o merge
            df_merged = pd.merge(df_csv, df_sql, on='KEY', how='left')
            print(df_merged.columns)

            #Convertendo a coluna 'seqturno' para string, se existir no DataFrame
            df_merged = df_merged.convert_dtypes(str)
            
            # Salvando o DataFrame resultante no CSV, com todos os dados como string
            df_merged.to_csv(
                f'C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj{nroLoja}.csv', 
                index=False, 
                sep=";", 
                encoding='ISO-8859-1', #ISO-8859-1
                na_rep='',  # Para evitar 'NaN' nos campos vazios
                
            )




           

   
    
            
                           


            
                
                 