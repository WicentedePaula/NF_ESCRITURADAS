from subprocess import Popen
from pywinauto import Application
import pyautogui
import Digitacao
import time
import FuncoesAuxiliares
import pyperclip
import os
from decimal import Decimal
import locale
import RepositorioDAO
import Calculos
import pandas as pd
import csv


if __name__  == '__main__':
    print('EXECUTANDO TEXTE DE MODULOS')
    caminho_csv =f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA3\\10-2024\\arquivoNDD_Loja3.csv"
    caminho_txt =f"\\\\10.11.10.3\\arcomixfs$\\Dados_Contabilidade\\FISCAL\\CONFRONTO_SPED\\LOJA3\\10-2024\\entradaLoja3.txt"
    
    varFunca = FuncoesAuxiliares.Funcao_Apoio()
    varFunca.GeraLogsInfo("Numero 1")
    varFunca.GeraLogsInfo("Numero 2")


    pyautogui.sleep(100000)
    funcao = FuncoesAuxiliares.Funcao_Apoio()
    funcao.xlsx_to_csv("14")
    pyautogui.sleep(3)
    funcao.GeraSeqTurnoCSV("14","09/09/2024")

    """ 
    guptamdiframeAcer_Operador = Application().connect(title_re=".*Tesouraria.*")
    guptamdiframeAcer_Operador = guptamdiframeAcer_Operador.window(class_name='Gupta:MDIFrame') 
    guptamdiframeAcer_Operador.Wait('ready')
    guptamdiframeAcer_Operador[u'&Edit15'].click_input()

    
    varExecute = RepositorioDAO.DAO()
    con = varExecute.getConection()
    varFuncao = FuncoesAuxiliares.Funcao_Apoio()

    dtaMovimento ="12/08/2024"
    ljnro ="27"

    varFuncao.GeraSeqTurnoCSV(ljnro,dtaMovimento)
    

     
   # df_csv = pd.read_csv('C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj3.csv', sep=';')
   # dtaMovimento ="20/07/2024"
   # ljnro = "3"
   

  
    sql_query = 
        SELECT h.seqturno, TO_CHAR(h.dtamovimento, 'DD/MM/YYYY') as Data, h.coo, h.nroempresa || h.nrocheckout || h.coo as KEY
        FROM consincomonitor.tb_docto h
        WHERE h.nroempresa = 3 
        AND h.dtamovimento = TO_DATE(SYSDATE - 23, 'DD/MM/YY')
        
        """    
    """
    sql_query = f 
        SELECT h.seqturno, TO_CHAR(h.dtamovimento, 'DD/MM/YYYY') as Data, h.coo, h.nroempresa || h.nrocheckout || h.coo as KEY
        FROM consincomonitor.tb_docto h
        WHERE h.nroempresa = '{3}' 
        AND h.dtamovimento = TO_DATE('{dtaMovimento}', 'DD/MM/YYYY')"""  
    
    ##df_csv['KEY'] = df_csv['KEY'].astype(str)
    #df_sql['KEY'] = df_sql['KEY'].astype(str)

    #df_csv['KEY'] = df_csv['KEY'].astype(str)
    #df_sql['KEY'] = df_sql['KEY'].astype(str)


   # df_merged = pd.merge(df_csv, df_sql, on=['KEY','KEY'], how='inner')

  #  f_merged = df_merged.convert_dtypes(str)
    
    #print(f_merged)
   # f_merged.to_csv(f'C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj{ljnro}.csv', index=False, sep=";")
   
    
   



    """
    movBase = Application().connect(title_re=".*Movimento Detalhado.*")
    mov = movBase.window(class_name='Gupta:Dialog') 
    mov.Wait('ready')
    mov.close()


    pyautogui.sleep(10000000)

    
    guptamdiframeAcer_Detalhamento = Application().connect(title_re=".*Movimento Detalhado.*")
    guptadialog = guptamdiframeAcer_Detalhamento[u'Gupta:Dialog']
    guptadialog.Wait('ready')
    guptachildtable = guptadialog[u'Gupta:ChildTable']
    guptachildtable.click_input() # Clica na tabela   window.send_close()

    varcalculos = Calculos.Operacoes()
    valor = varcalculos.removeCifraoRetornaString("R$ 9,99")

    pyautogui.sleep(0.35)
    pyautogui.press('ctrl')
    pyautogui.sleep(0.35)
    pyautogui.press('insert') # Abre a linha para digitação 
    pyautogui.sleep(0.35)
    pyautogui.press('insert') # Inseri a linha 
    pyautogui.press('insert') # ABRE linha 
    pyautogui.sleep(1.35)
    #pyautogui.write("787 - DEVOLUCAO")
    pyautogui.write("787 - ")
    pyautogui.press('down')
    pyautogui.sleep(0.35)
    pyautogui.press('up')
    pyautogui.sleep(0.35)
    pyautogui.press('tab')
    pyautogui.sleep(0.35)
    pyautogui.write(valor) # Informando o  valor 
   
    guptadialog[u'Button3'].click_input() #Click na no  botão conciliar
    pyautogui.sleep(2)

    guptadialog[u'Button6'].click_input() # Confirma e fecha a janela
    
    jnErro = varFuncao.check_window_exists("Movimento Detalhado")
    if jnErro == True:
         appDlgConf_Exclusao = Application().connect(title_re=".*Movimento Detalhado.*")
         window = appDlgConf_Exclusao.Dialog
         window.Wait('ready')
         button = window[u'Button6']
         button.click_input()


    pyautogui.sleep(100000)


    varArquivosLojas =f"C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj28.csv"
    if os.path.exists(varArquivosLojas): #Testa para ver se o arquivo existe.
   
      with open(varArquivosLojas, 'r') as c:
        
         lines = c.readlines()
      
      indices_para_remover = set()

      #LOOP DA PLANILHA
      for index, line in enumerate(lines[1:], start=1):
         line_data = line.strip().split(";")
      
        
         codigo=line_data[0]
         usuario=line_data[1]
         pdv=line_data[2]
         data=line_data[3]
         dinheiro=line_data[4]
         devolucao=line_data[5]
         sobra=line_data[6]
         quebra=line_data[7]
         loja=line_data[8]
         coo=line_data[9]
         key=line_data[10] 
         turnoCsv=str(line_data[11])

         if index == len(lines) -1:
             print(usuario)
         else:
             continue

        

        

         
         
            
      

    pyautogui.sleep(100000)
    # 1. Carregar o arquivo CSV
    df_csv = pd.read_csv("C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj28.csv")
        
    query = SELECT h.seqturno, TO_CHAR(h.dtamovimento, 'DD/MM/YYYY') as Data, h.coo, h.nroempresa || h.nrocheckout || h.coo as key
      FROM consincomonitor.tb_docto h
                     WHERE h.nroempresa = 27 
                     AND h.dtamovimento = TO_DATE(SYSDATE - 1, 'DD/MM/YY')
    
    

    df_sql = pd.read_sql_query(query, con.conectar())     
    
    con.desconectar()

    df_merged = pd.merge(df_csv, df_sql, on='Key', how='left')  # 'how' pode ser 'inner', 'outer', 'left', 'right'

    # 4. Salvar o resultado em um arquivo CSV
    df_merged.to_csv('C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\arquivo_resultante.csv', index=False)

    
  
    varArquivosLojas =f"C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj27.csv"
    if os.path.exists(varArquivosLojas): #Testa para ver se o arquivo existe.
   
      with open(varArquivosLojas, 'r') as c:
        
         lines = c.readlines()

      
      indices_para_remover = set()

      #LOOP DA PLANILHA
      for index, line in enumerate(lines[1:], start=1):
         line_data = line.strip().split(";")
                  
         codigo=line_data[0]
         usuario=line_data[1]
         pdv=line_data[2]
         data=line_data[3]
         dinheiro=line_data[4]
         devolucao=line_data[5]
         sobra=line_data[6]
         quebra=line_data[7]
         loja=line_data[8]
         coo=line_data[9]
         key=line_data[10] 
         turnoCsv=str(line_data[11])

         
         if codigo == "11667":
            print(f" PDV ENCONTRATO : {pdv} -- TURNO :{turnoCsv} -- OPERADOR : {usuario}")
            break
         else:
            print(f"LINHA A SER EXCLUÍDA : {pdv} -- TURNO :{turnoCsv} -- OPERADOR : {usuario}")
            indices_para_remover.add(index)


         
         with open(varArquivosLojas, 'w') as c:
            c.write(lines[0])  # Escreve o cabeçalho
            for index, line in enumerate(lines[1:], start=1):
               if index not in indices_para_remover:
                  c.write(line)

             

      
        

    pyautogui.sleep(100000)
    guptamdiframeAcer_dinheiro = Application().connect(title_re=".*Tesouraria.*")
    guptadialog = guptamdiframeAcer_dinheiro[u'Gupta:Dialog']
    guptadialog.Wait('ready')

    
    guptadialog[u'&Dinheiro'].click_input() #  Clicando em no botão dinheiro 
    pyautogui.sleep(0.25)
    guptadialog[u'Edit2'].type_keys("15,98") # Informando o dinheiro
    pyautogui.sleep(0.25)

    

    guptadialog[u'Button1'].click_input() # clicando em ok

    pyautogui.sleep(100000)
    guptamdiframeAcer_Operador = Application().connect(title_re=".*Tesouraria.*")
    guptamdiframeAcer_Operador = guptamdiframeAcer_Operador.window(class_name='Gupta:MDIFrame') 
    guptamdiframeAcer_Operador.Wait('ready')
   
    guptamdiframeAcer_Operador[u'Edit49'].type_keys("^c") # Executa a copia  do controle digitacao  [u'Edit49', u'Diferen\xe7aEdit']
    srtDiferenca = pyperclip.paste() # Insere na variável
    print(srtDiferenca)

    pyautogui.sleep(10000)
    guptamdiframeAcer_Detalhamento = Application().connect(title_re=".*Movimento Detalhado.*")
    guptadialog = guptamdiframeAcer_Detalhamento[u'Gupta:Dialog']
    guptadialog.Wait('ready')
                     
    guptadialog[u'Button3'].click_input() #Click na no  botão conciliar
    pyautogui.sleep(2)
    guptadialog[u'Button6'].click_input() # Confirma e fecha a janela                                                                                                   

    pyautogui.sleep(100000)

    guptamdiframeAcer_Detalhamento = Application().connect(title_re=".*Movimento Detalhado.*")
    guptadialog = guptamdiframeAcer_Detalhamento[u'Gupta:Dialog']
    guptadialog.Wait('ready')
    guptachildtable = guptadialog[u'Gupta:ChildTable']
    guptachildtable.click_input()
   
    pyautogui.sleep(0.35)
    pyautogui.press('ctrl')
    pyautogui.sleep(0.35)
    pyautogui.press('insert')
    pyautogui.sleep(0.35)
    pyautogui.press('insert')
    pyautogui.sleep(0.35)
    pyautogui.press('insert')
    pyautogui.write("787 - DEVOLUCAO")
    pyautogui.sleep(0.35)
    pyautogui.press('tab')
    pyautogui.sleep(0.35)
    pyautogui.write("9,99")

    guptadialog[u'Button3'].click_input() #Click na no campo data  [u'Button3', u'Concilia Todos', u'Concilia TodosButton']
    pyautogui.sleep(0.35)
    guptadialog[u'Button6'].click_input()
    



    pyautogui.sleep(300000)
    varExec = RepositorioDAO.DAO()
    varCalculos = Calculos.Operacoes()
    cal = varCalculos.convertStringEmDecimal("250.439,39")
    print(cal)

    pyautogui.sleep(300000)
    guptamdiframeAcer_Operador = Application().connect(title_re=".*Tesouraria.*")
    guptamdiframeAcer_Operador = guptamdiframeAcer_Operador.window(class_name='Gupta:MDIFrame') 
    guptamdiframeAcer_Operador.Wait('ready')
    Gt_Final_Formulario = guptamdiframeAcer_Operador.Edit49
    Gt_Final_Formulario.SetFocus()
   # Gt_Final_Formulario.type_keys("^a") #Seleciona o texto a ser copiado     //guptamdiframe.Edit49
    Gt_Final_Formulario.type_keys("^c") # Executa a copia

    pyautogui.sleep(1)
    texto_copiado = pyperclip.paste()

    print("Texto copiado:", texto_copiado)
    
    
    pyautogui.sleep(300000)
    # Definir a localidade para pt-BR
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    time.sleep(4)

    varFuncao = FuncoesAuxiliares.Funcao_Apoio() 
    pyautogui.click(1219,771, duration=0.35) # Clica no campo diferença
    varFuncao.SelecionaConteudoCampo()
    varFuncao.copiarCampo()  
    srtDiferenca = pyperclip.paste() #Captura o valor da diferença no inicio da digitação
   
    print(f"Valor do srtDiferenca ln 24 :{srtDiferenca}")
    time.sleep(300000)

    pyautogui.click(852,553, duration=0.25) #Clica na caixa do  dinheiro
    #pyautogui.click(1217,769, duration=0.25) # Clica no campo diferença

    #x, y = pyautogui.position()
    #print(x,y)

    #pyautogui.click(1226,765) # Clica no campo diferença
    varFuncao = FuncoesAuxiliares.Funcao_Apoio()
    # Exemplo de uso:
    #valor_com_cifrao = "R$ 10,15".lstrip()
    #valor_com_cifrao = "10.20"
    #valor_decimal = varFuncao.converter_para_decimal(valor_com_cifrao)
    #print("Valor decimal resul:", valor_decimal)
    
    #valor_decimal = varFuncao.converter_para_decimal(valor_com_cifrao)
    #if valor_decimal is not None:
    #   print("Valor decimal resul:", valor_decimal)



    time.sleep(10000)
    desc_dinheiro_sistema =0
    desc_dinhero_tabela = 0

    varArquivosLojas =f"C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj1.csv"
    if os.path.exists(varArquivosLojas): #Testa para ver se o arquivo existe.
   
         with open(varArquivosLojas) as c:
            next(c)
            #LOOP DA PLANILHA
            for line in c:
               line=line.strip()
               line=line.split(",")

               codigo=line[0]
               usuario=line[1]
               pdv=line[2]
               data=line[3]
               dinheiro=line[4]
               devolucao=line[5]
               sobra=line[6]
               quebra=line[7].rstrip()
               loja=line[8]
               coo=line[9]
               key=line[10]
               
               pyautogui.click(1227,766) # Clica no campo diferença
               varFuncao.SelecionaConteudoCampo()

               time.sleep(1)
               varFuncao.copiarCampo()
              
               time.sleep(1)
               srtDiferenca = pyperclip.paste() #Captura o valor da diferença no inicio da digitação
               
               

               print(f"valor da variavel :{srtDiferenca}")

               if srtDiferenca != '0' and srtDiferenca !='':
                  #Converteu o valor da devolucao(outros) e da diferença em float
                  desc_Diferenca = Decimal(srtDiferenca)
                  desc_Devolucao = Decimal(devolucao)
                     
               if desc_dinheiro_sistema != '0' and srtDiferenca !='':
                  #Converter o valor do dinheiro da tabela e do dinheiro do sistema
                  desc_dinheiro_sistema = Decimal("999")

               if desc_dinhero_tabela != '0' and srtDiferenca !='':
                  desc_dinhero_tabela = Decimal(dinheiro)
               
                                         

               dif_dinheiro_sistema_e_tabela = varFuncao.subtracao(desc_dinheiro_sistema, desc_dinhero_tabela)
               dif_devolucao_diferencao = varFuncao.subtracao(desc_Diferenca, desc_Devolucao)

               if dif_devolucao_diferencao != '0':
                  pyautogui.leftClick(772,310, duration=0.25)#Proximo Pdv
                  continue
                                    
               #######################################################################################
               pyautogui.click(649,547, duration=0.25) #Clica em dinheiro
               pyautogui.typewrite(dinheiro ,interval=0.15) #Informa o dinheiro
               pyautogui.click(983,562, duration=0.25) #Confirma 
               #######################################################################################

               print(f" Dados da Planilha: {codigo}, {usuario}, {pdv}, {data}, {dinheiro}, {devolucao}, {sobra}, {quebra}, {loja}, {coo}, {key}") 


    time.sleep(10000)
    x, y = pyautogui.position()
    print(x,y)

   # with open(f'C:\\Projetos_Python\\TESOURARIA\\arquivos\\digitacao\\movimentolj3.csv') as c:
    #    next(c)

        #LOOP DA PLANILHA
     #   for line in c:
       #     line=line.strip()
       #     line=line.split(",")
      #      print(" Dados da Planilha : ",line)

   # time.sleep(3)
   # pyautogui.click(pyauogui.locateCenterOnScreen("C:\\Projetos_Python\\TESOURARIA\\img\\campos_tesouraria\\tx_data.png", confidence=0.95), duration=0.45)
    #varExectaTeste = Digtitacao.Digitacao()
    #varExectaTeste.acertoOperador("13/03/2024")
   # pyautogui.leftClick(pyautogui.locateCenterOnScreen("C:\\Projetos_Python\\TESOURARIA\\img\\campos_tesouraria\\bt_movimento.png", confidence=0.890, duration=0.25)
            
    time.sleep(10000)
    varFuncao = FuncoesAuxiliares.Funcao_Apoio()
    resul = varFuncao.SeJanelaNaoExiste("C:\\Projetos_Python\\TESOURARIA\\img\\janelas\\jn_acerto_operador.png")
    if resul is False:
      print(" JENELA NÃO ENCONTRATA !!!!!")
    else:
       print("ENCONTRADA")
       

    print('CONCLUINDO')
   """