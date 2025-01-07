import ConectaBD
#from ConnectionBD import Connect_BD

class DAO:

        def __init__(self):
                self.con = ConectaBD.ConexaoOracle

        def getConection(self):
                return  ConectaBD.ConexaoOracle()

        def executaQuery(self, query):
                conexao = self.con.conectar(self)
                resultados = []

                # Verifica se a conexão foi bem-sucedida
                if conexao is not None:
                        try:
                                cursor = conexao.cursor()

                                # Executa a consulta
                                for row in cursor.execute(query):
                                        resultados.append(row)

                                # Fecha o cursor
                                cursor.close()

                        except Exception as e:
                         print(f"Erro ao executar a consulta: {e}")

                        finally:
                        # Sempre desconecta, mesmo em caso de exceção
                         self.con.desconectar(self)
                         

                else:
                        print("Falha ao conectar ao banco de dados. CONEXAO É NULL.")

                return resultados
        

        def RetornaMovimento(self,nro_loja, dta_movimento):
                query = f"""
                        SELECT 
                        FI_TSMOVTOOPERADOR.NROEMPRESA, 
                        TO_CHAR(FI_TSMOVTOOPERADOR.DTAMOVIMENTO, 'dd/mm/yyyy') AS data,
                        FI_TSMOVTOOPERADOR.NROTURNO,
                        FI_TSMOVTOOPERADOR.GTINICIO,
                        FI_TSMOVTOOPERADOR.GTFINAL,
                        FI_TSMOVTOOPERADOR.ENCARGOS,
                        FI_TSMOVTOOPERADOR.VLRTOTALNFENFCESAT,
                        FI_TSMOVTOOPERADOR.TOTAL,
                        FI_TSMOVTOOPERADOR.VLRBANCARIO,
                        FI_TSMOVTOOPERADOR.QTDEDOCBANCARIO,
                        FI_TSMOVTOOPERADOR.NROEMPRESAMAE,
                        FI_TSMOVTOOPERADOR.ACERTADO,
                        FI_TSMOVTOOPERADOR.FECHADO,
                        FI_TSMOVTOOPERADOR.USUFECHOU,
                        FI_TSMOVTOOPERADOR.DTAFECHOU,
                        FI_TSMOVTOOPERADOR.USUALTERACAO,
                        FI_TSMOVTOOPERADOR.VERSAO,
                        FI_TSMOVTOOPERADOR.SEQIDENTIFICA,
                        NVL(FI_TSMOVTOOPERADOR.INDQUEBRAECF, 'N'),
                        FI_TSMOVTOOPERADOR.NROPDV,
                        FI_TSMOVTOOPERADOR.SOFTPDV,
                        FI_TSMOVTOOPERADOR.CODOPERADOR,
                        NVL(FI_TSMOVTOOPERADOR.SOFTPDV, 'DIG.MANUALMENTE'), 
                        FI_TSMOVTOOPERADOR.USUACERTOU,
                        FI_TSMOVTOOPERADOR.DTAHORAACERTOU,
                        usu.NOME
                        FROM FI_TSMOVTOOPERADOR, ge_usuario usu
                        WHERE FI_TSMOVTOOPERADOR.NROEMPRESA = :nro_loja
                        AND NVL(VERSAO, 'A') = 'N'
                        AND FI_TSMOVTOOPERADOR.DTAMOVIMENTO = TO_DATE(:dta_movimento, 'dd/mm/yyyy')
                        AND usu.SEQUSUARIO = FI_TSMOVTOOPERADOR.CODOPERADOR
                        ORDER BY NROPDV, NROTURNO
                """

                conexao = self.con.conectar(self)
                resultados = []

                # Verifica se a conexão foi bem-sucedida
                if conexao is not None:
                     try:
                           cursor = conexao.cursor()

                        # Executa a consulta com parâmetros
                           cursor.execute(query, {'nro_loja': nro_loja, 'dta_movimento': dta_movimento})
                           resultados = cursor.fetchall()

                        # Fecha o cursor 
                           cursor.close()

                     except Exception as e:
                          print(f"Erro ao executar a consulta: {e}")

                     finally:
                        # Sempre desconecta, mesmo em caso de exceção
                        self.con.desconectar(self)

                else:
                        print("Falha ao conectar ao banco de dados. CONEXAO É NULL.")

                return resultados
        
        def NF_pendentes_Ato_da_Entrega(self, nro_loja, dta_inical, dta_final):
                conexao = self.con.conectar(self)
                resultados = []

                query = f"""
                        select MRLV_NFEIMPORTACAO.SEQNOTAFISCAL, COUNT(DISTINCT MRLV_NFEIMPORTACAO.SEQPESSOA), MRLV_NFEIMPORTACAO.NUMERONF, MRLV_NFEIMPORTACAO.NROEMPRESA, MRLV_NFEIMPORTACAO.SERIENF, MAX(MRLV_NFEIMPORTACAO.SEQPESSOA), MAX(MRLV_NFEIMPORTACAO.NOMERAZAO), MRLV_NFEIMPORTACAO.NATOPERACAO, to_date(MRLV_NFEIMPORTACAO.DTAEMISSAO,'DD-MM-YYYY'), MRLV_NFEIMPORTACAO.PEDIDO, MRLV_NFEIMPORTACAO.VLRTOTNF, MRLV_NFEIMPORTACAO.CHAVEACESSO, MRLV_NFEIMPORTACAO.CODOPERMANIFESTDEST, MRLV_NFEIMPORTACAO.STATUSRETMANIFESTDEST, MAX(MRLV_NFEIMPORTACAO.NROCGCCPF), MAX(MRLV_NFEIMPORTACAO.FISICAJURIDICA), NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  
                        from MRLV_NFEIMPORTACAO 

                        where NROEMPRESA ={nro_loja}
                        AND NOT EXISTS(SELECT *
                                        FROM	MRL_NFEIMPORTACAO
                                        WHERE	MRL_NFEIMPORTACAO.SEQNOTAFISCAL = MRLV_NFEIMPORTACAO.SEQNOTAFISCAL
                                        AND 	NVL(MRL_NFEIMPORTACAO.INDNFORCAMENTO,'N') = 'S') 
                        AND NROEMPRESA ={nro_loja}                             
                        AND DTAEMISSAOC5 between TO_date('{dta_inical}', 'dd-mm-yyyy') and TO_date('{dta_final}', 'dd-mm-yyyy') AND NOT EXISTS (SELECT X.SEQNOTAFISCAL
                  					FROM MLF_AUXNOTAFISCAL X
                 					WHERE X.NUMERONF   =  MRLV_NFEIMPORTACAO.NUMERONF
          					   AND X.SEQPESSOA  =  MRLV_NFEIMPORTACAO.SEQPESSOA
           					   AND  LPAD(TRIM(X.SERIENF),3,'0') = LPAD(MRLV_NFEIMPORTACAO.SERIENF,3,'0')
           					   AND X.NROEMPRESA =  MRLV_NFEIMPORTACAO.NROEMPRESA
					   AND X.TIPNOTAFISCAL = 'E'
					   AND nvl(to_char(X.DTAEMISSAO, 'dd/mm/yyyy'),MRLV_NFEIMPORTACAO.DTAEMISSAO) = MRLV_NFEIMPORTACAO.DTAEMISSAO
					   AND NVL(X.NFECHAVEACESSO, NVL(X.NFECHAVEACESSOCOPIA, MRLV_NFEIMPORTACAO.CHAVEACESSO)) = MRLV_NFEIMPORTACAO.CHAVEACESSO
					
					  UNION
					 
					  SELECT XX.SEQNOTAFISCAL
                  				   FROM MLF_AUXNOTAFISCAL XX, MAX_CODGERALOPER PP
                 				   WHERE XX.NFREFERENCIANRO =    MRLV_NFEIMPORTACAO.NUMERONF
           					 AND    XX.NFREFERENCIASERIE =  MRLV_NFEIMPORTACAO.SERIENF
           					 AND    XX.SEQPESSOA       =    MRLV_NFEIMPORTACAO.SEQPESSOA            
           				 	 AND    XX.NROEMPRESA      =    MRLV_NFEIMPORTACAO.NROEMPRESA
           				 	 AND    XX.TIPNOTAFISCAL   =    'E'
           					 AND    PP.CODGERALOPER                    =       XX.CODGERALOPER
           					 AND    PP.TIPUSO                          =       'E'
           					 AND    PP.INDNFREFPRODRURAL               =       'S'
           					 AND    XX.STATUSNF               !=       'C'
					  					
					UNION
					
					SELECT	ANF.SEQNOTAFISCAL
					FROM	MLF_AUXNOTAFISCAL ANF, MAX_CODGERALOPER CGO
					WHERE	ANF.CODGERALOPER		= CGO.CODGERALOPER
					AND	ANF.NFREFERENCIANRO 		= MRLV_NFEIMPORTACAO.NUMERONF
					AND	ANF.NFREFERENCIASERIE		= MRLV_NFEIMPORTACAO.SERIENF
					AND	ANF.NROEMPRESA		= MRLV_NFEIMPORTACAO.NROEMPRESA
					AND	ANF.SEQPESSOA			= MRLV_NFEIMPORTACAO.SEQPESSOA
					AND	ANF.TIPNOTAFISCAL		= 'E'
					AND	CGO.TIPCGO			= 'E'
					AND	CGO.TIPUSO			= 'E'
					)
     			 AND NOT EXISTS( 	SELECT Y.SEQNOTAFISCAL
                  			   		FROM MLF_NOTAFISCAL Y
                  			  		WHERE Y.NUMERONF   =  MRLV_NFEIMPORTACAO.NUMERONF
           				    	   AND Y.SEQPESSOA  =  MRLV_NFEIMPORTACAO.SEQPESSOA
           				    	   AND LPAD(TRIM(Y.SERIENF),3,'0') = LPAD(MRLV_NFEIMPORTACAO.SERIENF,3,'0')
           				    	   AND Y.NROEMPRESA =  MRLV_NFEIMPORTACAO.NROEMPRESA
   					   AND Y.TIPNOTAFISCAL = 'E'
					   AND to_char(Y.DTAEMISSAO, 'dd/mm/yyyy') = MRLV_NFEIMPORTACAO.DTAEMISSAO
					   AND NVL(Y.NFECHAVEACESSO, 0) = MRLV_NFEIMPORTACAO.CHAVEACESSO
					
					UNION
           
           					SELECT YY.SEQNOTAFISCAL
           					FROM   MLF_NOTAFISCAL YY, MAX_CODGERALOPER PP
           					WHERE  YY.NFREFERENCIANRO =    MRLV_NFEIMPORTACAO.NUMERONF
           					AND    YY.NFREFERENCIASERIE =  MRLV_NFEIMPORTACAO.SERIENF
           					AND    YY.SEQPESSOA       =    MRLV_NFEIMPORTACAO.SEQPESSOA            
           					AND    YY.NROEMPRESA      =    MRLV_NFEIMPORTACAO.NROEMPRESA
           					AND    YY.TIPNOTAFISCAL   =    'E'
           					AND    to_char(YY.NFREFERENCIADTAEMISSAO, 'dd/mm/yyyy') = MRLV_NFEIMPORTACAO.DTAEMISSAO
           					AND    NVL(YY.NFEREFERENCIACHAVE,0)       =       MRLV_NFEIMPORTACAO.CHAVEACESSO      
           					AND    PP.CODGERALOPER                    =       YY.CODGERALOPER
           					AND    PP.TIPUSO                          =       'E'
           					AND    PP.INDNFREFPRODRURAL               =       'S'
					
					UNION
					
					SELECT	NF.SEQNOTAFISCAL
					FROM	MLF_NOTAFISCAL NF, MAX_CODGERALOPER CGO
					WHERE	NF.CODGERALOPER	= CGO.CODGERALOPER
					AND	NF.NFREFERENCIANRO 	= MRLV_NFEIMPORTACAO.NUMERONF
					AND	NF.NFREFERENCIASERIE	= MRLV_NFEIMPORTACAO.SERIENF
					AND	NF.NROEMPRESA		= MRLV_NFEIMPORTACAO.NROEMPRESA
					AND	NF.SEQPESSOA		= MRLV_NFEIMPORTACAO.SEQPESSOA
					AND	NF.TIPNOTAFISCAL	= 'E'
					AND	CGO.TIPCGO		= 'E'
					AND	CGO.TIPUSO		= 'E'
					) 

                        group by SEQNOTAFISCAL,NUMERONF,SERIENF,NATOPERACAO, DTAEMISSAO,PEDIDO, VLRTOTNF,CHAVEACESSO,NROEMPRESA, CODOPERMANIFESTDEST, DTAEMISSAOC5, 
                        STATUSRETMANIFESTDEST

                        order by DTAEMISSAOC5, NUMERONF 

                                """

                # Verifica se a conexão foi bem-sucedida
                if conexao is not None:
                        try:
                                cursor = conexao.cursor()

                                # Executa a consulta
                                for row in cursor.execute(query):
                                        resultados.append(row)

                                # Fecha o cursor
                                cursor.close()

                        except Exception as e:
                         print(f"Erro ao executar a consulta: {e}")

                        finally:
                        # Sempre desconecta, mesmo em caso de exceção
                         self.con.desconectar(self)
                         

                else:
                        print("Falha ao conectar ao banco de dados. CONEXAO É NULL.")
                
                resultados = list(resultados)
                return resultados
             
             
             
                
                
                    

        
                     