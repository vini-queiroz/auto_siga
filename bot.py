from botcity.core import DesktopBot
from tkinter import messagebox
import pyodbc

class Bot(DesktopBot):
    def action(self, execution=None):
        def abrir_produção():
                # Abre o ambiente de apontamento de produção:
            if not self.find( "producao", matching=0.97, waiting_time=1500):  # Clicar em Produção
                if self.find( "TOTVS", matching=0.97, waiting_time=2000):
                    self.click_relative(16, 28)
                else:
                    self.not_found("producao")
            self.find( "producao", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "incluir", matching=0.97, waiting_time=10000)
            self.click()
            self.wait(200)

            # Após a rotina finalizada o ambiente estará pronta para realizar o apontamento das OP
        def apontar_op(OP):

            #Conexa com banco de dados
            server = 'AVANCODB'
            database = 'p12'
            username = 'ssql'
            password = 'ssql987'

            connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

            conn = pyodbc.connect(connection_string)

            cursor = conn.cursor()

            quary = '''
            WITH APONTAMENTO AS (

            SELECT
                
                D4_COD
                , D4_LOCAL
                , D4_OP
                ,	D4_QUANT
                --, B2_QATU
                , CASE 
                    WHEN D4_LOCAL = '01' THEN B2_QATU  --AND B2_LOCAL = '01'
                    WHEN D4_LOCAL = '77' THEN B2_QATU  --AND B2_LOCAL = '77'
                    END AS SALDOS
                --, CASE 
                    --WHEN D4_QUANT <= B2_QATU THEN 'NAO'
                    --ELSE 'SIM' END AS FALTA_SALDO
        
                FROM SD4010 
                INNER JOIN SB2010 
                    ON D4_COD = B2_COD 
                    AND D4_LOCAL = B2_LOCAL
                    AND  SB2010.D_E_L_E_T_<>'*'
                    AND B2_FILIAL = '01'
                    AND (B2_LOCAL = '01' OR B2_LOCAL= '77')
        
                WHERE
                
                    SD4010.D_E_L_E_T_<>'*' 
                    AND (D4_OP LIKE 'A01175%' 
				        OR D4_OP LIKE 'A01174%')
                    AND D4_QTDEORI = D4_QUANT 
                    AND D4_QUANT != 0 
                    AND D4_FILIAL = '01'
                    AND D4_COD NOT LIKE 'MOD%'
                    )	
            SELECT DISTINCT D4_OP
                FROM APONTAMENTO
                WHERE D4_OP NOT IN (
                    SELECT D4_OP
                    FROM APONTAMENTO
                    WHERE D4_QUANT > SALDOS
                );'''
            cursor.execute(quary)
               #"SELECT DISTINCT(D4_OP) FROM SD4010 WHERE SD4010.D_E_L_E_T_<>'*' AND D4_OP LIKE ? AND D4_QTDEORI = D4_QUANT AND D4_QUANT != 0 AND D4_FILIAL = '01'  ORDER BY D4_OP DESC", OP


            resultados = [row.D4_OP for row in cursor.fetchall()]

            if resultados != '':
                print('Ordens de Produção obtidas!')
            else:
                print('Erro ao obter ops!')
                return 0

            conn.close()

            # Rotina para apontar as ordens de produção
            self.type_keys(['ctrl', 'alt', 'tab'])
            if not self.find( "TOTVS_encontrar", matching=0.97, waiting_time=1500):
                self.not_found("TOTVS_encontrar")          
            self.click_relative(114, 61)
            self.wait(300)

            for resultado in resultados:

                self.find_until( "Ord_producao", matching=0.97, waiting_time=1000)
                    #self.not_found("Ord_producao")
                self.double_click_relative(7, 32)
                self.paste(resultado)

                # Condicionais para o apontamento...
                if self.find( "encerrada", matching=0.97, waiting_time=150):                                                                                        # Se a OP estiver encerrada ou não existir o programa passa para a proxima OP
                    self.enter()
                elif self.find( "n_existe", matching=0.97, waiting_time=150):
                    self.enter()

                    # Caso o contrário, buscar pelo campo Tipo de Movimentação e digitar '010'
                else:
                    self.find( "tipo_mov", matching=0.97, waiting_time=300)
                    self.click_relative(6, 34)
                    self.kb_type('010')
                    self.find( "Salvar", matching=0.97, waiting_time=300)
                    self.click()

                    if self.find( "No_saldo", matching=0.97, waiting_time=3000):
                        self.enter()
                        if not self.find( "fechar_saldo", matching=0.97, waiting_time=300):                                                                         # Se faltar saldo, fechar o campo da indicação de saldo e passar para a próxima OP a ser apontada.
                            self.not_found("fechar_saldo")
                        self.click()

                    # Se não, o Apontamento estará ok.
                    elif self.find_until("documento_vazio", matching=0.97, waiting_time=14000):                                                                  #(TIME) PARAMETRO A SER MUDADO CONFORME A MAIOR ORDEM DE PRODUÇÃO A SER APONTADA.
                        b=0

            messagebox.showinfo('Apontamentos', 'Apontamentos Finalizados!')
            return
        def firmar_ops(op_inicial, op_final):
            #   Rotina para firmar ops previstas.
                #  Abrir ops previstas

            # while self.find("marcar_111105", matching=0.97, waiting_time=800):  # 111105 - Puxadores.
            #     self.enter()
            #     self.type_down()
            # while self.find("click_111106", matching=0.97, waiting_time=800):  # 111106 - Montagem.
            #     self.enter()
            #     self.type_down()
            # while self.find("111108", matching=0.97, waiting_time=10000):  # 111108 - Expedição.
            #     self.enter()
            #     self.type_down()
            #     #
            #     if not self.find( "cnt_custo", matching=0.97, waiting_time=10000):
            #         self.not_found("cnt_custo")
            #     self.click_relative(60, 14)
            #     self.wait(3000)                                                             # Procurando pelo C.Custo da Elétrica
            #     self.kb_type('111112')
            #     self.wait(2000)
            #     self.enter()
            #     self.enter()
            #
            # if not self.find( "clik_1111112", matching=0.97, waiting_time=10000):
            #     self.not_found("clik_1111112")
            # self.click()
            #
            #
            # while self.find( "marcar_111112", matching=0.97, waiting_time=1000):      # 111112 - Elétrica.
            #     self.enter()
            #     self.type_down()
            #
            # while self.find( "121104", matching=0.97, waiting_time=1000):
            #     self.type_down()

            while self.find( "121108", matching=0.97, waiting_time=1000):             # 121108 - Almoxarifado.
                self.enter()
                self.type_down()

            # Definir momento de parada...(Melhoria).

            #  Excluindo Ops que não utilizaremos...

            if not self.find( "Numero_da_op", matching=0.97, waiting_time=10000):
                self.not_found("Numero_da_op")
            self.double_click()
            self.wait(2500)
            self.find( "Outras_acoes", matching=0.97, waiting_time=10000)
            self.click()
        def baixa_doc():
            a = 0
            if not self.find( "pendencias", matching=0.97, waiting_time=10000):
                self.not_found("pendencias")
            self.click()
            
            while a < 1:
                if not self.find( "baixar", matching=0.97, waiting_time=10000):
                    self.not_found("baixar")
                self.click()
               
                if not self.find( "cofima", matching=0.97, waiting_time=10000):
                    self.not_found("cofima")
                self.click()
                if not self.find( "yes", matching=0.97, waiting_time=10000):
                    self.not_found("yes")
                self.click()

        # Utilização das Rotinas.

        #firmar_ops('A0116601001', 'A0117701zzz')
        #abrir_produção()
        apontar_op('A00ZEI%')
        #baixa_doc()

    def not_found(self, label):
        print(f"Element not found: {label}")
if __name__ == '__main__':
    Bot.main()







