import pyautogui as py
from botcity.core import DesktopBot
from tkinter import messagebox

class Bot(DesktopBot):
    def action(self, execution=None):

        def iniciar_siga(meu_usuario, senha, amb):
                # Rotina para iniciar o micro siga
            self.execute(r'C:\Users\Antonio\Desktop\SmartClient R33 Prod.lnk')    # Abrir o atalho para o siga
            if not self.find( "btn_ok", matching=0.97, waiting_time=10000):
                self.not_found("btn_ok")
            self.click()
            if not self.find( "usuario", matching=0.97, waiting_time=100000): # Encontra o campo usuário
                self.not_found("usuario")
            self.kb_type(meu_usuario)       # Digitar usuário
            self.tab()
            self.kb_type(senha)       # Digitar senha
            self.enter()
            if not self.find( "ambiente", matching=0.97, waiting_time=100000):
                self.not_found("ambiente")
            for i in range(5):  # Tecla tab 2 vezes, depois o enter.
                self.tab()
            self.kb_type(amb)
            for i in range(2):      # Tecla tab 2 vezes, depois o enter.
                self.tab()
            self.enter()

        def abrir_produção():

                # Abre o ambiente de apontamento de produção:
            if not self.find( "atualizacao", matching=0.97, waiting_time=10000):       # Clicar em Atualizações
                self.not_found("atualizacao")
            self.click()
            if not self.find( "producao", matching=0.97, waiting_time=10000):      # Clicar em Produção
                self.not_found("producao")
            self.click()
            if not self.find( "incluir", matching=0.97, waiting_time=10000):
                self.not_found("incluir")
            self.click()
            self.wait(200)

            # Após a rotina finalizada o ambiente estará pronta para realizar o apontamento das OPs

        def relatorio_estoque(armazem):
            self.find( "relatorios", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "saldos_estoques", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "planilha", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "tipo_planilha", matching=0.97, waiting_time=10000)
            self.click_relative(183, 24)
            self.find( "linhas_brancas", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "outras_acoes", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "parametros", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "do_armazem", matching=0.97, waiting_time=10000)
            self.click_relative(210, 11)
            self.kb_type(armazem)
            self.kb_type(armazem)
            self.find( "ok", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "imprimir", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "abrir", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "sim", matching=0.97, waiting_time=10000)
            self.click()

        def apontar_op():
                # Rotina para apontar as ordens de produção
                  
            self.type_keys(['ctrl', 'alt', 'tab'])
            if not self.find( "Apontamentos", matching=0.97, waiting_time=10000):
                self.not_found("Apontamentos")
            self.click_relative(121, 100)
            if not self.find( "Prod_excel", matching=0.97, waiting_time=10000):
                self.not_found("Prod_excel")
            self.click()
            self.type_keys(['ctrl', 'down'])
            self.type_keys(['ctrl','left'])

                #Looping para apontamento das OPs.
            a=0
            while a < 1: #not self.find( "descricao", matching=0.97, waiting_time=1000):
                self.type_keys(['ctrl', 'up'])
                self.control_c()
                self.type_keys(['alt', 'tab'])              #A célula é copiada da planilha e colada no local de apontamento
                if not self.find( "Ord_producao", matching=0.97, waiting_time=1000):
                    self.not_found("Ord_producao")
                self.double_click_relative(7, 32)
                self.control_v()
                self.enter()
                
                # Condicionais para o apontamento...

                    # Se caso aparecer uma mensagem de op encerrada ou sem op, voltar na planilha e copiar a próxima da sequência.
                if self.find( "Op_encerrada", matching=0.97, waiting_time=500) or self.find( "Sem_op", matching=0.97, waiting_time=500):
                    self.enter()
                    self.type_keys(['alt', 'tab'])

                    # Caso o contrário, buscar pelo campo Tipo de Movimentação e digitar '010'
                else:
                    self.find( "tipo_mov", matching=0.97, waiting_time=600)
                    self.click_relative(6, 34)
                    self.kb_type('010')
                    self.find( "Salvar", matching=0.97, waiting_time=600)  
                    self.click()

                        # Se faltar saldo, fechar o campo da indicação de saldo e passar para a próxima OP a ser apontada.
                    if self.find( "Falta_saldo", matching=0.97, waiting_time=2800):
                        self.enter()
                        self.find( "Falta_saldo_ok", matching=0.97, waiting_time=700)
                        self.click()
                        self.type_keys(['alt', 'tab'])
                        # Implementar: print da falta de saldo...

                        # Se não, o Apontamento está ok.
                    elif self.find( "documento_vazio", matching=0.97, waiting_time=3200):    # (TIME) PARAMETRO A SER MUDADO CONFORME A MAIOR ORDEM DE PRODUÇÃO A SER APONTADA.
                        self.type_keys(['alt', 'tab'])
                        # Condição para finalizar o apontamento.
                    elif self.find( "final", matching=0.97, waiting_time=3000): #self.find( "fim_apontamento", matching=0.97, waiting_time=5000) and
                        messagebox.showinfo('Apontamentos', 'Apontamentos Finalizados!')
                        a=1
            return 0
                    
        def firmar_ops(op_inicial, op_final):

        #   Rotina para firmar ops previstas.
            #   Abrir ops previstas
            self.find( "atualizacao", matching=0.97, waiting_time=10000)
            self.click()
            self.find( "previstas", matching=0.97, waiting_time=10000)
            self.click()

            #    Digitar ops a serem firmadas...

            self.find( "op_inicio", matching=0.97, waiting_time=10000)
            self.click_relative(206, 12)
            self.paste(op_inicial)
            self.tab()
            self.paste(op_final)
            self.find( "ok_firmar", matching=0.97, waiting_time=10000)
            self.click()

            #   Marcando as OPs

            if not self.find( "Nmr_Op", matching=0.97, waiting_time=10000):
                self.not_found("Nmr_Op")
            self.double_click_relative(-102, 0)
            if not self.find( "marcadas", matching=0.97, waiting_time=100000):
                self.not_found("marcadas")
            if not self.find( "Centro_custo", matching=0.97, waiting_time=10000):
                self.not_found("Centro_custo")
            self.double_click()
            self.wait(1000)
            self.type_keys(['ctrl','home'])


            #  Desmarcar as ops que não serão excluidas.

            while not self.find( "111105", matching=0.97, waiting_time=800):        # Procurando pelo C.Custo dos puxadores
                self.page_down()
            self.find( "111105", matching=0.97, waiting_time=10000)
            self.click()

            #   Centro de custo desmarcados

            while self.find( "marcar_111105", matching=0.97, waiting_time=800):     # 111105 - Puxadores.
                self.enter()
                self.type_down()
            while self.find( "click_111106", matching=0.97, waiting_time=800):      # 111106 - Montagem.
                self.enter()
                self.type_down()
            while self.find( "111108", matching=0.97, waiting_time=10000):          # 111108 - Expedição.
                self.enter()
                self.type_down()


            while not self.find( "111112", matching=0.97, waiting_time=900):        # Procurando pelo C.Custo da elétrica.
                self.page_down()
            self.find( "111112", matching=0.97, waiting_time=1000)
            self.click()

            while self.find( "marcar_111112", matching=0.97, waiting_time=1000):      # 111112 - Elétrica.
                self.enter()
                self.type_down()


            while self.find( "121108", matching=0.97, waiting_time=1000):             # 121108 - Almoxarifado.
                self.enter()
                self.type_down()

            #  Excluindo Ops que não utilizaremos...

            if not self.find( "Numero_da_op", matching=0.97, waiting_time=10000):
                self.not_found("Numero_da_op")
            self.double_click()
            self.wait(2500)
            self.find( "Outras_acoes", matching=0.97, waiting_time=10000)
            self.click()
            # self.find( "Excluir_ops", matching=0.97, waiting_time=10000)
            # self.click()
            # self.find( "Sim_", matching=0.97, waiting_time=10000)
            #self.click()
            

        # Utilização das Rotinas.

        iniciar_siga('PCP4','15450616','10')
        #firmar_ops('A00X0401001', 'A00X4401zzz')
        #abrir_produção()
        #apontar_op()
        #relatorio_estoque('77')

    def not_found(self, label):
        print(f"Element not found: {label}")
if __name__ == '__main__':
    Bot.main()




