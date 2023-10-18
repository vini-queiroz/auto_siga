from botcity.core import DesktopBot
from time import sleep

class Bot(DesktopBot):
    def action(self, execution=None):
        def iniciar_siga(meu_usuario, senha, amb):      # Rotina para iniciar o micro siga
            self.execute(r'C:\Users\Antonio\Desktop\SmartClient R33 Prod.lnk')    # Abrir o atalho para o siga
            if not self.find( "ok", matching=0.97, waiting_time=100000):    # Clicar no botão ok
                self.not_found("ok")
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
                self.tab(wait=5)
            self.kb_type(amb)
            for i in range(2):      # Tecla tab 2 vezes, depois o enter.
                self.tab(wait=5)
            self.enter()        #rotu#
        def abrir_produção():       # Abre o ambiente de apontamento de produção
            if not self.find( "atualizacao", matching=0.97, waiting_time=100000):       # Clicar em Atualizações
                self.not_found("atualizacao")
            self.click()
            if not self.find( "producao", matching=0.97, waiting_time=100000):      # Clicar em Produção
                self.not_found("producao")
            self.click()
            if not self.find( "incluir", matching=0.97, waiting_time=10000):
                self.not_found("incluir")
            self.click()
            self.wait(2000)
            
            ## Após a rotina finalizada o ambiente estará pronta para realizar o apontamento das OPs




        def apontar_op(): # Rotina para apontar as ordens de produção
            self.type_keys(['ctrl', 'alt', 'tab'])
            if not self.find( "apontamentos", matching=0.97, waiting_time=10000):       # Abri a planilha de apontamentos
                self.not_found("apontamentos")
            self.click_relative(125, 122)

            # Ponto de melhoria... (Condicional para comecar e terminar)

            i = 2               # Looping para copiar e apontar cada OP da planilha
            while i > 1:
                self.type_keys(['ctrl', 'up'])
                self.control_c()
                self.type_keys(['alt', 'tab'])
                if not self.find( "campo_op", matching=0.97, waiting_time=10000):
                    self.not_found("campo_op")
                self.click_relative(7, 30)
                self.control_v()
                self.enter()
                if self.find( "op_encerrada", matching=0.97, waiting_time=10000):
                    self.enter()
                    self.key_esc()
                    #self.not_found("op_encerrada")
                else:
                    if not self.find( "tp_movimento", matching=0.97, waiting_time=10000):
                        self.not_found("tp_movimento")
                    self.click_relative(7, 31)
                    self.kb_type('010')
                                         
                if self.find( "falta_saldo", matching=0.97, waiting_time=1000):
                    self.enter()
                    self.key_esc()
                    #self.not_found("falta_saldo")
                      
                if not self.find( "salvar", matching=0.97, waiting_time=10000):
                    self.not_found("salvar")
                self.click()
                if not self.find( "producoes", matching=0.97, waiting_time=10000):
                    self.not_found("producoes")
                self.wait(200)
                if not self.find( "incluir_producoes", matching=0.97, waiting_time=10000):
                    self.not_found("incluir_producoes")
                self.wait(200)
                self.type_keys(['alt', 'tab'])
                
                # for i in range(3):
                     #self.enter()
                
                
                
        #iniciar_siga('PCP4','15450616','10')
        #abrir_produção()
        apontar_op()



    def not_found(self, label):
        print(f"Element not found: {label}")
if __name__ == '__main__':
    Bot.main()





