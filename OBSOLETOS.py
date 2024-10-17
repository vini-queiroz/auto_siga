def relatorio_estoque(armazem):
    if not self.find("saldos_estoques", matching=0.97, waiting_time=1500):
        if self.find("relatorios", matching=0.97, waiting_time=1500):
            self.click()
        else:
            self.not_found("relatorios")
    self.find("saldos_estoques", matching=0.97, waiting_time=1500)
    self.click()
    self.find("planilha", matching=0.97, waiting_time=10000)
    self.click()
    self.find("tipo_planilha", matching=0.97, waiting_time=10000)
    self.click_relative(183, 24)
    self.find("linhas_brancas", matching=0.97, waiting_time=10000)
    self.click()
    self.find("outras_acoes", matching=0.97, waiting_time=10000)
    self.click()
    self.find("parametros", matching=0.97, waiting_time=10000)
    self.click()
    self.find("do_armazem", matching=0.97, waiting_time=10000)
    self.click_relative(210, 11)
    self.kb_type(armazem)
    self.kb_type(armazem)
    self.find("ok", matching=0.97, waiting_time=10000)
    self.click()
    self.find("imprimir", matching=0.97, waiting_time=10000)
    self.click()
    self.find("abrir", matching=0.97, waiting_time=10000)
    self.click()
    self.find("btn_siim", matching=0.97, waiting_time=10000)
    self.click()


def iniciar_siga(meu_usuario, senha, amb):
    # Rotina para iniciar o micro siga
    self.execute(r'C:\Users\Antonio\Desktop\SmartClient R33 Prod.lnk')  # Abrir o atalho para o siga
    if not self.find("btn_ok", matching=0.97, waiting_time=10000):
        self.not_found("btn_ok")
    self.click()
    if not self.find("Boas_vindas", matching=0.97, waiting_time=100000):
        self.not_found("Boas_vindas")
    self.kb_type(meu_usuario)  # Digitar usuário
    self.tab()
    self.kb_type(senha)  # Digitar senha
    self.enter()
    if not self.find("Ambiente", matching=0.97, waiting_time=10000):
        self.not_found("Ambiente")
    for i in range(2):  # Tecla tab 2 vezes, depois o enter.
        self.tab()
    self.kb_type(amb)
    self.find("btn_entrar", matching=0.97, waiting_time=10000)
    self.click()


def apontar_op_backup():
    # Rotina para apontar as ordens de produção
    self.type_keys(['ctrl', 'alt', 'tab'])
    if not self.find("TOTVS_encontrar", matching=0.97, waiting_time=10000):
        self.not_found("TOTVS_encontrar")
    self.click_relative(114, 61)
    self.wait(500)
    self.type_keys(['ctrl', 'alt', 'tab'])
    if not self.find("Apontamentos", matching=0.97, waiting_time=10000):
        self.not_found("Apontamentos")
    self.click_relative(121, 100)

    if not self.find("Prod_excel", matching=0.97, waiting_time=10000):
        self.not_found("Prod_excel")
    self.click()  # Alterna para a planilha excel e procura a ultima OP
    self.type_keys(['ctrl', 'down'])
    self.type_keys(['ctrl', 'left'])

    # Looping para apontamento das OPs.
    a = 0
    while a < 1:
        self.type_keys(['ctrl', 'up'])
        self.control_c()
        self.type_keys(['alt', 'tab'])  # A célula é copiada da planilha e colada no local de apontamento
        if not self.find("Ord_producao", matching=0.97, waiting_time=1000):
            self.not_found("Ord_producao")
        self.double_click_relative(7, 32)
        self.control_v()
        self.enter()

        # Condicionais para o apontamento...

        # Se caso aparecer uma mensagem de op encerrada ou sem op, voltar na planilha e copiar a próxima da sequência.
        if self.find("Op_encerrada", matching=0.97, waiting_time=300) or self.find("Sem_op", matching=0.97,
                                                                                   waiting_time=300):
            self.enter()
            self.type_keys(['alt', 'tab'])

            # Caso o contrário, buscar pelo campo Tipo de Movimentação e digitar '010'
        else:
            self.find("tipo_mov", matching=0.97, waiting_time=300)
            self.click_relative(6, 34)
            self.kb_type('010')
            self.find("Salvar", matching=0.97, waiting_time=300)
            self.click()
            # Se faltar saldo, fechar o campo da indicação de saldo e passar para a próxima OP a ser apontada.
            # if self.find( "Falta_saldo", matching=0.97, waiting_time=400):
            #    self.wait(500)
            #    self.enter()
            #    self.find( "Falta_saldo_ok", matching=0.97, waiting_time=700)
            #    self.click()
            #    self.type_keys(['alt', 'tab'])
            # Implementar: print da falta de saldo...

            if self.find("No_saldo", matching=0.97, waiting_time=3000):
                self.enter()
                if not self.find("fechar_saldo", matching=0.97, waiting_time=300):
                    self.not_found("fechar_saldo")
                self.click()
                self.type_keys(['alt', 'tab'])

                # Se não, o Apontamento está ok.
            elif self.find("documento_vazio", matching=0.97,
                           waiting_time=7000):  # (TIME) PARAMETRO A SER MUDADO CONFORME A MAIOR ORDEM DE PRODUÇÃO A SER APONTADA.
                self.type_keys(['alt', 'tab'])
                # Condição para finalizar o apontamento.
            elif self.find("Finalizado", matching=0.97, waiting_time=10000):
                messagebox.showinfo('Apontamentos', 'Apontamentos Finalizados!')
                a = 1

        def relatorio_SD4(op_inicial, op_final):
            if not self.find("amb_genericos", matching=0.97, waiting_time=1500):
                if self.find("btn_consultas", matching=0.97, waiting_time=1500):
                    self.click()
                else:
                    self.not_found("btn_consultas")
            self.find("amb_genericos", matching=0.97, waiting_time=1500)
            self.click()
            if not self.find("campo_pesquisa", matching=0.97, waiting_time=100000):
                self.not_found("campo_pesquisa")
            self.click_relative(82, 6)
            self.kb_type('SD4')
            self.enter()
            if not self.find("amb_requisicoes", matching=0.97, waiting_time=10000):
                self.not_found("amb_requisicoes")
            self.click()
            self.double_click()
            if not self.find("bnt_filtrar", matching=0.97, waiting_time=10000):
                self.not_found("bnt_filtrar")
            self.click()
            if not self.find("btn_criar_filtro", matching=0.97, waiting_time=10000):
                self.not_found("btn_criar_filtro")
            self.click()
            if not self.find("campo_campo", matching=0.97, waiting_time=10000):
                self.not_found("campo_campo")
            self.click_relative(135, 31)

            if not self.find("campo_produto", matching=0.97, waiting_time=10000):
                self.not_found("campo_produto")
            self.click()
            if not self.find("campo_operador", matching=0.97, waiting_time=10000):
                self.not_found("campo_operador")
            self.click_relative(171, 27)
            if not self.find("campo_nao_contem", matching=0.97, waiting_time=10000):
                self.not_found("campo_nao_contem")
            self.click()
            if not self.find("campo_expressao", matching=0.97, waiting_time=10000):
                self.not_found("campo_expressao")
            self.click_relative(9, 27)
            self.kb_type('MOD')
            if not self.find("btn_adicionar", matching=0.97, waiting_time=10000):
                self.not_found("btn_adicionar")
            self.click()
            if not self.find("btn_e", matching=0.97, waiting_time=10000):
                self.not_found("btn_e")
            self.click()
            if not self.find("campo_campo", matching=0.97, waiting_time=10000):
                self.not_found("campo_campo")
            self.click_relative(135, 31)
            if not self.find("campo_saldo_empenho", matching=0.97, waiting_time=10000):
                self.not_found("campo_saldo_empenho")
            self.click()
            if not self.find("campo_operador", matching=0.97, waiting_time=10000):
                self.not_found("campo_operador")
            self.click_relative(171, 27)
            if not self.find("campo_diferente_de", matching=0.97, waiting_time=10000):
                self.not_found("campo_diferente_de")
            self.click()
            if not self.find("btn_adicionar", matching=0.97, waiting_time=10000):
                self.not_found("btn_adicionar")
            self.click()
            if not self.find("btn_salvar", matching=0.97, waiting_time=10000):
                self.not_found("btn_salvar")
            self.click()
            if not self.find("selecionar_filtro", matching=0.97, waiting_time=10000):
                self.not_found("selecionar_filtro")
            self.click_relative(-17, -1)
            if not self.find("btn_aplicar_filtro", matching=0.97, waiting_time=10000):
                self.not_found("btn_aplicar_filtro")
            self.click()
            messagebox.askyesno("Aguardando Filtro", "O filtro já foi processado???")

            def selecionar_dicionario():
                if not self.find("btn_dicionario", matching=0.97, waiting_time=10000):
                    self.not_found("btn_dicionario")
                self.click()
                if not self.find("campo_marcar", matching=0.97, waiting_time=10000):
                    self.not_found("campo_marcar")
                self.move()
                self.mouse_down()
                self.mouse_up()
                if not self.find("campo_marcado", matching=0.97, waiting_time=10000):
                    self.not_found("campo_marcado")
                self.move()
                self.click()
                if not self.find("btn_filial", matching=0.97, waiting_time=10000):
                    self.not_found("btn_filial")
                self.click_relative(-70, 2)

                for i in range(0, 5):
                    self.enter()
                    self.type_down()
                self.type_down()
                self.type_down()
                for i in range(0, 2):
                    self.enter()
                    self.type_down()
                self.page_down()
                if not self.find("campo_OP_origem", matching=0.97, waiting_time=10000):
                    self.not_found("campo_OP_origem")
                self.click_relative(-69, 2)
                self.enter()
                self.page_down()
                self.page_down()
                if not self.find("Produto_pai", matching=0.97, waiting_time=10000):
                    self.not_found("Produto_pai")
                self.click_relative(-70, 1)
                self.enter()

            selecionar_dicionario()
            if not self.find("botao_ok", matching=0.97, waiting_time=10000):
                self.not_found("botao_ok")
            self.click()
            self.wait(1000)
            if not self.find("exportar_csv", matching=0.97, waiting_time=10000):
                self.not_found("exportar_csv")
            self.click()
            if not self.find("btn_formato", matching=0.97, waiting_time=10000):
                self.not_found("btn_formato")
            self.click()
            if not self.find("formato_xml", matching=0.97, waiting_time=10000):
                self.not_found("formato_xml")
            self.click()
            if not self.find("btn_confirmar", matching=0.97, waiting_time=10000):
                self.not_found("btn_confirmar")
            self.click()