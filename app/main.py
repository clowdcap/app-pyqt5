from PyQt5 import uic, QtWidgets
from datetime import date
import win32com.client as win32
from openpyxl import Workbook



class SistemaGeral:
    ### INICIO
    def __init__(self):
        ### SETUP LAYOUT APP
        # Chamando sistema Layout
        app = QtWidgets.QApplication([])

        # Layout's
        self.tela_login = uic.loadUi('./pyqt5-templates/login.ui')
        self.tela_geral = uic.loadUi('./pyqt5-templates/geral.ui')
        self.tela_projetos = uic.loadUi('./pyqt5-templates/projetos.ui')
        self.tela_email = uic.loadUi('./pyqt5-templates/email.ui')
        self.tela_calculo = uic.loadUi('./pyqt5-templates/calculo.ui')
        self.tela_analise = uic.loadUi('./pyqt5-templates/analise.ui')

        # Layouts's Projetos        
        self.tela_projetos_resumo = uic.loadUi('./pyqt5-templates/projetos-resumo.ui')

        # Executar funções ao chamar
        ### TELA LOGIN BOTOES
        self.tela_login.btn_login.clicked.connect(self.logar_login) # botao logar
        self.tela_login.btn_cadastrar.clicked.connect(self.cadastrar_login) # botao cadastrar        
        
        ### TELA GERAL BOTOES
        ### Linha 1
        self.tela_geral.btn_projetos.clicked.connect(self.projetos_geral) # botao projetos
        self.tela_geral.btn_email.clicked.connect(self.email_geral) # botao email
        self.tela_geral.btn_calc.clicked.connect(self.calculo_geral) # botao calculo
        
        ### Linha 2
        self.tela_geral.btn_analise.clicked.connect(self.analise_geral) # botao analise
        
        self.tela_geral.btn_voltar.clicked.connect(self.voltar_geral) # botao voltar

        ### TELA PROJETOS BOTOES
        self.tela_projetos.btn_resumo.clicked.connect(self.resumo_projetos) # botao calculo
        self.tela_projetos.btn_situacao.clicked.connect(self.situacao_projetos) # botao situacao
        self.tela_projetos.btn_cadastro.clicked.connect(self.cadastros_projetos) # botao cadastro
        self.tela_projetos.btn_anterior.clicked.connect(self.anteriores_projetos) # botao anterior
        self.tela_projetos.btn_entrega.clicked.connect(self.entrega_projetos) # botao entrega
        self.tela_projetos.btn_voltar.clicked.connect(self.voltar_projetos) # botao voltar

        ### TELA RESUMO BOTOES
        self.tela_projetos_resumo.btn_voltar.clicked.connect(self.voltar_projetos_resumo) # botao voltar

        ### TELA E-MAIL BOTOES
        self.tela_email.btn_voltar.clicked.connect(self.voltar_email) # botao voltar
        self.tela_email.btn_enviar.clicked.connect(self.enviar_email) # envia email

        ### TELA CALCULO BOTOES 
        self.tela_calculo.btn_calcular.clicked.connect(self.calcular_estatisticas) # botao calcular
        self.tela_calculo.btn_voltar.clicked.connect(self.voltar_calculo) # botao voltar

        ### TELA ANALSIE BOTOES
        self.tela_analise.btn_voltar.clicked.connect(self.voltar_analise) # botao voltar

        ### SETUP

        # Mostrar layout
        self.tela_login.show()

        # Executar sistema
        app.exec_()



    ### LOGIN
    def logar_login(self):
        # Coleta dados dos input
        login = self.tela_login.input_login.text()
        senha = self.tela_login.input_senha.text()
        print(f'Login: {login}\nSenha: {senha}')
        
        # Conferencia do login e senha
        if login == '' and senha == '':
            print('Login Autorizado')
            self.tela_login.close()
            self.tela_geral.show()
        else: 
            print('Erro no login')

    def cadastrar_login(self):
        print('Cadastrar')



    ### GERAL         
    def projetos_geral(self):
        print('Projetos')
        
        # Fecha Geral e abre Projetos
        self.tela_geral.close()
        self.tela_projetos.show()

    def email_geral(self):
        print('Email')

        # Fecha Geral e abre Email
        self.tela_geral.close()
        self.tela_email.show()

    def calculo_geral(self):
        print('Calculo de Estatistica')

        # Fecha Geral e abre Calculo
        self.tela_geral.close()
        self.tela_calculo.show()

    def analise_geral(self):
        print('Analise')

        # Fecha Geral e abre Analise
        self.tela_geral.close()
        self.tela_analise.show()    

    def voltar_geral(self):
        print('Deslogar')

        # Fecha Geral e abre Login
        self.tela_geral.close()
        self.tela_login.show()



    ### PROJETOS
    def resumo_projetos(self):
        print('Resumo')

        # Fecha Projetos e abre Resumo
        self.tela_projetos.close()
        self.tela_projetos_resumo.show()

    def situacao_projetos(self):
        print('Situação')

    def cadastros_projetos(self):
        print('Cadastros')

    def anteriores_projetos(self):
        print('Anteriores')

    def entrega_projetos(self):
        print('Entrega')  
   
    def voltar_projetos(self):
        print('Voltar')

        # Fecha Projetos e abre Geral
        self.tela_projetos.close()
        self.tela_geral.show()



    ### PROJETO - RESUMO
    def voltar_projetos_resumo(self):
        print('Voltar')

        # Fecha Resumo e abre Projeto
        self.tela_projetos_resumo.close()
        self.tela_projetos.show()



    ### EMAIL
    def enviar_email(self):
        
            self.tela_email = uic.loadUi('./pyqt5-templates/email.ui')
            print('E-Mail')
            
            # criar integração com o outlook
            outlook = win32.Dispatch('outlook.application')

            # criar um email
            email = outlook.CreateItem(0)     
            
            # configurar as informações
            email.To = 'jose.marinho56@gmail.com'
            email.Subject = 'Prefeitura Municipal de Campo Magro - Modelo Carimbo'
            
            # variváveis
            nome = self.tela_email.input_nome.text()
            sobrenome = self.tela_email.input_snome.text()
            # adicionando anexo
            anexo = r'E:\Python\app-pyqt5\app\emaildef\anexos\legenda-pmcm.dwg'
            
            try:
                email.Attachments.Add(anexo)
                print('Anexado documento...')
            except:
                email.Attachments.Add(anexo)
                print('Sem Anexo...')
            
            css = '''
                        <style>
                        * {
                            margin: 0;
                            padding: 0;
                            box-sizing: border-box;
                        }

                        body {
                            width: 100%;
                            overflow-x: hidden;
                        }

                        .email {
                            margin: 2%;
                        }

                        .topo {
                            padding: 0.3rem 0; 
                            background-color: brown;
                            width: 100%;
                            text-align: center !important;
                            margin-bottom: 35px;
                            font-family: Arial, Helvetica, sans-serif;
                        }

                        .topo h2 {
                            font-size: 28px;
                            color: white;
                            padding-top: 25px;
                        }

                        img {
                            width: 80px;
                            height: 80px;
                        }

                        .capa {
                            text-align: center;
                            width: 100%;
                            padding: 0.5rem 0;
                            background-color: cadetblue;
                        }
                        .capa p {
                            font-size: 16px;
                            font-family: Arial, Helvetica, sans-serif;
                            color: white;
                            text-align: left !important;
                            padding: 0 30px;
                        }

                        .conteudo {
                            text-align: center;
                            width: 100%;
                            padding: 2rem;
                            background-color: gainsboro;
                        }

                        .conteudo p{
                            font-size: 20px;
                            font-family: Arial, Helvetica, sans-serif;
                        }

                        .conteudo p a {
                            text-decoration: none;
                            color: red;
                        }

                        .assinatura {
                            background-color: royalblue;
                            color: white;
                            font-family: Arial, Helvetica, sans-serif;
                            margin-top: 20px;
                            padding: 2%;
                        }
                       </style>
                        '''

            assinatura = '''
                <div class="assinatura">
                    <p>Atenciosamente,</p>
                        <br>
                        <br>
                    <h3>Jose Marinho</h3>
                        <br>
                    <p><a style="color: white;" href="https://wa.me/qr/LQM5O2QPPRDOH1">Whatsapp: (41) 9 9272-5388</a></p>
                    <p>Telefone: (41) 3677-4000 - Central Prefeitura</p>
                    <p>Telefone: (41) 3677-4050 - SEDUA</p>
                    <p style="color: white;">jm.arquiteturacwb@gmail.com</p>
                    <p>Prefeitura Municipal de Campo Magro / PR</p>
                </div>
                '''

            topo = '''
                <div class="topo">
                    <img src="https://leismunicipais.com.br/img/cidades/pr/campo-magro.png" alt="campo-magro">
                    <h2>Prefeitura Municipal de Campo Magro - SEDUA</h2>
                </div> <!--topo-->
                '''

            conteudo = f'''
                <div class="conteudo">
                    <p>Bom dia {nome} {sobrenome},</p>
                    <br>
                    <p>Entro em contato para atender a sua solicitação</p> 
                    <br>
                    <p>Está anexado a esse mensagem, um arquivo em dwg, onde o mesmo contém a estrutura de carimbo padrão da prefeitura, logo, a tabela de estatística está junto.</p> 
                    <br>
                    <p>Nome do arquivo: <b>legenda-pmcm.dwg</b></p>
                    <p>Tamanho do arquivo: <b>75,3 KB</b></p>
                    <p>Caso eu não tenha esclarecido totalmente a sua dúvida, estou à disposição</p>
                </div>
                '''
            
            capa = '''
                <div class="capa">
                    <p>Atendimento via E-mail - A/C: <b>José Marinho - Estagiário</b></p>
                </div> <!--capa-->
                '''

            email.HTMLBody = f'''
                <!DOCTYPE html>
                <html>
                    <head>
                        <meta charset="utf-8">
                        <meta http-equiv="X-UA-Compatible" content="IE=edge">
                        <meta name="viewport" content="width=device-width, initial-scale=1">
                        {css}
                    </head>
                    <body>
                        <section class="email">

                            {topo}

                            {capa}

                            {conteudo}

                            {assinatura}

                        </section>
                    </body>
                </html>

                '''
            # body
      
            # finalizando email
            email.Send()

    def voltar_email(self):
        print('Voltar')

        # Fecha Email e abre Geral
        self.tela_email.close()
        self.tela_geral.show()



    ### CALCULO
    def calcular_estatisticas(self):
        print('Calcular')
        
        # Setup inicial
        # __init__
        arquivo_excel = Workbook()
        planilha1 = arquivo_excel.active
        data_atual = date.today()
    
        # Coleta dados dos input e arnazena em variáveis
        area_do_terreno = float(self.tela_calculo.input_area_do_terreno.text())
        area_anteriormente_construido = float(self.tela_calculo.input_area_anteriormente_construido.text())
        area_computavel_subsolo = float(self.tela_calculo.input_area_computavel_a_construir_no_subsolo.text())
        area_nao_computavel_subsolo = float(self.tela_calculo.input_area_nao_computavel_a_construir_no_subsolo.text())
        area_computavel_terreo = float(self.tela_calculo.input_area_computavel_a_construir_terreo.text())
        area_nao_computavel_terreo = float(self.tela_calculo.input_area_nao_computavel_a_construir_terreo.text())
        area_computavel_superior = float(self.tela_calculo.input_area_computavel_a_construir_pavSup.text())
        area_nao_computavel_superior = float(self.tela_calculo.input_area_nao_computavel_a_construir_pavSup.text())
        area_nao_computavel_a_construir_atico = float(self.tela_calculo.input_area_nao_computavel_a_construir_atico.text())
        area_livre = float(self.tela_calculo.input_area_livre.text())
        numero_de_pavimento = int(self.tela_calculo.input_numero_de_pavimento.text())
        altura_total = float(self.tela_calculo.input_altura_total.text())
        area_de_cobertura = float(self.tela_calculo.input_area_de_cobertura.text())

        # Gerando calculos
        projecao_edificacao = area_anteriormente_construido + area_computavel_terreo + area_nao_computavel_terreo
        taxa_ocupacao = (projecao_edificacao / area_do_terreno) * 100 
        taxa_permeabilidade = (area_livre / area_do_terreno) * 100
        area_total_contruir_subsolo = area_computavel_subsolo + area_nao_computavel_subsolo
        area_total_contruir_terreo = area_computavel_terreo + area_nao_computavel_terreo
        area_total_contruir_superior = area_computavel_superior + area_nao_computavel_superior
        area_total_contruir_computavel = area_computavel_subsolo + area_computavel_terreo + area_computavel_superior
        area_total_contruir_nao_computavel = area_nao_computavel_subsolo + area_nao_computavel_terreo + area_nao_computavel_superior
        coeficiente_aproveitamento = ((area_total_contruir_computavel + area_anteriormente_construido)/area_do_terreno)
        area_total_construida_liberada = area_total_contruir_computavel + area_total_contruir_nao_computavel

        dicionario_com_resultado = {
            'Area do terreno': area_do_terreno,
            'Area anteriormente construido' : area_anteriormente_construido,
            'Area computavel a construir no subsolo': area_computavel_subsolo,
            'Area nao computavel a construir no subsolo': area_nao_computavel_subsolo,
            'Area computavel a construir no pavimento terreo': area_computavel_terreo,
            'Area nao computavel a construir no no pavimento terreo': area_nao_computavel_terreo,
            'Area computavel a construir no no pavimento superior': area_computavel_superior,
            'Area nao computavel a construir no no pavimento superior': area_nao_computavel_superior,
            'Area nao computavel a construir no atico': area_nao_computavel_a_construir_atico,
            'Area livre': area_livre,
            'Numero de pavimento': numero_de_pavimento,
            'Altura total': altura_total,
            'Projecao da edificacao': projecao_edificacao,
            'Taxa de ocupacao': taxa_ocupacao,
            'Taxa de permeabilidade': taxa_permeabilidade,
            'Area total a contruir no subsolo': area_total_contruir_subsolo,
            'Area total a contruir no pavimneto terreo': area_total_contruir_terreo,
            'Area total a contruir no pavimneto superior': area_total_contruir_superior,
            'Area total a contruir computavel': area_total_contruir_computavel,
            'Area total a contruir nao computavel': area_total_contruir_nao_computavel,
            'Coeficiente de aproveitamento': coeficiente_aproveitamento,
            'Area de cobertura': area_de_cobertura,
            'Area total a ser construida': area_total_construida_liberada
        }

        # Coletando dados input do protocolo e do interessado
        numero_do_protocolo = str(self.tela_calculo.input_protocolo.text())
        interessado_projeto = str(self.tela_calculo.input_interessado.text())
        
        # Colunas
        planilha1['A2'] = 'Item'
        planilha1['B2'] = 'Descrição'
        planilha1['C2'] = 'Dado'
        planilha1['D2'] = 'Unidade'
        planilha1['A1'] = 'Registro:'
        planilha1['B1'] = data_atual
        planilha1['C1'] = f'Protocolo {numero_do_protocolo}'
        planilha1['D1'] = f'Interessado {interessado_projeto}'
        
        # Passar por todos os dados do dicionarios e adicionar em linhas na planilha
        for item, descricao in enumerate(dicionario_com_resultado):
            linha = (item+1, descricao, dicionario_com_resultado[descricao])
            planilha1.append(linha)

        # Retornando dados
        # Colocando unidades nas linhas
        planilha1['D3'] = 'M²' # Item 1
        planilha1['D4'] = 'M²' # Item 2
        planilha1['D5'] = 'M²' # Item 3
        planilha1['D6'] = 'M²' # Item 4
        planilha1['D7'] = 'M²' # Item 5
        planilha1['D8'] = 'M²' # Item 6
        planilha1['D9'] = 'M²' # Item 7
        planilha1['D10'] = 'M²' # Item 8
        planilha1['D11'] = 'M²' # Item 9
        planilha1['D12'] = 'M²' # Item 10
        planilha1['D13'] = 'Pavimentos' # Item 11
        planilha1['D14'] = 'M' # Item 12
        planilha1['D15'] = 'M²' # Item13
        planilha1['D16'] = '%' # Item 14
        planilha1['D17'] = '%' # Item 15
        planilha1['D18'] = 'M²' # item 16
        planilha1['D19'] = 'M²' # item 17
        planilha1['D20'] = 'M²' # item 18
        planilha1['D21'] = 'M²' # item 19
        planilha1['D22'] = 'M²' # item 20
        planilha1['D23'] = '' # Item 21
        planilha1['D24'] = 'M²' # Item 22
        planilha1['D25'] = 'M²' # Item 23

        try:
            arquivo_excel.save(f"./calc/relatorios/Relatorio {numero_do_protocolo} - Estatístico.xlsx")
            print('Relatorio gerado com sucesso')
        except:
            print("Erro ao salvar o Relatório.")
    
    def voltar_calculo(self):
        print('Voltar')
        
        # Fecha Calculo e abre Geral
        self.tela_calculo.close()
        self.tela_geral.show()



    ### ANALISES
    def voltar_analise(self):
        print('Voltar')

        # Fecha Analise e abre Geral
        self.tela_analise.close()
        self.tela_geral.show()

SistemaGeral()
### jose.marinho56@gmail.com
