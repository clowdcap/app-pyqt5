from PyQt5 import uic, QtWidgets
from openpyxl import Workbook
import datetime

class SistemaGeral:
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

        # Layouts's Projetos        
        self.tela_projetos_resumo = uic.loadUi('./pyqt5-templates/projetos-resumo.ui')


        # Executar funções ao chamar
        ### TELA LOGIN BOTOES
        self.tela_login.btn_login.clicked.connect(self.logar_login) # botao logar
        self.tela_login.btn_cadastrar.clicked.connect(self.cadastrar_login) # botao cadastrar        
        
        ### TELA GERAL BOTOES
        self.tela_geral.btn_projetos.clicked.connect(self.projetos_geral) # botao projetos
        self.tela_geral.btn_email.clicked.connect(self.email_geral) # botao email
        self.tela_geral.btn_calc.clicked.connect(self.calculo_geral) # botao calculo
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

        ### TELA CALCULO BOTOES
        self.tela_calculo.btn_calcular.clicked.connect(self.calcular_calculo) # botao calcular
        self.tela_calculo.btn_voltar.clicked.connect(self.voltar_calculo) # botao voltar

        ### SETUP

        # Mostrar layout
        self.tela_login.show()

        # Executar sistema
        app.exec_()
    

    
    ### LOGIN
    def logar_login(self):
        login = self.tela_login.input_login.text()
        senha = self.tela_login.input_senha.text()
        print(f'Login: {login}\nSenha: {senha}')
        
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
        self.tela_geral.close()
        self.tela_projetos.show()
        
    def email_geral(self):
        print('Email')
        self.tela_geral.close()
        self.tela_email.show()
        
    def calculo_geral(self):
        print('Calculo de Estatistica')
        self.tela_geral.close()
        self.tela_calculo.show()
        
    def voltar_geral(self):
        print('Deslogar')
        self.tela_geral.close()
        self.tela_login.show()


    ### PROJETOS
    def resumo_projetos(self):
        print('Resumo')
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
        self.tela_projetos.close()
        self.tela_geral.show()


    ### PROJETO - RESUMO
    def voltar_projetos_resumo(self):
        print('Voltar')
        self.tela_projetos_resumo.close()
        self.tela_projetos.show()


    ### EMAIL
    def voltar_email(self):
        print('Voltar')
        self.tela_email.close()
        self.tela_geral.show()


    ### CALCULO
    def calcular_calculo(self):
        print('Calcular')

        # Setup inicial
        # __init__
        arquivo_excel = Workbook()
        planilha1 = arquivo_excel.active

        # Colunas
        planilha1['A2'] = 'Item'
        planilha1['B2'] = 'Descrição'
        planilha1['C2'] = 'Dado'
        planilha1['D2'] = 'Unidade'
        planilha1['A1'] = 'Registro:'
        planilha1['B1'] = datetime.datetime.now()
        
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

        if area_computavel_subsolo == 0 and area_nao_computavel_subsolo == 0:
            if numero_de_pavimento < 0 or numero_de_pavimento > 2:
                print('Numero de pavimentos fora do permitido')

        if area_de_cobertura < (area_de_cobertura*0.9):
            print('Rever tamanho da cobertura')

        dicionario_com_resultado = {
            'Area do terreno': area_do_terreno,
            'Area anteriormente construido' : area_anteriormente_construido,
            'Area computavel subsolo': area_computavel_subsolo,
            'Area nao computavel subsolo': area_nao_computavel_subsolo,
            'Area computavel terreo': area_computavel_terreo,
            'Area nao computavel terreo': area_nao_computavel_terreo,
            'Area computavel superior': area_computavel_superior,
            'Area nao computavel superior': area_nao_computavel_superior,
            'Area nao computavel a construir atico': area_nao_computavel_a_construir_atico,
            'Area livre': area_livre,
            'Numero de pavimento': numero_de_pavimento,
            'Altura total': altura_total,
            'Projecao da edificacao': projecao_edificacao,
            'Taxa de ocupacao': taxa_ocupacao,
            'Taxa de permeabilidade': taxa_permeabilidade,
            'Area total a contruir subsolo': area_total_contruir_subsolo,
            'Area total a contruir terreo': area_total_contruir_terreo,
            'Area total a contruir superior': area_total_contruir_superior,
            'Area total a contruir computavel': area_total_contruir_computavel,
            'Area total a contruir nao computavel': area_total_contruir_nao_computavel,
            'Coeficiente de aproveitamento': coeficiente_aproveitamento,
            'Area de cobertura': area_de_cobertura,
            'Area total construida liberada': area_total_construida_liberada
        }

        for item, descricao in enumerate(dicionario_com_resultado):
            linha = (item+1, descricao, dicionario_com_resultado[descricao])
            planilha1.append(linha)
        
        # Retornando dados
        # Colocando unidades nos resultados
        planilha1['D3'] = 'M²' # 1
        planilha1['D4'] = 'M²' # 2
        planilha1['D5'] = 'M²' # 3
        planilha1['D6'] = 'M²' # 4
        planilha1['D7'] = 'M²' # 5
        planilha1['D8'] = 'M²' # 6
        planilha1['D9'] = 'M²' # 7
        planilha1['D10'] = 'M²' # 8
        planilha1['D11'] = 'M²' # 9
        planilha1['D12'] = 'M²' # 10
        planilha1['D13'] = 'Pavimentos' # 11
        planilha1['D14'] = 'M' # 12
        planilha1['D15'] = 'M²' # 13
        planilha1['D16'] = '%' # 14
        planilha1['D17'] = '%' # 15
        planilha1['D18'] = 'M²' # 16
        planilha1['D19'] = 'M²' # 17
        planilha1['D20'] = 'M²' # 18
        planilha1['D21'] = 'M²' # 19
        planilha1['D22'] = 'M²' # 20
        planilha1['D23'] = '' # 21
        planilha1['D24'] = 'M²' # 22
        planilha1['D25'] = 'M²' # 23
        
        try:
            arquivo_excel.save("teste.xlsx")
            print('Relatorio gerado com sucesso')
        except:
            print("Erro ao salvar o Relatório.")

    def voltar_calculo(self):
        print('Voltar')
        self.tela_calculo.close()
        self.tela_geral.show()


SistemaGeral()
