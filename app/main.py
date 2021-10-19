from PyQt5 import uic, QtWidgets
from datetime import date, datetime
import win32com.client as win32
from openpyxl import Workbook
from emaildef import send
from calc import calculo


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
        self.tela_email.btn_enviar.clicked.connect(send.Enviar_Email) # envia email

        ### TELA CALCULO BOTOES 
        self.tela_calculo.btn_calcular.clicked.connect(calculo.Calculodora_de_Estatistica) # botao calcular
        self.tela_calculo.btn_voltar.clicked.connect(self.voltar_calculo) # botao voltar

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
    def voltar_email(self):
        print('Voltar')

        # Fecha Email e abre Geral
        self.tela_email.close()
        self.tela_geral.show()


    ### CALCULO
    def voltar_calculo(self):
        print('Voltar')
        
        # Fecha Calculo e abre Geral
        self.tela_calculo.close()
        self.tela_geral.show()


SistemaGeral()


### jose.marinho56@gmail.com