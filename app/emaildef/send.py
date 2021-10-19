import win32com.client as win32
from PyQt5 import uic

class Enviar_Email:
    def __init__(self):
        self.tela_email = uic.loadUi('./pyqt5-templates/email.ui')
        print('E-Mail')
        # criar integração com o outlook
        outlook = win32.Dispatch('outlook.application')

        # criar um email
        email = outlook.CreateItem(0)
        
        
        # configurar as informações
        email.To = 'jose.marinho56@gmail.com'
        email.Subject = 'Teste'
        
        # variváveis
        ### nome = self.tela_email.input_email.text()
        ### sobrenome = str(input('Sobrenome: '))
        # adicionando anexo
        anexo = r'E:\Python\app-pyqt5\app\emaildef\anexos\legenda-pmcm.dwg'
        
        try:
            email.Attachments.Add(anexo)
            print('Anexado documento...')
        except:
            email.Attachments.Add(anexo)
            print('Sem Anexo...')
        
        email.HTMLBody = 'Segue em Anexo'
        
        # finalizando email
        email.Send()

# jose.marinho56@gmail.com
