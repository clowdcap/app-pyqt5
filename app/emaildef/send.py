import win32com.client as win32
from PyQt5 import uic

class Enviar_Email:
    def __init__(self):
        print('E-Mail')
        # criar integração com o outlook
        outlook = win32.Dispatch('outlook.application')

        # criar um email
        email = outlook.CreateItem(0)
        
        # configurar as informações
        email.To = ''
        email.Subject = ''
        
        # variváveis
        ### nome = str(input('Nome: '))
        ### sobrenome = str(input('Sobrenome: '))
        # adicionando anexo
        anexo = r'./anexos/legenda-pmcm.dwg'
        
        try:
            email.Attachments.Add(anexo)
            print('Anexado documento...')
        except:
            print('Sem Anexo...')
        
        email.HTMLBody = 'Segue em Anexo'
        
        # finalizando email
        email.Send()

# jose.marinho56@gmail.com
