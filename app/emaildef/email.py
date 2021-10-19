import win32com.client as win32
from admin import admin


class Email:
    def __init__(self):
        print('Sistema de E-mail')

    def enviar_email(self):
        print('Enviar E-mail')
        # criar integração com o outlook
        outlook = win32.Dispatch('outlook.application')

        # criar um email
        email = outlook.CreateItem(0)

        # configurar as informações
        email.To = ''
        email.Subject = ''

        if email.To == '':
            email.To = str(input('Para quem vai esse e-mail? '))

        if email.Subject == '':
            email.Subject = str(input('Qual o assunto do e-mail? '))

        # variváveis
        nome = str(input('Nome: '))
        sobrenome = str(input('Sobrenome: '))

        if nome == '' and sobrenome == '':
            nome = 'Senhor'
            sobrenome = '(a)'
        if nome == '' and sobrenome != '':
            nome = 'Sr (a).'

              
        #anexo = fr'app\email\anexos\legenda-pmcm.dwg'
        #anexo = fr'./anexos/legenda-pmcm.dwg'
    
        # adicionando anexo
        def anexar_documento(nome_arquivo):
            if nome_arquivo == '':
                anexo = ''
                if nome_arquivo == '':
                    nome_arquivo = str(input('Qual o nome do arquivo?: '))
                    if nome_arquivo != '':
                        anexo = fr'E:\python\system\menu\sistema\files\{nome_arquivo}'
            else:
                anexo = fr'E:\python\system\menu\sistema\files\{nome_arquivo}'

            if anexo == '':
                print('Sem Anexo...')
            else:
                email.Attachments.Add(anexo)
                print('Anexado documento...')


        # CORPO DO EMAIL AUTOMATIZADO
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
                <p>jm.arquiteturacwb@gmail.com</p>
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
                <p>Nome do arquivo: <b>{self.nome_arquivo} </b></p>
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

