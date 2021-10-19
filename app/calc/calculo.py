# AMBIENTE DE TESTE - CALCULADORA
from datetime import date, datetime
from openpyxl import Workbook
from PyQt5 import uic


class Calculodora_de_Estatistica:
    def __init__(self):
        self.tela_calculo = uic.loadUi('./pyqt5-templates/calculo.ui')
        self.calcular_estatisticas()

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
            #arquivo_excel.save(f"./calc/relatorios/Relatorio {numero_do_protocolo} - Estatístico.xlsx")
            print('Relatorio gerado com sucesso')
        except:
            print("Erro ao salvar o Relatório.")
        return arquivo_excel.save(f"./calc/relatorios/Relatorio {numero_do_protocolo} - Estatístico.xlsx")