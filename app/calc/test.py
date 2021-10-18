## TESTE A APLICAÇÃO AQUI
from openpyxl import Workbook
import datetime 
from datetime import date

# Setup inicial
# __init__
arquivo_excel = Workbook()
planilha1 = arquivo_excel.active

# Colunas
planilha1['A2'] = 'Item'
planilha1['B2'] = 'Descrição'
planilha1['C2'] = 'Dado'
planilha1['A1'] = 'Registro:'
planilha1['B1'] = datetime.datetime.now()

area_do_terreno = 1000
area_anteriormente_construido = 0
area_computavel_subsolo = 0
area_nao_computavel_subsolo = 0
area_computavel_terreo = 500
area_nao_computavel_terreo = 0
area_computavel_superior = 0
area_nao_computavel_superior = 0
area_nao_computavel_a_construir_atico = 0
area_livre = 0
numero_de_pavimento = 0
altura_total = 0
area_de_cobertura = 0

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
    'area_anteriormente_construido' : area_anteriormente_construido,
    'area_computavel_subsolo': area_computavel_subsolo,
    'area_nao_computavel_subsolo': area_nao_computavel_subsolo,
    'area_computavel_terreo': area_computavel_terreo,
    'area_nao_computavel_terreo': area_nao_computavel_terreo,
    'area_computavel_superior': area_computavel_superior,
    'area_nao_computavel_superior': area_nao_computavel_superior,
    'area_nao_computavel_a_construir_atico': area_nao_computavel_a_construir_atico,
    'area_livre': area_livre,
    'numero_de_pavimento': numero_de_pavimento,
    'altura_total': altura_total,
    'projecao_edificacao': projecao_edificacao,
    'taxa_ocupacao': taxa_ocupacao,
    'taxa_permeabilidade': taxa_permeabilidade,
    'area_total_contruir_subsolo': area_total_contruir_subsolo,
    'area_total_contruir_terreo': area_total_contruir_terreo,
    'area_total_contruir_superior': area_total_contruir_superior,
    'area_total_contruir_computavel': area_total_contruir_computavel,
    'area_total_contruir_nao_computavel': area_total_contruir_nao_computavel,
    'coeficiente_aproveitamento': coeficiente_aproveitamento,
    'area_de_cobertura': area_de_cobertura,
    'area_total_construida_liberada': area_total_construida_liberada
}

for item, descricao in enumerate(dicionario_com_resultado):
    linha = (item+1, descricao, dicionario_com_resultado[descricao])
    planilha1.append(linha)

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
numero_do_protocolo = '1234/2021'
interessado_projeto = 'Aole'
planilha1['C1'] = f'Protocolo {numero_do_protocolo}'
planilha1['D1'] = f'Interessado {interessado_projeto}'


data_atual = date.today()
print(data_atual)
arquivo_excel.save(fr"E:\python\app-pyqt5\app\calc\relatorios\relatorio-{data_atual}.xlsx")
