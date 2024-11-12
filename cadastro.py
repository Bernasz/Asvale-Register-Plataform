import pandas as pd
from openpyxl import load_workbook
import os
import subprocess

# Nome do arquivo Excel
file_name = 'cadastro_entregas.xlsx'

def abrir_excel(file_name):
    """Abre o arquivo Excel automaticamente após salvar ou finalizar a pesquisa"""
    if os.name == 'posix':  # Para sistemas Linux/Mac
        subprocess.call(['open', file_name])
    elif os.name == 'nt':  # Para sistemas Windows
        os.startfile(file_name)

# Verifica se o arquivo já existe
if os.path.exists(file_name):
    # Carrega o arquivo existente, especificando que colunas com zeros à esquerda devem ser strings
    df = pd.read_excel(file_name, dtype={'CPF/CNPJ': str, 'Fone': str, 'Data': str, 'I.E/Produtor': str})
else:
    # Cria um novo DataFrame se o arquivo não existir
    df = pd.DataFrame(columns=[ 
        'Produtor/Empresa', 'CPF/CNPJ', 'I.E/Produtor', 'Endereço', 'Município', 'Revenda(s)', 'UF', 'Fone', 'Data',
        'Embalagens Lavadas', 'Embalagens Lavadas 1L', 'Embalagens Lavadas 5L', 'Embalagens Lavadas 10L', 'Embalagens Lavadas 20L',
        'Embalagens Lavadas ML', 'Peso Lavadas ML', 'Embalagens Lavadas GR', 'Peso Lavadas GR', 'Embalagens Lavadas KG', 'Peso Lavadas KG',
        'Embalagens Não Lavadas', 'Embalagens Não Lavadas 1L', 'Embalagens Não Lavadas 5L', 'Embalagens Não Lavadas 10L', 'Embalagens Não Lavadas 20L',
        'Embalagens Não Lavadas ML', 'Peso Não Lavadas ML', 'Embalagens Não Lavadas GR', 'Peso Não Lavadas GR', 'Embalagens Não Lavadas KG', 'Peso Não Lavadas KG',
        'Embalagens Impróprias', 'Embalagens Impróprias 1L', 'Embalagens Impróprias 5L', 'Embalagens Impróprias 10L', 'Embalagens Impróprias 20L',
        'Embalagens Impróprias ML', 'Peso Impróprias ML', 'Embalagens Impróprias GR', 'Peso Impróprias GR', 'Embalagens Impróprias KG', 'Peso Impróprias KG'
    ])

while True:
    # Captura as informações do cadastro
    produtor = input("Produtor/Empresa: ")
    cpf_cnpj = input("CPF/CNPJ: ")
    # Verifica se o valor inserido é CPF ou CNPJ e preenche com zeros à esquerda se necessário
    if len(cpf_cnpj) == 11:
        cpf_cnpj = cpf_cnpj.zfill(11)  # CPF com 11 dígitos
    elif len(cpf_cnpj) == 14:
        cpf_cnpj = cpf_cnpj.zfill(14)  # CNPJ com 14 dígitos
    ie_produtor = input("I.E/Produtor: ")
    endereco = input("Endereço: ")
    municipio = input("Município: ")
    revendas = input("Revenda(s): ")
    uf = input("UF: ")
    fone = input("Fone: ").zfill(11)  # Formata telefone para manter 11 dígitos (exemplo)
    data = input("Data (dd/mm/yyyy): ")

    # Pergunta sobre as embalagens e armazena como "Sim" ou "Não"
    embalagens_lavadas = input("Teve entrega de embalagens lavadas? (s/n): ").lower()
    embalagens_lavadas = "Sim" if embalagens_lavadas == 's' else "Não"

    embalagens_nao_lavadas = input("Teve entrega de embalagens não lavadas? (s/n): ").lower()
    embalagens_nao_lavadas = "Sim" if embalagens_nao_lavadas == 's' else "Não"

    embalagens_improprias = input("Teve entrega de embalagens impróprias? (s/n): ").lower()
    embalagens_improprias = "Sim" if embalagens_improprias == 's' else "Não"

    # Inicializa as variáveis para as embalagens
    embalagens_lavadas_1L = embalagens_lavadas_5L = embalagens_lavadas_10L = embalagens_lavadas_20L = 0
    embalagens_lavadas_ml = embalagens_lavadas_gr = embalagens_lavadas_kg = 0
    peso_lavadas_ml = peso_lavadas_gr = peso_lavadas_kg = 0

    embalagens_nao_lavadas_1L = embalagens_nao_lavadas_5L = embalagens_nao_lavadas_10L = embalagens_nao_lavadas_20L = 0
    embalagens_nao_lavadas_ml = embalagens_nao_lavadas_gr = embalagens_nao_lavadas_kg = 0
    peso_nao_lavadas_ml = peso_nao_lavadas_gr = peso_nao_lavadas_kg = 0

    embalagens_improprias_1L = embalagens_improprias_5L = embalagens_improprias_10L = embalagens_improprias_20L = 0
    embalagens_improprias_ml = embalagens_improprias_gr = embalagens_improprias_kg = 0
    peso_improprias_ml = peso_improprias_gr = peso_improprias_kg = 0

    # Pergunta sobre as embalagens para cada tipo
    if embalagens_lavadas == "Sim":
        embalagens_lavadas_1L = int(input("Quantas embalagens de 1L (Lavadas)? "))
        embalagens_lavadas_5L = int(input("Quantas embalagens de 5L (Lavadas)? "))
        embalagens_lavadas_10L = int(input("Quantas embalagens de 10L (Lavadas)? "))
        embalagens_lavadas_20L = int(input("Quantas embalagens de 20L (Lavadas)? "))

        embalagens_lavadas_ml = int(input("Quantas embalagens de ML (Lavadas)? "))
        if embalagens_lavadas_ml > 0:
            peso_lavadas_ml = '{:.2f}'.format(float(input("Qual o volume de cada embalagem em ML (Lavadas)? ")))

        embalagens_lavadas_gr = int(input("Quantas embalagens de GR (Lavadas)? "))
        if embalagens_lavadas_gr > 0:
            peso_lavadas_gr = '{:.2f}'.format(float(input("Qual o peso de cada embalagem em GR (Lavadas)? ")))

        embalagens_lavadas_kg = int(input("Quantas embalagens de KG (Lavadas)? "))
        if embalagens_lavadas_kg > 0:
            peso_lavadas_kg = '{:.2f}'.format(float(input("Qual o peso de cada embalagem em KG (Lavadas)? ")))

    if embalagens_nao_lavadas == "Sim":
        embalagens_nao_lavadas_1L = int(input("Quantas embalagens de 1L (Não Lavadas)? "))
        embalagens_nao_lavadas_5L = int(input("Quantas embalagens de 5L (Não Lavadas)? "))
        embalagens_nao_lavadas_10L = int(input("Quantas embalagens de 10L (Não Lavadas)? "))
        embalagens_nao_lavadas_20L = int(input("Quantas embalagens de 20L (Não Lavadas)? "))

        embalagens_nao_lavadas_ml = int(input("Quantas embalagens de ML (Não Lavadas)? "))
        if embalagens_nao_lavadas_ml > 0:
            peso_nao_lavadas_ml = '{:.2f}'.format(float(input("Qual o volume de cada embalagem em ML (Não Lavadas)? ")))

        embalagens_nao_lavadas_gr = int(input("Quantas embalagens de GR (Não Lavadas)? "))
        if embalagens_nao_lavadas_gr > 0:
            peso_nao_lavadas_gr = '{:.2f}'.format(float(input("Qual o peso de cada embalagem em GR (Não Lavadas)? ")))

        embalagens_nao_lavadas_kg = int(input("Quantas embalagens de KG (Não Lavadas)? "))
        if embalagens_nao_lavadas_kg > 0:
            peso_nao_lavadas_kg = '{:.2f}'.format(float(input("Qual o peso de cada embalagem em KG (Não Lavadas)? ")))

    if embalagens_improprias == "Sim":
        embalagens_improprias_1L = int(input("Quantas embalagens de 1L (Impróprias)? "))
        embalagens_improprias_5L = int(input("Quantas embalagens de 5L (Impróprias)? "))
        embalagens_improprias_10L = int(input("Quantas embalagens de 10L (Impróprias)? "))
        embalagens_improprias_20L = int(input("Quantas embalagens de 20L (Impróprias)? "))

        embalagens_improprias_ml = int(input("Quantas embalagens de ML (Impróprias)? "))
        if embalagens_improprias_ml > 0:
            peso_improprias_ml = '{:.2f}'.format(float(input("Qual o volume de cada embalagem em ML (Impróprias)? ")))

        embalagens_improprias_gr = int(input("Quantas embalagens de GR (Impróprias)? "))
        if embalagens_improprias_gr > 0:
            peso_improprias_gr = '{:.2f}'.format(float(input("Qual o peso de cada embalagem em GR (Impróprias)? ")))

        embalagens_improprias_kg = int(input("Quantas embalagens de KG (Impróprias)? "))
        if embalagens_improprias_kg > 0:
            peso_improprias_kg = '{:.2f}'.format(float(input("Qual o peso de cada embalagem em KG (Impróprias)? ")))

    dados = {
        'Produtor/Empresa': produtor,
        'CPF/CNPJ': cpf_cnpj,
        'I.E/Produtor': ie_produtor,
        'Endereço': endereco,
        'Município': municipio,
        'Revenda(s)': revendas,
        'UF': uf,
        'Fone': fone,
        'Data': data,
        'Embalagens Lavadas': embalagens_lavadas,
        'Embalagens Lavadas 1L': embalagens_lavadas_1L,
        'Embalagens Lavadas 5L': embalagens_lavadas_5L,
        'Embalagens Lavadas 10L': embalagens_lavadas_10L,
        'Embalagens Lavadas 20L': embalagens_lavadas_20L,
        'Embalagens Lavadas ML': embalagens_lavadas_ml,
        'Peso Lavadas ML': peso_lavadas_ml,
        'Embalagens Lavadas GR': embalagens_lavadas_gr,
        'Peso Lavadas GR': peso_lavadas_gr,
        'Embalagens Lavadas KG': embalagens_lavadas_kg,
        'Peso Lavadas KG': peso_lavadas_kg,
        'Embalagens Não Lavadas': embalagens_nao_lavadas,
        'Embalagens Não Lavadas 1L': embalagens_nao_lavadas_1L,
        'Embalagens Não Lavadas 5L': embalagens_nao_lavadas_5L,
        'Embalagens Não Lavadas 10L': embalagens_nao_lavadas_10L,
        'Embalagens Não Lavadas 20L': embalagens_nao_lavadas_20L,
        'Embalagens Não Lavadas ML': embalagens_nao_lavadas_ml,
        'Peso Não Lavadas ML': peso_nao_lavadas_ml,
        'Embalagens Não Lavadas GR': embalagens_nao_lavadas_gr,
        'Peso Não Lavadas GR': peso_nao_lavadas_gr,
        'Embalagens Não Lavadas KG': embalagens_nao_lavadas_kg,
        'Peso Não Lavadas KG': peso_nao_lavadas_kg,
        'Embalagens Impróprias': embalagens_improprias,
        'Embalagens Impróprias 1L': embalagens_improprias_1L,
        'Embalagens Impróprias 5L': embalagens_improprias_5L,
        'Embalagens Impróprias 10L': embalagens_improprias_10L,
        'Embalagens Impróprias 20L': embalagens_improprias_20L,
        'Embalagens Impróprias ML': embalagens_improprias_ml,
        'Peso Impróprias ML': peso_improprias_ml,
        'Embalagens Impróprias GR': embalagens_improprias_gr,
        'Peso Impróprias GR': peso_improprias_gr,
        'Embalagens Impróprias KG': embalagens_improprias_kg,
        'Peso Impróprias KG': peso_improprias_kg
    }
    
    # Ordena o DataFrame pela coluna 'Produtor/Empresa'
    df = df.sort_values(by='Produtor/Empresa')

    # Adiciona os dados ao DataFrame
    df = pd.concat([df, pd.DataFrame([dados])], ignore_index=True)

    # Salva o DataFrame no arquivo Excel
    df.to_excel(file_name, index=False)

    # Pergunta se o usuário deseja adicionar mais dados
    mais_dados = input("Deseja adicionar mais dados? (s/n): ").lower()
    if mais_dados != 's':
        break

# Abre o arquivo Excel automaticamente
abrir_excel(file_name)
