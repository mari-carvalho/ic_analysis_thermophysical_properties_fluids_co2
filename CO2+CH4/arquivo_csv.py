import numpy as np
import pandas as pd
import openpyxl
import csv

file_path = r'C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Metan Databank NOVO.xlsx'

workbook = openpyxl.load_workbook(file_path, data_only=True) # carregar a planilha com data_only=True para ler os resultados das fórmulas

sheet = workbook.worksheets[2] # selecionar apenas a primeira aba

caminho_novo = r'C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Metan Databank.csv' # caminho onde será salvo o novo arquivo

with open(caminho_novo, mode='w', newline='', encoding='utf-8') as file_csv: # abre o arquivo para escrita ou para ser sobrescrito (se já existir); newline evita a inserção de linhas em branco; encoding='utf-8' define a codificação do arquivo para suportar caracteres especiais; as file_csv garante que será aberto e fechado ao final do bloco
    writer = csv.writer(file_csv, delimiter=';') # cria um objeto writer usado para escrever no arquivo csv

    for row in sheet.iter_rows(values_only=True): # itera sobre as linhas da planilha, retornando valores das células; values_only=True garante que apenas valores sejam retornados (em vez de objetos)
        writer.writerow(row) # escreve a linha atual no arquivo csv, que é lida como uma tupla (escreve uma nova linha no arquivo csv, com os valores da linha separados por vírgulas)