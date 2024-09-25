import pandas as pd
from openpyxl import load_workbook

file_path = r'C:\Users\03950025081\Desktop\Metan Databank.csv'

df = pd.read_csv(file_path, delimiter=';', encoding='utf-8', dtype=str)

def corrigir_pontos(valor):
    if '.' in valor:
        partes = valor.split('.')
        if len(partes[0]) > 3:
            valor_corrigido = float(valor.replace('.', '')) / 1000
            return f"{valor_corrigido: .2f}".replace('.', ',')
        else:
            return valor.replace('.', ',')
    return valor

df = df.applymap(corrigir_pontos)

caminho_novo_csv = r'C:\Users\03950025081\Desktop\Metan Databank Modificado Modificado.csv'
df.to_csv(caminho_novo_csv, index=False, sep=';', encoding='utf-8')



