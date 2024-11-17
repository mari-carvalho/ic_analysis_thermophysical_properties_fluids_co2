import pandas as pd
import os
import glob

caminhos_arquivos_excel = [
    r"C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Seleção Treino_Teste\arquivo_convertido_0.xlsx"
    r"C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Seleção Treino_Teste\arquivo_convertido_1.xlsx"
    r"C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Seleção Treino_Teste\arquivo_convertido_2.xlsx"
]

arquivo_saida = r"C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Seleção Treino_Teste\arquivo_completo.xlsx"

arquivos_excel = []

for caminho in caminhos_arquivos_excel:
    arquivos_excel_excel.extend(glob.glob(caminho))

with pd.ExcelWriter(arquivo_saida, engine='openpyxl'):
    for arquivo in arquivos_excel:
        nome_aba = arquivo.split("\\")[-1].replace(".xlsx", "")

        df = pd.read_excel(arquivo)

        df.to_excel(writer, sheet_name=nome_aba, index=False)

