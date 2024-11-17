import pandas as pd
import os

xls = r"C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Seleção Treino_Teste\Metan Databank NOVO.xlsx"

df = pd.read_excel(xls, sheet_name=0)

arquivo_csv = r"C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Seleção Treino_Teste\arquivo_csv_final.csv"
df.to_csv(arquivo_csv, index=False, sep=';')