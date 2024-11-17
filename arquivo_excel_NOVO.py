import pandas as pd
import os

arquivo_csv = r"C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Seleção Treino_Teste\inliers_dataset_sheet_2.csv"

arquivo_excel = r"C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Seleção Treino_Teste\arquivo_convertido_2.xlsx"

df = pd.read_csv(arquivo_csv)

df.to_excel(arquivo_excel, index=False)

