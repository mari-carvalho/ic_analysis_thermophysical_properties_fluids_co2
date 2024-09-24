import pandas as pd

file_path = r'C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Metan Databank.csv'

df = pd.read_csv(file_path, delimiter=';', encoding='utf-8')

for column in df.select_dtypes(include=['float64', 'int64']).columns:
    df[column] = df[column].astype(str).str.replace('.', ',', regex=False)

caminho_novo_csv = r'C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Metan Databank Modificado.csv'
df.to_csv(caminho_novo_csv, index=False, sep=';', encoding='utf-8')



