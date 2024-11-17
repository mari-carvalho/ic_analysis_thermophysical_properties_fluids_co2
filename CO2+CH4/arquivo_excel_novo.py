import pandas as pd

file_path = r'C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Metan Databank NOVO.csv'

df = pd.read_csv(file_path, delimiter=';')

df.to_excel(r'C:\Users\03950025081\Desktop\Simulações Thermobuilder\CH4+CO2\Metan Databank NOVO2.xlsx', index=False)