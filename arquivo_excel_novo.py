import pandas as pd

file_path = r'C:\Users\03950025081\Desktop\Metan Databank Modificado Moificado.csv'

df = pd.read_csv(file_path, delimiter=';')

df.to_excel(r'C:\Users\03950025081\Desktop\Metan Databank NOVO.xlsx', index=False)