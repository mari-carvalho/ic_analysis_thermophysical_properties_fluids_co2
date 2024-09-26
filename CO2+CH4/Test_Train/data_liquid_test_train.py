import pandas as pd
from sklearn.model_selection import train_test_split
import os

# Ler a planilha com várias abas:
file_path = r"C:\Users\03950025081\Desktop\Metan Databank NOVO.xlsx"

# Ler a segunda e a terceira aba:
data_second_sheet = pd.read_excel(file_path, sheet_name=1, engine='openpyxl')
data_third_sheet = pd.read_excel(file_path, sheet_name=2, engine='openpyxl')

# Selecionar as colunas de interesse, incluindo a densidade como feature:
features_second_sheet = data_second_sheet[['ID', 'x CH4', 'T[K]', 'P atm', 'V cc/g-mol', 'ρ [kg/m3]']]
features_third_sheet = data_third_sheet[['ID', 'co2', 'x CH4', 'T[K]', 'P atm', 'ρ [kg/m3]']]

combined_data = pd.concat([features_second_sheet, features_third_sheet])

# Dividir os dados em treino e teste:
X_train, X_test = train_test_split(combined_data, test_size=0.25, random_state=42) # random_state=42 garante que a separação dos dados será sempre a mesma na execução do código

# Definir o caminho da pasta onde deseja salvar os arquivos:
folder_path = r"C:\Users\03950025081\Desktop"

# Criar o caminho completo para os arquivos:
train_file_path = os.path.join(folder_path, 'dados_treinamento_com_densidade_liquid.xlsx')
test_file_path = os.path.join(folder_path, 'dados_teste_com_densidade_liquid.xlsx')

# Salvar os dados em novas planilhas:
# Dados de Treinamento:
X_train.to_excel(train_file_path, index=False) #index=false impede que os índices das linhas sejam salvos no arquivo Excel

# Dados de Teste:
X_test.to_excel(test_file_path, index=False)




