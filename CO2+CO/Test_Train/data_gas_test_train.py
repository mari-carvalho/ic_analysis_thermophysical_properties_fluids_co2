import pandas as pd
from sklearn.model_selection import train_test_split
import os

# Ler a planilha com várias abas:
file_path = r"C:\Users\03950025081\Desktop\Simulações Thermobuilder\CO+CO2\Seleção Treino_Teste\densidades_gásl_líquido_treino_teste.xlsx"
sheet_names = pd.ExcelFile(file_path).sheet_names # Obtém os nomes das abas

# Lista para armazenar DataFrames:
dfs = []

for sheet in sheet_names:
    df = pd.read_excel(file_path, sheet_name=1)
    dfs.append(df)

# Preparar os dados:
data = pd.concat(dfs) # combina os dados de todas as abas

print(data.columns)
print(data)

# Selecionar as colunas de interesse, incluindo a densidade como feature:
features = data[['CO2', 'CO', 'TEMPERATURE', 'PRESSURE', 'EXPE. DENSITY', 'ID']]

# Dividir os dados em treino e teste:
X_train, X_test = train_test_split(features, test_size=0.25, random_state=42) # random_state=42 garante que a separação dos dados será sempre a mesma na execução do código

# Definir o caminho da pasta onde deseja salvar os arquivos:
folder_path = r"C:\Users\03950025081\Desktop\Simulações Thermobuilder\CO+CO2\Seleção Treino_Teste\Train_Test Gas"

# Criar o caminho completo para os arquivos:
train_file_path = os.path.join(folder_path, 'dados_treinamento_com_densidade_gas.xlsx')
test_file_path = os.path.join(folder_path, 'dados_teste_com_densidade_gas.xlsx')

# Salvar os dados em novas planilhas:
# Dados de Treinamento:
X_train.to_excel(train_file_path, index=False) #index=false impede que os índices das linhas sejam salvos no arquivo Excel

# Dados de Teste:
X_test.to_excel(test_file_path, index=False)




