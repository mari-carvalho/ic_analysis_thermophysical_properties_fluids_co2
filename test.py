import openpyxl

# Caminho para o arquivo Excel
caminho_arquivo = 'C:\\Users\\03950025081\\Desktop\\Simulações Thermobuilder\\CH4+CO2\\Metan Databank - Copia.xlsx'

# Carregar o arquivo Excel
workbook = openpyxl.load_workbook(caminho_arquivo)

# Selecionar a primeira planilha
sheet = workbook.active

# Iterar pelas células de uma coluna específica (por exemplo, coluna A)
for row in sheet.iter_rows(min_row=2):  # Começar a partir da linha 2 para ignorar cabeçalhos
    for cell in row:
        if isinstance(cell.value, str) and '.' in cell.value:  # Verifica se o valor é uma string e contém um ponto
            # Trocar ponto por vírgula
            cell.value = cell.value.replace('.', ',')

# Salvar o arquivo Excel
workbook.save(caminho_arquivo)

print("Substituição concluída!")
