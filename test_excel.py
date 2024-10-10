import openpyxl

# Caminho do arquivo
caminho_arquivo = r"C:\Users\diego.augusto\OneDrive - Accenture\Unilever_backup\Release\Team\teste docs\CTP_SAP ECC_V1.0_New Tasa Protección Ambiental Pilar Argentina 2 - Copy.xlsx"

# Carregar o arquivo Excel (funciona tanto para .xlsx quanto para .xlsm)
workbook = openpyxl.load_workbook(caminho_arquivo)

# Selecionar a aba "Document History"
sheet = workbook["Document History"]

# Iterar sobre as células na coluna A
for row in sheet.iter_rows(min_row=1, max_col=2):  # Considerando até a coluna B
    cell_value = row[0].value
    
    # Debug: imprime o valor da célula e seu tipo
    print(f'Valor encontrado: "{cell_value}" | Tipo: {type(cell_value)}')  
    
    # Verifica se o valor da célula não é None e compara com "File Name" sem espaços extras
    if cell_value is not None and str(cell_value).strip().rstrip(':') == "File Name":
        row[1].value = "TESTE FEITO"  # Modifica o valor na coluna B
        print('Modificado para "TESTE FEITO" na coluna B.')  # Mensagem de confirmação
        break  # Para sair do loop após encontrar a primeira ocorrência
else:
    print('A palavra "File Name" não foi encontrada.')

# Salvar as alterações
workbook.save(caminho_arquivo)
