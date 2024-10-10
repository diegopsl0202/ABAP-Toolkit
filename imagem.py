import win32com.client

# Conectar ao Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Manter o Excel invisível

# Abrir a planilha existente
workbook = excel.Workbooks.Open(r"C:\Users\diego.augusto\OneDrive - Accenture\Unilever_backup\Release\Team\teste docs\CTP_SAP ECC_V1.0_New Tasa Protección Ambiental Pilar Argentina 2 - Copy.xls")

# Variável para armazenar a aba desejada
first_cycle_sheet_without_images = None

# Percorrer todas as planilhas no livro
for sheet in workbook.Worksheets:
    # Verificar se o nome da aba contém "Cycle"
    if "Cycle" in sheet.Name:
        # Verificar se a aba não contém imagens
        if sheet.Shapes.Count == 0:  # Se não há formas (imagens, gráficos, etc.)
            first_cycle_sheet_without_images = sheet.Name  # Armazena a primeira aba que contém "Cycle" e não tem imagens
            break  # Sai do loop se encontrar a primeira aba com "Cycle" e sem imagens

# Decidir qual aba retornar
if first_cycle_sheet_without_images:
    print(f"A primeira aba que contém 'Cycle' e não tem imagens: '{first_cycle_sheet_without_images}'.")
else:
    print("Nenhuma aba encontrada que contenha 'Cycle' sem imagens.")

# Fechar o arquivo e sair do Excel
workbook.Close(SaveChanges=False)
excel.Quit()
