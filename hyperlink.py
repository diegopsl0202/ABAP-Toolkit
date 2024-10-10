import win32com.client

# Conectar ao Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Manter o Excel invisível

# Abrir a planilha existente
workbook = excel.Workbooks.Open(r"C:\caminho\para\arquivo.xlsx")
sheet = workbook.Worksheets("NomeDaAba")  # Aba onde o hiperlink será inserido

# Adicionar o hiperlink para a célula A3 da aba "Cycle 1"
cell = sheet.Cells(1, 1)  # Define a célula A1 onde o link será inserido
sub_address = "'Cycle 1'!A3"  # SubAddress para a célula A3 da aba "Cycle 1"
display_text = "Ir para Cycle 1 - A3"

sheet.Hyperlinks.Add(Anchor=cell, Address="", SubAddress=sub_address, TextToDisplay=display_text)

# Salvar e fechar o arquivo
workbook.Save()
workbook.Close()
excel.Quit()
