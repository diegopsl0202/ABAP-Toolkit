import xlwings as xw

def convert_xls_to_xlsm(input_file, output_file):
    # Inicia uma nova aplicação Excel
    app = xw.App(visible=False)
    # Abre o arquivo .xls
    wb = app.books.open(input_file)

    # Salva como .xlsm para preservar macros
    wb.save(output_file)
    wb.close()
    app.quit()

# Exemplo de uso
input_file_path = r"C:\Users\diego.augusto\OneDrive - Accenture\Unilever_backup\Release\Team\teste docs\CTP_SAP ECC_V9.0_5th Flag Automation Cockpit - Copy.xls"
output_file_path = r"C:\Users\diego.augusto\OneDrive - Accenture\Unilever_backup\Release\Team\teste docs\CTP_SAP ECC_V9.0_5th Flag Automation Cockpit - Copy.xlsm"

convert_xls_to_xlsm(input_file_path, output_file_path)
