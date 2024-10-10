import os
import win32com.client

def convert_xls_to_xlsx_with_macros(xls_file_path, xlsx_file_path):
    # Verifica se o arquivo .xls existe
    if not os.path.isfile(xls_file_path):
        print(f"O arquivo {xls_file_path} não existe.")
        return
    
    # Cria uma instância do Excel
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Mantenha o Excel oculto durante a execução

    try:
        # Abre o arquivo .xls
        workbook = excel.Workbooks.Open(xls_file_path)
        
        # Salva o arquivo diretamente como .xlsx
        workbook.SaveAs(xlsx_file_path, FileFormat=51)  # 51 é o código para .xlsx
        
        # Aqui aplicamos a opção "Save and erase features"
        workbook.Close(SaveChanges=True)  # Fecha o workbook salvando as mudanças
        
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        # Fecha o Excel
        excel.Quit()
    
    print(f"Arquivo convertido: {xlsx_file_path}")

# Caminho do arquivo .xls
xls_file = r"C:\Users\diego.augusto\OneDrive - Accenture\Unilever_backup\Release\Team\teste docs\CTP_SAP ECC_V1.0_New Tasa Protección Ambiental Pilar Argentina 2 - Copy.xls"
# Caminho para salvar o arquivo .xlsx
xlsx_file = r"C:\Users\diego.augusto\OneDrive - Accenture\Unilever_backup\Release\Team\teste docs\CTP_SAP ECC_V1.0_New Tasa Protección Ambiental Pilar Argentina 2 - Copy.xlsx"

# Converte o arquivo
convert_xls_to_xlsx_with_macros(xls_file, xlsx_file)
