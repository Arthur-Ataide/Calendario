import win32com.client
from gerador import ANO

# Caminho do arquivo Excel
excel_file = "Agenda_"+str(ANO)+".xlsx"
pdf_file = "Agenda_"+str(ANO)+".pdf"

# Cria um objeto Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

# Abre o arquivo
workbook = excel.Workbooks.Open(excel_file)

# Exporta para PDF
workbook.ExportAsFixedFormat(0, pdf_file)  # 0 é para PDF

# Fecha o Excel
workbook.Close()
excel.Quit()

print(f"Arquivo PDF criado: {pdf_file}")
