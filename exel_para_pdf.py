import win32com.client

# Caminho do arquivo Excel
excel_file = "Agenda_2025.xlsx"
pdf_file = "Agenda_2025.pdf"

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
