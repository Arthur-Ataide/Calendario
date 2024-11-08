import calendar
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.worksheet.page import PageMargins

# Função para gerar a lista de datas de 2025
def generate_dates(year=2025):
    dias_semana = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
    meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

    dates = []
    for month in range(1, 13):
        num_days = calendar.monthrange(year, month)[1]
        for day in range(1, num_days + 1):
            weekday = dias_semana[calendar.weekday(year, month, day)]
            month_name = meses[month - 1]
            date = f"{day} | {weekday} | {month_name} | {year}"
            dates.append(date)
    return dates

# Função para configurar a página de cada dia
def setup_day_page(ws, day_text):
    # Alinhamento centralizado e fonte em negrito
    align_center = Alignment(horizontal='center', vertical='center')
    bold_font = Font(size=14, bold=True)

    # Estilos de borda (grades)
    thick_border = Border(
        left=Side(style='thick'), right=Side(style='thick'),
        top=Side(style='thick'), bottom=Side(style='thick')
    )

    # Cabeçalho (Dia | Semana | Mês | Ano)
    ws.merge_cells('A1:H2')  # Mescla células A1 até H2
    header_cell = ws['A1']
    header_cell.value = day_text
    header_cell.alignment = align_center
    header_cell.font = bold_font
    header_cell.border = thick_border

    # Divisões de escrita (com bordas visíveis)
    for row in range(4, 40, 9):  # Cria 4 divisões de escrita
        ws.merge_cells(f'A{row}:H{row+8}')
        for r in range(row, row+9):
            for col in range(1, 9):
                cell = ws.cell(row=r, column=col)
                cell.border = thick_border

# Função principal para criar a agenda
def create_agenda():
    wb = Workbook()
    dates = generate_dates()

    for i, date in enumerate(dates):
        ws = wb.create_sheet(title=f"Dia {i+1}")
        setup_day_page(ws, date)
        ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
        ws.print_area = 'A1:H39'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1

    wb.remove(wb['Sheet'])  # Remove a planilha padrão
    wb.save('Agenda_2025.xlsx')

# Executa a função principal
create_agenda()