from openpyxl import Workbook

wb = Workbook()

planilha = wb.worksheets[0]

planilha['A1'] = "banana"
planilha['B1'] = "paçoca"

planilha.title = "Planilha de comida"
wb.save("/home/vitor/projetos/int-py-xl/meuarquivo.xlsx")