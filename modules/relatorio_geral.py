from xls2xlsx import XLS2XLSX


def formata_geral():
    x2x = XLS2XLSX("excel files/grade-exportacao.xls")
    x2x.to_xlsx("excel files/grade-exportacao1.xlsx")

    