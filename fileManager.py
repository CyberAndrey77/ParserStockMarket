# coding: utf8
import xlsxwriter

def write_exel(rowsUSD, rowsEUR):
    workbook = xlsxwriter.Workbook('USD_EUR.xlsx')
    cellFormat = workbook.add_format()
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:G', 10)
    cellFormat.set_center_across()

    worksheet.write('A1', 'Дата', cellFormat)
    worksheet.write('B1', 'Значение', cellFormat)
    worksheet.write('C1', 'Изменение', cellFormat)
    worksheet.write('D1', 'Дата', cellFormat)
    worksheet.write('E1', 'Значение', cellFormat)
    worksheet.write('F1', 'Изменение', cellFormat)
    worksheet.write('G1', 'Отношение', cellFormat)

    rowTable = 1
    colTable = 0

    dateFormat = workbook.add_format({'num_format': 'dd.mm.yy'})
    numberFormat = workbook.add_format({'num_format': '_-* #,##0.00 ₽_-'})
    for row in rowsUSD:
        worksheet.write(rowTable, colTable, row[0], dateFormat)
        worksheet.write(rowTable, colTable + 1, row[1], numberFormat)
        rowTable += 1

    rowTable = 1

    for i in range(len(rowsUSD) - 1):
        worksheet.write(rowTable, colTable + 2, rowsUSD[i][1] - rowsUSD[i+1][1], numberFormat)
        rowTable += 1

    rowTable = 1
    colTable += 3

    for row in rowsEUR:
        worksheet.write(rowTable, colTable, row[0], dateFormat)
        worksheet.write(rowTable, colTable + 1, row[1], numberFormat)
        rowTable += 1

    rowTable = 1

    for i in range(len(rowsEUR) - 1):
        worksheet.write(rowTable, colTable + 2, rowsEUR[i][1] - rowsEUR[i+1][1], numberFormat)
        rowTable += 1

    colTable += 1

    for i in range(rowTable):
        worksheet.write_formula('G{0}'.format(i+2), '=E{0}/B{0}'.format(i+2), numberFormat)
    workbook.close()
    return rowTable + 1

def main():
    workbook = xlsxwriter.Workbook('hello.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Дата')
    worksheet.write('B1', 'Значение')
    worksheet.write('C1', 'Изменение')

    workbook.close()

if __name__ == "__main__":
    main()