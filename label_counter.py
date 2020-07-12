import xlrd
import xlsxwriter

records = {}
workbook = xlrd.open_workbook('rus_Records.xlsx')
worksheet = workbook.sheet_by_index(0)
row = 1
while True:
    try:
        if records.get(worksheet.cell(row, 4).value) is None:
            if str.lower(worksheet.cell(row, 1).value) != worksheet.cell(row, 4).value:
                records[worksheet.cell(row, 4).value] = 1
        else:
            records[worksheet.cell(row, 4).value] += 1
        row += 1
    except:
        break

records = {k: v for k, v in reversed(sorted(records.items(), key=lambda item: item[1]))}

workbook = xlsxwriter.Workbook('rus_Labels.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})
bold.set_align('center')

worksheet.set_column(0, 0, 20)
worksheet.set_column(1, 1, 10)

worksheet.write('A1', 'Label', bold)
worksheet.write('B1', 'Count', bold)

row = 1
col = 0
for skill in records.items():
    worksheet.write_string(row, col,  skill[0])
    worksheet.write_string(row, col + 1, str(skill[1]))
    row += 1
workbook.close()
print('OK')
