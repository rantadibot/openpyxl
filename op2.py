import openpyxl as op
wb=op.load_workbook(r'C:\Users\xkr04\Downloads\Downloads\연습.xlsx',data_only=True)
ws=wb.active
for cols in ws.iter_cols():
    for cell in cols:
        print(cell.value,end=' ')