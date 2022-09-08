import xlsxwriter as xw

folder=r'D:\coding'
file=folder+'\XlsxWriter_chart_01.xlsx'
wb=xw.Workbook(file)
ws=wb.add_worksheet()

headings=['분류','점유율']
data=[['네이버','카카오','SKT','LG','KT'],[50,30,10,5,5]]
bold=wb.add_format({'bold':1})

ws.write_row('A1',headings,bold)
ws.write_column('A2',data[0])
ws.write_column('B2',data[1])

pi_chart=wb.add_chart({'type':'pie'})
pi_chart.add_series({
    'name':'통신사 점유율',
    'categories':['Sheet1',1,0,5,0],
    'values':['Sheet1',1,1,5,1]
})

pi_chart.set_title({'name':'한국 통신사 점유율'})
pi_chart.set_style(10)
ws.insert_chart('C2',pi_chart,{'x_offset':30,'y_offset':20})

wb.close()