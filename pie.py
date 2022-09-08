import xlsxwriter as xw

folder=r'D:\coding'
file=folder+'\XlsxWriter_chart_01.xlsx'
wb=xw.Workbook(file)
ws=wb.add_worksheet()

headings=['분류','점유율']
data=[['삼성','애플','샤오미'],[60,30,10]]
bold=wb.add_format({'bold':1})

ws.write_row('A1',headings,bold)
ws.write_column('A2',data[0])
ws.write_column('B2',data[1])

pi_chart=wb.add_chart({'type':'pie'})
pi_chart.add_series({
    'name':'스마트폰 점유율',
    'categories':['Sheet1',1,0,3,0],
    'values':['Sheet1',1,1,3,1]
})

pi_chart.set_title({'name':'한국 스마트폰 점유율'})
pi_chart.set_style(10)
ws.insert_chart('C2',pi_chart,{'x_offset':30,'y_offset':20})

wb.close()