import xlsxwriter as xw

folder=r'D:\coding'
file=folder+'\XlsxWriter_cell_table.xlsx'
wb=xw.Workbook(file)
ws=wb.add_worksheet()

datas=[
    ['국방부',65,70,75],
    ['기재부',55,90,40],
    ['농식품부',80,30,60],
    ['산업부',90,20,30],
    ['과기부',40,50,60]
    ]
formula='=SUM(marklist[@[21년]:[24년]])'

ws.add_table("A2:E7",{'data':datas,
             'autofilter':False,
             'name':'marklist',         
             'columns':[
                {'header':'21년'},
                {'header':'22년'},
                {'header':'23년'},
                {'header':'24년'},
                {'header':'총합','formula':formula},
             ]})
alpha2=['B','C','D','E']
for i in range(4):
    ws.write(0,i+1,f'=SUBTOTAL(9,{alpha2[i]}2:{alpha2[i]}7)')
wb.close()