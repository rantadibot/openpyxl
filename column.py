import xlsxwriter as xw

folder=r'D:\coding'
file=folder+'\XlsxWriter_cell_format_06.xlsx'
wb=xw.Workbook(file)
ws=wb.add_worksheet()

data=[30,20,40,60,50,80,70,10]
ws.write_column('A1',data)

chart=wb.add_chart({'type':'column'})
chart.add_series({'values':'=Sheet1!$A$1:$A$8',
                  'gap':2,
                  'data_labels':{'value':True},
                  'marker':{
                     'type':'circle',
                     'size':6,
                     'border':{'color':'black'},
                     'fill':{'color':'blue'}
                     }})

chart.set_legend({'position': 'none'})
# chart.set_legend({'position': 'top'})
chart.set_title({'name': '연습용'})
chart.set_x_axis({'name': 'index'})
chart.set_y_axis({'name': 'value'})

ws.insert_chart('C1',chart,{'x_scale':1,'y_scale':1.2})
wb.close()