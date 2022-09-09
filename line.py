import xlsxwriter as xw

folder=r'D:\coding'
file=folder+'\XlsxWriter_cell_format_05.xlsx'
wb=xw.Workbook(file)
ws=wb.add_worksheet()

data=[44.5,75.5,103.9,131.1,158.7,185.7]
ws.write_column('A1',data)

chart=wb.add_chart({'type':'line'})
chart.add_series({'values':'=Sheet1!$A$1:$A$6',
                  'data_labels':{'value':True},
                  'marker':{
                     'type':'circle',
                     'size':6,
                     'border':{'color':'black'},
                     'fill':{'color':'blue'}
                     }})
chart.set_title({'name': '대한민국 국가채무비율'})
chart.set_x_axis({'name': 'x_data'})
chart.set_y_axis({'name': 'y_data'})

ws.insert_chart('C1',chart,{'x_scale':1,'y_scale':1.2})
wb.close()