import xlsxwriter as xw

folder=r'D:\coding'
file=folder+'\XlsxWriter_cell_format_07.xlsx'
wb=xw.Workbook(file)
ws=wb.add_worksheet()

names=['프로그램','21년','22년','23년']
ws.write_row('A1',names)

data=[['장병 보건 및 복지향상','국방정보화','군인사 및 교육훈련',
        '국방행정지원','정책기획 및 협력'],
      [4991,6424,8024,8132,11912],
      [7981,7329,9069,7474,13994],
      [12212,7347,9610,6887,12773]]
ws.write_column('A2',data[0])
ws.write_column('B2',data[1])
ws.write_column('C2',data[2])
ws.write_column('D2',data[3])

chart=wb.add_chart({'type':'column'})
chart.add_series({'name':'=Sheet1!$B$1','values':'=Sheet1!$B$2:$B$6','categories':'=Sheet1!$A$2:$A$6'})
chart.add_series({'name':'=Sheet1!$C$1','values':'=Sheet1!$C$2:$C$6','categories':'=Sheet1!$A$2:$A$6'})
chart.add_series({'name':'=Sheet1!$D$1','values':'=Sheet1!$D$2:$D$6','categories':'=Sheet1!$A$2:$A$6','data_labels':{'value':True}})

chart.set_legend({'position': 'right'})
# chart.set_legend({'position': 'top'})
chart.set_title({'name': '국방부 주요예산'})
chart.set_x_axis({'name': '프로그램','name_font':{'size':14,'bold':True}})
chart.set_y_axis({'name': '예산액','name_font':{'size':14,'bold':True},'min':4000,'max':14000})
# chart.set_table({'show_keys': True})

ws.insert_chart('E1',chart,{'x_scale':1.5,'y_scale':1})
wb.close()