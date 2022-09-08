import xlsxwriter as xw
import pandas as pd

folder=r'D:\coding'
file=folder+'\XlsxWriter_cell_format_01.xlsx'
wb=xw.Workbook(file)
ws=wb.add_worksheet()

#너비 설정(0:A열, 1: B열 등)
ws.set_column(0,0,21)
ws.set_column(1,3,11)

title=['구분','22 예산','23 예산(안)','증감','%']
budget=['보건,복지,고용','교육','문화 체육 관광','환경',
        'R&D','산업,중소기업,에너지','SOC','농림,수산,식품','국방',
        '외교,통일','공공질서,안전','일반,지방행정']
budget22=[217.7,84.2,9.1,11.9,29.8,31.3,28.0,23.7,54.6,6.0,22.3,98.1] 
budget23=[226.6,96.1,8.5,12.4,30.7,25.7,25.1,24.2,57.1,6.4,22.9,111.7]
differ=[8.9,12.0,-0.6,0.5,0.9,-5.6,-2.8,0.6,2.5,0.4,0.5,13.6]

#format 설정, add_format걸고 나머지 걸기
cell_format1=wb.add_format({'font_name':'맑은 고딕',
                            'bold':True,
                            'font_color':'navy',
                            'align':'center'})
cell_format1.set_border(1)
cell_format1.set_font_size(11)
cell_format1.set_bg_color('yellow')

cell_format2=wb.add_format()
cell_format2.set_border(1)
cell_format2.set_border_color('green')
cell_format2.set_num_format('0.0')

cell_format3=wb.add_format()
cell_format3.set_border(1)
cell_format3.set_border_color('green')
cell_format3.set_num_format('0.0%')

ws.write(0,0,"전력운영비",cell_format1)
ws.write(0,1,"=sum(b3:b14)",cell_format2)
ws.write(0,2,"=sum(c3:c14)",cell_format2)
ws.write(0,3,"=sum(d3:d14)",cell_format2)
ws.write(0,4,"=d1/b1",cell_format3)

for i in range(len(title)):
    ws.write(1,i,title[i],cell_format1)

ws.write_column(row=2,col=0,data=budget,cell_format=cell_format1)
ws.write_column(row=2,col=1,data=budget22,cell_format=cell_format2)
ws.write_column(row=2,col=2,data=budget23,cell_format=cell_format2)
ws.write_column(row=2,col=3,data=differ,cell_format=cell_format2)

for j in range(len(budget)):
    ws.write(j+2,4,differ[j]/budget22[j],cell_format3)

wb.close()