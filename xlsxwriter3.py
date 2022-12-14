import xlsxwriter as xw
import pandas as pd

folder=r'D:\coding'
file=folder+'\XlsxWriter_cell_format_news_03.xlsx'
wb=xw.Workbook(file)
ws=wb.add_worksheet()

#너비 설정(0:A열, 1: B열 등)
ws.set_column(0,0,21)
ws.set_column(1,3,11)

title=['프로그램','22 예산','23 예산(안)','증감','%']
budget=['국방정보화','장병 보건 및 복지향상','군수지원 및 협력','군인사 및 교육훈련',
        '군사시설 건설 및 운영','복지기금 전출금','예비전력관리','책임운영기관 운영','정책기획 및 협력',
        '국방행정지원']
budget22=[7329,7981,60219,9069,50318,200,2612,2157,13994,7474] 
budget23=[7347,12212,65132,9610,49471,0,2616,2360,12773,6887]
differ=[18,4231,4913,541,-847,-200,4,202,-1222,-586]

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
cell_format2.set_num_format(3)

cell_format3=wb.add_format()
cell_format3.set_border(1)
cell_format3.set_border_color('green')
cell_format3.set_num_format('0.0%')

ws.write(0,0,"전력운영비",cell_format1)
ws.write(0,1,"=sum(b3:b13)",cell_format2)
ws.write(0,2,"=sum(c3:c13)",cell_format2)
ws.write(0,3,"=sum(d3:d13)",cell_format2)
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