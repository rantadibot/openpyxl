import xlsxwriter as xw
import pandas as pd

folder=r'C:\Users\xkr04\OneDrive\바탕 화면\코딩\pyexcel-master\pyexcel-master\data\ch06'
file=folder+'\XlsxWriter_cell_format_news_01.xlsx'
wb=xw.Workbook(file)
ws=wb.add_worksheet()
title=['프로그램','22 예산','23 예산(안)','증감','%']
budget=['국방정보화','장병 보건 및 복지향상','군수지원 및 협력','군인사 및 교육훈련',
        '군사시설 건설 및 운영','복지기금 전출금','예비전력관리','책임운영기관 운영','정책기획 및 협력',
        '국방행정지원']
budget22=[7329,7981,60219,9069,50318,200,2612,2157,13994,7474] #합 : 45686
budget23=[7347,12212,65132,9610,49471,0,2616,2360,12773,6887]
differ=[18,4231,4913,541,-847,-200,4,202,-1222,-586]
cell_format1=wb.add_format({'font_name':'맑은 고딕',
                            'bold':True,
                            'font_color':'navy'})
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

for i in range(len(title)):
    ws.write(1,i+1,title[i],cell_format1)

for j in range(len(budget)):
    ws.write(j+2,1,budget[j],cell_format1)
    ws.write(j+2,2,budget22[j],cell_format2)
    ws.write(j+2,3,budget23[j],cell_format2)
    ws.write(j+2,4,differ[j],cell_format2)
    ws.write(j+2,5,differ[j]/budget22[j],cell_format3)
wb.close()