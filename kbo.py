import xlrd
import pandas as pd

ad0=pd.read_html('https://sports.news.naver.com/kbaseball/record/index?category=kbo')[0]
folder=r'D:\coding'
file=folder+'\XlsxWriter_kbo.xlsx'
writer=pd.ExcelWriter(file,engine='openpyxl')
ad0.to_excel(writer,sheet_name='팀 순위',index=False)

writer.save()

wb=xlrd.open_workbook(file)
ws=wb.sheet_by_index(0)

ncol=ws.ncols
nrow=ws.nrows

def print_score(rank):
    for i in range(nrow):
        print(ws.row_values(rank)[i+1])

rank=input("순위를 입력 하세요 : ")
print_score(int(rank))