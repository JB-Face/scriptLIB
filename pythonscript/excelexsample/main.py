import xlrd
import time
from datetime import datetime
import xlwings as xw
path = r"C:\Users\57376\Documents\我的POPO\Easy美术部外派2.xls"
data = xlrd.open_workbook(path)
import xlrd
import xlwt
from xlutils.copy import copy

def write_excel_xls(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿
    print("xls格式表格写入数据成功！")


table = data.sheets()[0]
maxrows = table.nrows
maxcols = table.ncols
pool = []
for i in range(maxrows):
    for ii in range(maxcols):
        text = table.cell(i,ii)
        #print(text)
        if text.value == '日期':
            pool.append((i,ii))
print(list(pool))
# 获取 内容池 开始建立对应的名字的信息库
lib = {}
for i in pool:
    name = table.cell(i[0]-1,i[1]).value

    pllname = name.split(' ')
    lib[pllname[0]] = {}
    lib[pllname[0]]['num'] = pllname[-1]
    lib[pllname[0]]['coord'] = i

    lib[pllname[0]]['day'] = {}
    a = 1
    while table.cell(i[0]+a,i[1]).value:
        hand = str(table.cell(i[0]+a,i[1]).value)
        di = lib[pllname[0]]['day'][hand] = {'worktime':8}
        di['week'] = ['星期一','星期二','星期三','星期四','星期五','星期六','星期日',][datetime.strptime(hand, '%Y-%m-%d').weekday()]
        di['workon'] = table.cell(i[0]+a,i[1]+1).value
        di['wokoff'] = table.cell(i[0]+a,i[1]+2).value
        di['sometime'] = '19:00'
        wo = datetime.strptime(table.cell(i[0]+a,i[1]+1).value, '%H:%M')
        wf = datetime.strptime(table.cell(i[0]+a,i[1]+2).value, '%H:%M')
        worktime = wf - wo
        realworktime = int(str(worktime).split(':')[0])
        if realworktime >= 8:
            realworktime = 8
        di['worktime'] = realworktime
        overtime = str(int(str(worktime).split(':')[0]) - 8) + ':' + str(worktime).split(':')[1]
        di['overtime'] = overtime
        di['day&person'] = None
        di['reverso'] = 900
        a = a+1



print(lib)


