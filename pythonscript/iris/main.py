import xlrd
import time
from datetime import datetime
path = r".\\pythonscript\\iris\\testfile\\Easy美术部外派2.xls"

import xlrd
import xlwt
from xlutils.copy import copy
import tkinter 
from tkinter.filedialog import askdirectory,askopenfilename

def write_excel_xls(path, sheet_name, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            sheet.write(i, j, value[i][j])  # 像表格中写入数据（对应的行和列）
    workbook.save(path)  # 保存工作簿



def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i + rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
 
def iris(path):
    data = xlrd.open_workbook(path)

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
            fullday = 0
            if realworktime >= 8:
                realworktime = 8
                fullday = 1        
            di['worktime'] = realworktime
            overtime = str(int(str(worktime).split(':')[0]) - 8) + ':' + str(worktime).split(':')[1]
            di['overtime'] = overtime

            # 比值计算
            otime = int(overtime.split(':')[0]) + int(overtime.split(':')[1])/60
            rtimr = realworktime

            if fullday:
                di['day&person'] =round(1+1.5*(float(otime)/8.0),1)
            else:
                di['day&person'] = round(rtimr/8.0,1)
            di['reverso'] = 900
            a = a+1

    title = [["姓名","外派日期","是否周末","上班时间","下班时间"," ","上班时长","加班时间","人天","单价"]]


    for i in lib:
        # p = path.split('\\')[-1]
        # p = path.split(p)[0] + str(i)+'.xls'
        p = ".\\"+str(i)+'.xls'
        write_excel_xls(p,'main',title)
        t = lib[i]['day']
        she = []
        for k in t:
            kk = t[k]
            newar = [str(i),str(k),kk['week'],kk['workon'],kk['wokoff'],kk['sometime'],kk['worktime'],kk['overtime'],kk['day&person'],kk['reverso']]
            she.append(newar)
        write_excel_xls_append(p,she)


def main():
    top = tkinter.Tk()
    top.title('iris 保护协会')

    path = tkinter.StringVar() 
    def selectPath():
        path_ = askopenfilename(title='iris 保护协会', initialdir='/', filetypes=[('xls','*.xls'),('xlsx','*.xlsx')])
        path.set(path_)
        if path_:
            iris(path.get())
    tkinter.Label(top,text = '选择文件').grid(row = 0, column = 0)
    tkinter.Entry(top, textvariable = path).grid(row = 0, column = 1)
    tkinter.Button(top, text = "选择文件", command = selectPath).grid(row = 0, column = 2)



    # blenderpath = gui_decal_path(top,0,'blender路径')
    # decalpath = gui_decal_path(top,1,'decal路径')
    # gui_decalnormalize_button(top,2,blenderpath,decalpath)
    # gui_decaladdnormal_button(top,2,blenderpath,decalpath)

    top.mainloop()


if __name__ == "__main__":
    main()