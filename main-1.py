import openpyxl
wb = openpyxl.load_workbook('in.xlsx')
ws2 = wb.create_sheet()
ws = wb.active

allworkname = ['部门','姓名','总工时']        #代号lists
allworktime = []        #工时list
cloumn_length = len(list(ws.rows))-2      #总表长度


#——————————————————————————————————————————生成总代号list—————allworkname———————————————————————————————————————
for n in range(cloumn_length):
    #计算总工时
    worktime = 0        #初始化总工时计数器
    col = 11     #数据位置指针初始化 
    for x in range(7):      #每周7次工时累加
        ex1 = ws.cell(row=n+3,column=col)     #项目代号位置
        ex2 = ws.cell(row=n+3,column=col+1)       #项目工时位置
        col = col + 3       #数据位置指针周内移动
        if ex2.value != None:       #排除空单元格
            worktime = worktime + ex2.value    #总工时累加
    col = col + 1       #数据位置指针移动至周开始