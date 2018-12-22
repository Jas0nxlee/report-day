import openpyxl
wb = openpyxl.load_workbook('in.xlsx')
ws2 = wb.create_sheet()
ws = wb.active
worktime = 0
null = ''
cloumn_length = len(list(ws.rows))-2      #总表长度


#——————————————————————————————————             总工时———————————————————————————————————————
for n in range(cloumn_length):    #计算总工时
    col = 10
    for x in range(5):      #每周7次工时累加
        ex1 = ws.cell(row=n+3,column=col+1)     #项目代号位置
        ex2 = ws.cell(row=n+3,column=col+2)       #项目工时位置
        col = col + 3
        if  ex2.value != null and ex2.value != None:       #排除空单元格
            worktime += float(ex2.value)    #总工时累加
print(worktime)


#—————————————————————————————————————个人项目工时列表————————————————————————————————————————
all = []
aa = []
bbtmp = []
worklist = []
allworktime = 0
for n in range(cloumn_length):
    ex1 = ws.cell(row=n+3,column=1)
    ex2 = ws.cell(row=n+3,column=2)
    if ex1.value not in bbtmp:
        aa.append(ex1.value)
        aa.append(ex2.value)
        bbtmp.append(ex1.value)
        for p in range(cloumn_length):
            ex3 = ws.cell(row=p+4,column=1)
            ex4 = ws.cell(row=p+4,column=2)
            if ex3.value == ex1.value:
                col = 10
                for x in range(5):
                    ex5 = ws.cell(row=p+3,column=col+1)
                    ex6 = ws.cell(row=p+3,column=col+2)
                    col = col + 3
                    if  ex6.value != null and ex6.value != None:       #排除
                        allworktime += float(ex6.value)
                        aa.append(ex5.value)
                        aa.append(ex6.value)
        aa.insert(2,allworktime)
        all.append(aa)
        aa = []
        allworktime = 0
#print(all)

#—————————————————————————————————————————————————————————————————————————————
for n in range(len(all)):
    aaa = all[n]
    print(aaa)
    111
    