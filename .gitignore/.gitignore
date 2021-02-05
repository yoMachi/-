# coding=utf-8
import os,sys
import re
import xlwt


path = sys.path[0]


for root, dirs, files in os.walk(path): 
    if root == path:
        for file in files:  
            (filename, extension) = os.path.splitext(file) 
            
            if extension == '.txt':
                xls = ('[转换]' + filename + '.xls')

                txt = (filename + extension) 
        
path_txt = path + txt
path_xls = path + xls
f = open(path_txt,'r')
workbook = xlwt.Workbook(encoding='utf-8')       #新建工作簿
sheet1 = workbook.add_sheet(filename)          #新建sheet
sheet1.col(0).width = 200*30
sheet1.col(1).width = 200*30
sheet1.write(0,0,'Y')
sheet1.write(0,1,'X')

zhPattern = re.compile(u'[\u4e00-\u9fa5]+')
row = 1
for line in f:
    match = zhPattern.search(line)
    line = line.rstrip('\n')
    line = line.split(',') #每一行以","分隔
    if match:
        print('这行不写')
    elif len(line) < 4:
        print('这行不写')
    else:
        sheet1.write(row,0, line[2])
        sheet1.write(row,1, line[3])
        workbook.save(path_xls) #输出在同一目录
        row += 1

f.close()
os.remove(path_txt)


