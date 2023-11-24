import os
import xlrd,xlwt
from xlwt.Workbook import Workbook
from xlwt.Worksheet import Worksheet


wenjianjia=input('输入需要处理的文件夹')
cell_names=[]
file_dir=os.path.realpath(__file__)
file_dir=file_dir[:-13] 

cell_filename=[]
i=0
for files in os.walk(file_dir):
    if i==0:
        i+=1
    elif i==1:
        cell_filename.append(files)
#取文件夹下文件名
     
cell_filename=cell_filename[0][2]

for p in cell_filename:
    if os.path.splitext(p)[1] == '.xls':
        cell_names.append(p)
#取xls文件名



for j in cell_names:#j是单个子电池的名称
    
    temp_filename=str(file_dir+wenjianjia+'\\'+j)
    temp_wb=xlrd.open_workbook(temp_filename)
    temp_ws=temp_wb.sheets()[0]
    #进入单个子电池工作表
    
    rowl=range(temp_ws.nrows)
    filename=str(j+'.txt')
    f = open(filename, 'w')
    for yui in rowl:
        Temp_AnodeV=str(temp_ws.cell(yui,3))[7:]
        f.write(Temp_AnodeV)
        f.write('\t')
        Temp_AnodeI=str(temp_ws.cell(yui,2))[7:]
        f.write(Temp_AnodeI)
        f.write('\n')
        f.flush()
    f.close()
           



      



    


