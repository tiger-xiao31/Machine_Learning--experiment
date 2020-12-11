import matplotlib.pyplot as plt
import os
import xlrd
import numpy as npy
import math
import seaborn as sn 
import pandas as pd


#读取excel文件中的数据存入lis
lis= []
file_addr = "D://xxe_two（数据源）.xlsx"
if os.path.exists(file_addr):
    xls_file = xlrd.open_workbook(file_addr)
    xls_sheet = xls_file.sheets()[0]
    # 获取行数、列数
    nrows = int(xls_sheet.nrows)
    ncols = int(xls_sheet.ncols)    
    for row in range(1,nrows):
        lis.append(xls_sheet.row_values(row))


#值为“NULL”的成绩赋值为0
#课程6-9原本是十分制，均换算成100分制计算
for i in range(len(lis)):
    for j in range(5,14):
        if(lis[i][j]=='NULL'):
            lis[i][j]=0
for i in range(len(lis)):
    for j in range(10,14):
        lis[i][j]=lis[i][j]*10


#把9门成绩放入lesson列表中，一个子列表有9门成绩
lesson=[]
for i in range(len(lis)):
    lesson.append([])
for i in range(len(lis)):
    for j in range(5,14):
        lesson[i].append(lis[i][j])

#-----------相关矩阵、混淆矩阵------------#
#第4题：计算出100x100的相关矩阵，并可视化出混淆矩阵（我这里是97x97）
#求标准差
std=[]
for i in range(len(lesson)):
    std.append([])
for i in range(len(lesson)):
        std[i].append(npy.std(lesson[i]))


#求均值
sum1=[]
level=[]
for i in range(len(lesson)):
    sum1.append(0)
    level.append(0)
for i in range(len(lesson)):
    for j in range(9):
        sum1[i]+=lesson[i][j]
    level[i]=sum1[i]/len(lesson[i])


#求各行数据之间的协方差
result1=[]
sum2=[]
for i in range(len(lesson)):
    sum2.append([])
    result1.append([])
for i in range(len(lesson)):
    for j in range(len(lesson)):
        sum2[i].append(0)
        result1[i].append(0)

for i in range(len(lesson)):
    for j in range(len(lesson)):
        for k in range(9):
            sum2[i][j]+=(lesson[i][k]-level[i])*(lesson[j][k]-level[j])
        sum2[i][j]=sum2[i][j]/len(lesson[i])
        result1[i][j]='%.6f'%(sum2[i][j]/(std[i][0]*std[j][0]))

#将各行时间的相关系数写入excel文件中，方便查看与验证
#同时将单纯的相关系数存入又一个excel文件中
#后将数据复制存入后缀为xlsx的文件中，方便可视化混淆矩阵时导入、传值
#原因：result1直接传值会出错，原因不详...，所以出此下策
result = open('D://第4题：各行之间相关系数.xls', 'w', encoding='gbk')
title=['姓名','行号']
for i in range(len(result1)):
    title.append(str(i))
for i in range(len(title)):
    result.write(title[i])
    result.write('\t')
result.write('\n')
for m in range(len(result1)):
    result.write(lis[m][1])
    result.write('\t')
    result.write(title[m+2])
    result.write('\t')
    for n in range(len(result1[m])):
        result.write(str(result1[m][n]))
        result.write('\t')
    result.write('\n')
result.close()


result = open('D://第4题：混淆矩阵所需导入数据.xls', 'w', encoding='gbk')
for m in range(len(result1)):
    for n in range(len(result1[m])):
        result.write(str(result1[m][n]))
        result.write('\t')
    result.write('\n')
result.close()


#打印相关矩阵
result2=npy.matrix(result1)
print('相关矩阵“\n%s'%result2)


#可视化混淆矩阵
b=[]
file_addr = "D://第4题：混淆矩阵需导入的数据.xlsx"
if os.path.exists(file_addr):
    xls_file = xlrd.open_workbook(file_addr)
    xls_sheet = xls_file.sheets()[0]
    # 获取行数、列数
    nrows = int(xls_sheet.nrows)
    ncols = int(xls_sheet.ncols)    
    for row in range(nrows):
        b.append(xls_sheet.row_values(row))

confusion_matrix=b
df_cm=pd.DataFrame(confusion_matrix)
sn.heatmap(df_cm,vmax=1,vmin=-1)
plt.show()


#----------距离最近的样本----------#
#第5题：根据相关矩阵，找到距离每个样本最近的三个样本，
#得到97x3的矩阵（每一行为对应三个样本的ID）输出到txt文件中，以\t,\n间隔
short=[]
for i in range(len(result1)):
    short.append([])
max1=[]      #记录绝对值最大的的三个数据本身
max_label=[] #记录绝对值最大的的三个数据的行数

for i in range(len(b)):
    max1.append(b[i][0])
    max1.append(b[i][1])
    max1.append(b[i][2])
    max1.sort()   #升序
    for j in range(len(b)):
        if(abs(float(b[i][j]))<abs(float(max1[0]))):
            pass
        elif(j!=i and abs(float(b[i][j]))>abs(float(max1[0]))):
            max1[0]=b[i][j]
            max1.sort()
        else:pass
    max_label.append(b[i].index(max1[0]))
    max_label.append(b[i].index(max1[1]))
    max_label.append(b[i].index(max1[2]))
    short[i].append(lis[max_label[0]][0])
    short[i].append(lis[max_label[1]][0])
    short[i].append(lis[max_label[2]][0])
    max1=[]
    max_label=[]


#打印距离某样本最近的另三个样本的ID
for i in range(len(short)):
    short[i].sort()
    print(short[i])

#把sshort列表数据转换为字符串，才能进行下面的导出到txt文件
for i in range(len(short)):
    for j in range(3):
        short[i][j]=str(short[i][j])

#导出到txt文件
file=open("D://第5题：距离某样本最近的另三个样本ID.txt",'a')
for i in range(len(short)):
    s = str(short[i]).replace('[','').replace(']','')
    s = s.replace("'",'').replace(',','') +'\n'
    file.write(s)
file.close()
print("保存文件成功")   

