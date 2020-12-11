import matplotlib.pyplot as plt
import os
import xlrd
import numpy as npy
import math

#显示中文需引入模块：pylab
#可设置中文字体，此处设置为黑体，同时解决保存图像时负号无法显示问题
from pylab import mpl
mpl.rcParams['font.sans-serif']=['SimHei']
mpl.rcParams['axes.unicode_minus']=False


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
for i in range(len(lis)):
    for j in range(5,14):
        if(lis[i][j]=='NULL'):
            lis[i][j]=0
for i in range(len(lis)):
    if(lis[i][15]=='NULL'):
        lis[i][15]=0

      
#课程1成绩和体能成绩分别放进两个列表中  
scatter_class1=[]
scatter_strong=[]
for i in range(len(lis)):
    scatter_class1.append(lis[i][5])
for i in range(len(lis)):
    if(lis[i][15]=='excellent'):
        scatter_strong.append(95)
    elif(lis[i][15]=='good'):
        scatter_strong.append(85)
    elif(lis[i][15]=='general'):
        scatter_strong.append(75)
    elif(lis[i][15]=='bad'):
        scatter_strong.append(65)
    else:
         scatter_strong.append(0)




###----------绘制散点图------------#
##
###第1题：以课程1成绩为x轴，体能成绩为y轴，画出散点图
###体能成绩处理方案：excellent=95，good=85,general=75,bad=65
###本来有97对数据，但散点图显示却只有45个，解释：因为有重叠！！！ 
##    
###定义标题，定义坐标轴x,y，以及进行刻度设置
##plt.title('体能成绩 & 课程1成绩--散点图',fontsize=20,color='blue')
##plt.xlabel('课程1成绩',fontsize=15,color='r')
##plt.ylabel('体能成绩',fontsize=15,color='r')
##my_x=npy.arange(0,100,5)
##my_y=npy.arange(0,100,10)
##plt.xticks(my_x)
##plt.yticks(my_y)
##
###绘制 体能成绩 & 课程1成绩--散点图 并显示
##plt.scatter(scatter_class1,scatter_strong,marker=',')
##plt.show()




###------------绘制直方图-------------#
##         
###第2题：2. 以5分为间隔，画出课程1的成绩直方图
###定义标题，定义坐标轴x,y，以及进行刻度设置
##plt.title('课程1--成绩直方图',fontsize=20,color='blue')
##plt.xlabel('课程1成绩',fontsize=15,color='r')
##plt.ylabel('在该成绩区间的人数',fontsize=12,color='r')
##my_x=npy.arange(0,100,5)
##plt.xticks(my_x)
##
###绘制 课程1--成绩直方图 并显示，间隔为5分
##jiange=[]
##for i in range(0,100,5):
##    jiange.append(i)
##n, bins, patches = plt.hist(scatter_class1,jiange)
##plt.show()



#-----------归一化-------------#

#第3题：对每门成绩进行z-score归一化，得到归一化的数据矩阵
#课程10成绩全为0，因此只算9门课程学习成绩和体能成绩，最后得到97行10列的数据矩阵
#初始化一些列表，分别是成绩和，均值，平方和，方差
sum1=[0,0,0,0,0,0,0,0,0,0]
level=[0,0,0,0,0,0,0,0,0,0]
squ=[0,0,0,0,0,0,0,0,0,0]
variance=[0,0,0,0,0,0,0,0,0,0]

#求各科成绩总和，然后除以列表lis总长度得到10个成绩的均值
for i in range(len(lis)):
    for j in range(5,14):
        sum1[j-5]+=lis[i][j]
for i in range(len(lis)):
    if(lis[i][15]=='excellent'):
        sum1[9]+=95
    elif(lis[i][15]=='good'):
        sum1[9]+=85
    elif(lis[i][15]=='general'):
        sum1[9]+=75
    elif(lis[i][15]=='bad'):
        sum1[9]+=65
    else:sum1[9]+=0
for i in range(len(sum1)):
    level[i]=sum1[i]/len(lis)

#求各科成绩与均值差的平方和，然后除以列表lis总长度，得到10个成绩的方差
for i in range(len(lis)):
    for j in range(5,14):
        squ[j-5]+=(lis[i][j]-level[j-5])**2
for i in range(len(lis)):
    if(lis[i][15]=='excellent'):
        squ[9]+=(95-level[9])**2
    elif(lis[i][15]=='good'):
        squ[9]+=(85-level[9])**2
    elif(lis[i][15]=='general'):
        squ[9]+=(75-level[9])**2
    elif(lis[i][15]=='bad'):
        squ[9]+=(65-level[9])**2
    else:squ[9]+=(0-level[9])**2
for i in range(len(squ)):
    variance[i]=math.sqrt(squ[i]/len(lis))

#进行Z-score归一化：（初值-均值）/ 方差，保留4位小数
for i in range(len(lis)):
    for j in range(5,14):
        lis[i][j]="%.4f" % ((lis[i][j]-level[j-5])/variance[j-5])
for i in range(len(lis)):
    if(lis[i][15]=='excellent'):
        lis[i][15]="%.4f" % ((95-level[9])/variance[9])
    elif(lis[i][15]=='good'):
        lis[i][15]="%.4f" % ((85-level[9])/variance[9])
    elif(lis[i][15]=='general'):
        lis[i][15]="%.4f" % ((75-level[9])/variance[9])
    elif(lis[i][15]=='bad'):
        lis[i][15]="%.4f" % ((65-level[9])/variance[9])
    else:lis[i][15]="%.4f" % ((0-level[9])/variance[9])

#分别将每一行的9门学习成绩和体能成绩放进嵌套列表这种对应位置的元素中
#然后进行列表转矩阵，并打印数据矩阵
#将列表a1写入excel文件中，方便查看验证
a1=[]
for i in range(len(lis)):
    a1.append([])
for i in range(len(lis)):
    for j in range(5,14):
        a1[i].append(lis[i][j])
    a1[i].append(lis[i][15])

a2=npy.matrix(a1)
print("归一化的数据矩阵：\n%s"%a2)

result = open('D://第3题：归一化后列表数据.xls', 'w', encoding='gbk')
title=['lesson1','lesson2','lesson3','lesson4','lesson5','lesson6','lesson7','lesson8','lesson9','体能成绩']
for i in range(len(title)):
    result.write(title[i])
    result.write('\t')
result.write('\n')
for m in range(len(a1)):
    for n in range(len(a1[m])):
        result.write(str(a1[m][n]))
        result.write('\t')
    result.write('\n')
result.close()
