import xlrd
import xlwt
import os
import math


#---------第一步----------#
#读入指导老师给的两个数据源，我重命名为：xxe1.xlsx，xxe2.txt
#读取excel文件中的数据，加到lis列表，excel中的每一行数据是lis列表中的一个元素，lis列表是嵌套列表
#读取excel的每一行时，一般不读第一行，因为第一行是表头，即从第二行（下表索引：1），开始读取
#拿到每行的数据，每行数据本身是一个列表，再加到lis列表中，后面可通过索引取每一个字段内容
lis= []
file_addr = "D://xxe1.xlsx"
if os.path.exists(file_addr):
    xls_file = xlrd.open_workbook(file_addr)
    xls_sheet = xls_file.sheets()[0]
    # 获取行数、列数
    nrows = int(xls_sheet.nrows)
    ncols = int(xls_sheet.ncols)    
    for row in range(1,nrows):
        lis.append(xls_sheet.row_values(row))


#按行读取txt文件中的数据，以逗号分割，读到的每一行数据本身是一个列表，再加到data列表中，data列表是一个嵌套列表
data=[]
fname = 'D://xxe2.txt'
with open(fname,'r+',encoding='utf-8') as f:
    for line in f.readlines(): 
        data.append(line[:-1].split(',')) 



#--------第二步------------#
#lis列表根据Name去重，自定义：不允许同名，即同名的都当作是同一个人处理，去掉一个，保留一个
#需要删除的lis中的元素的索引暂时存放在lisx列表中，扫描完整个lis列表再执行删除操作，注意下标索引问题！！
lisx=[]
for r in range(len(lis)):
    if(r!=len(lis)-1):
        if(lis[r+1][1]==lis[r][1]):
            lisx.append(r)
        else:
            pass
    else:pass
for i in range(len(lisx)):
    lis.pop(lisx[i]-i)


#data列表根据Name去重，自定义：不允许同名，即同名的都当作是同一个人处理，去掉一个，保留一个
#需要删除的data中的元素的索引暂时存放在datax列表中，扫描完整个data列表再执行删除操作，注意下标索引问题！！
datax=[]
for n in range(1,len(data)):
    if(n!=len(data)-1):
        if(data[n+1][1]==data[n][1]):
            datax.append(n)
        else:
            pass
    else:pass
for i in range(len(datax)):
    data.pop(datax[i]-i)


#学号去重，改变其中一个，且改变后的不能和其他学号重复
#基本措施：若两个学号重复了，若后面那个学生的学号简单+1处理后没有跟其他学生学号重复，则成功去重
#若基本措施不行，那么就把该学号改为列表最后一个元素的学号+1，同时该元素移到列表最后
for i in range(1,len(data)):
    if(i!=len(data)-1):
        if(data[i][0]==data[i+1][0]):
            if(str(int(data[i+1][0])+1!=str(int(data[i+2][0])))):
               data[i+1][0]=str(int(data[i+1][0])+1)
            else:
               data[i+1][0]=str(int(data[len(data)][0])+1)
               data,append(data[i+1])
               data.pop(data[i+1])
        else:pass
    else:pass



#-------------第三步----------------#
#以data列表为基准，合并两个列表的数据，利用lis列表中的数据最大限度补充data列表
num=max(len(lis),len(data))
for i in range(1,num):
    dt=data[i][1]
    for j in range(len(lis)):
        if(lis[j][1]==dt):
            for k in range(len(data[0])):
                if(data[i][k]==''):
                    data[i][k]=lis[j][k]
                else:pass
        else:pass



#----------第四步------------#
#数据规整化：
#1、没有值的写NULL
#2、Gender：以male/female表示
#3、Height：单位为m
#4、课程1-10成绩：不为NULL的均以int类型表示
#5、把处理后的data列表导出到【xxe3.xls】中，以便直接查看结果，以验证上面操作的正确性。
for e in range(1,len(data)):
    for q in range(ncols):
        if(data[e][q]==''):
            data[e][q]='NULL'
            
for i in range(1,len(data)):
    if(data[i][3]=='boy'):
        data[i][3]='male'
    elif(data[i][3]=='girl'):
        data[i][3]='female'
    else:
        pass

for i in range(1,len(data)):
    if(data[i][4]>'100'):
        data[i][4]='%.2f'% (float(data[i][4])/100.00)
    else:
        data[i][4]='%.2f' % float(data[i][4])
        
for i in range(1,len(data)):
    for c in range(5,15):
        if(data[i][c]!='NULL'):
            data[i][c]=int(data[i][c])        

result = open('D://xxe_final_data.xls', 'w', encoding='gbk')
for m in range(len(data)):
    for n in range(len(data[m])):
        result.write(str(data[m][n]))
        result.write('\t')
    result.write('\n')
result.close()



#---------第五步-----------#
#开始做1-4题

#第1题:学生中家乡在Beijing的所有课程的平均成绩。
#解释：即把家乡在北京的每一个学生的相同课程的分数加起来除以总人数，一共会得到十个平均数
#成绩索引范围是5-14，家乡索引值：2
p=0
lesson_level=[0,0,0,0,0,0,0,0,0,0]
for i in range(len(data)):
    if(data[i][2]=='Beijing'):
        p+=1
        for y in range(5,15):
            if(data[i][y]!='NULL'):
                lesson_level[y-5]+=int(data[i][y])
            else:pass
    else:pass
      
for i in range(10):
    lesson_level[i]='%.4f' %(lesson_level[i]/p)
print('题目一：\n','一共有【',p,'】名家乡在Beijing的学生，他们课程C0~C10的平均成绩按顺序如下：')
print(lesson_level)
print('\n=======================\n')


#第2题：学生中家乡在广州，课程1在80分以上，且课程9在9分以上的男同学的数量
k=0
for i in range(len(data)):
    if(data[i][2]=='Guangzhou' and data[i][3]=='male'and data[i][5]>80 and data[i][13]>9):
        k+=1
        print('题目二：\n',data[i][1],' 符合筛选条件')
print('一共有【',k,'】名家乡在Guangzhou的男同学，其课程1在80分以上且课程9在9分以上')
print('\n=======================\n')


#第3题：比较广州和上海两地女生的平均体能测试成绩，哪个地区的更强些？
#自定义：excellent=95，good=85,general=75,bad=65，空的就不计算进去人数了
score_G=0
num_G=0
score_S=0
num_S=0
for i in range(len(data)):
    if(data[i][2]=='Guangzhou' and data[i][3]=='female'):
        if(data[i][15]!='NULL'):
            num_G+=1
            if(data[i][15]=='excellent'):
                score_G+=95
            elif(data[i][15]=='good'):
                score_G+=85
            elif(data[i][15]=='general'):
                score_G+=75
            elif(data[i][15]=='bad'):
                score_G+=65
            else:pass
        else:pass
    elif(data[i][2]=='Shanghai' and data[i][3]=='female'):
        if(data[i][15]!='NULL'):
            num_S+=1
            if(data[i][15]=='excellent'):
                score_S+=95
            elif(data[i][15]=='good'):
                score_S+=85
            elif(data[i][15]=='general'):
                score_S+=75
            elif(data[i][15]=='bad'):
                score_S+=65
            else:pass
        else:pass
    else:pass
score_level_G =score_G/num_G
score_level_S=score_S/num_S    
print('题目三：\n','Guangzhou的女生平均体能测试成绩为：',score_level_G)
print(' Shanghai的女生平均体能测试成绩为：',score_level_S)
if(score_level_G>score_level_S):
    print('由此可得：Guangzhou地区的更强些！')
elif(score_level_G<score_level_S):
    print('由此可得：Shanghai地区的更强些！')
else:print('由此可得：两个地区女生体能相当！')
print('\n=======================\n')


#第4题：学习成绩和体能测试成绩，两者的相关性是多少？
#分别计算所有学生课程1-9的学习成绩（即一共9组数据）和体能测试（一组数据）的相关性
#NULL当作0分计算
sum1=[0,0,0,0,0,0,0,0,0,0]
squ_sum=[0,0,0,0,0,0,0,0,0,0]
mean=[0,0,0,0,0,0,0,0,0,0]
standard=[0,0,0,0,0,0,0,0,0,0]
p=[0,0,0,0,0,0,0,0,0]
X_EX=[0,0,0,0,0,0,0,0,0]

#分别求课程1-9学习成绩的总和
for i in range(1,len(data)):
    for y in range(5,14):
        if(data[i][y]!='NULL'):
            sum1[y-5]+=float(data[i][y])
        else:
            sum1[y-5]+=0

#体能测试成绩的总和
for i in range(1,len(data)):
    if(data[i][15]=='excellent'):
        sum1[9]+=95
    elif(data[i][15]=='good'):
        sum1[9]+=85
    elif(data[i][15]=='general'):
        sum1[9]+=75
    elif(data[i][15]=='bad'):
        sum1[9]+=65
    else:sum1[9]+=0

#9门课程学习成绩及体能测试成绩的平均值   
for i in range(len(sum1)):
    mean[i]='%.4f' % (sum1[i]/(len(data)-1))
    if(i==0):
        print('题目四：\n','9门课程平均成绩:',mean[i])
    elif(i==9):
        print('体能测试平均成绩：',mean[i],'\n')
    else:print('                ',mean[i])


#9门课程学习成绩平方和
for i in range(1,len(data)):
    for y in range(5,14):
        if(data[i][y]!='NULL'):
            squ_sum[y-5]+=(float(data[i][y])-float(mean[y-5]))*(float(data[i][y])-float(mean[y-5]))
        else:
            squ_sum[y-5]+=(0-float(mean[y-5]))*(0-float(mean[y-5]))

#体能测试成绩平方和
for i in range(1,len(data)):
    if(data[i][15]=='excellent'):
        squ_sum[9]+=(95-float(mean[9]))*(95-float(mean[9]))
    elif(data[i][15]=='good'):
        squ_sum[9]+=(85-float(mean[9]))*(85-float(mean[9]))
    elif(data[i][15]=='general'):
        squ_sum[9]+=(75-float(mean[9]))*(75-float(mean[9]))
    elif(data[i][15]=='bad'):
        squ_sum[9]+=(65-float(mean[9]))*(65-float(mean[9]))
    else:
        squ_sum[9]+=(0-float(mean[9]))*(0-float(mean[9]))

#9门课程学习成绩及体能测试成绩的标准差
for i in range(len(squ_sum)):
    standard[i]='%.4f' % (math.sqrt(squ_sum[i]/(len(data)-1)))
    if(i==0):
        print('9门课程成绩的标准差:',standard[i])
    elif(i==9):
        print('体能测试成绩的标准差：',standard[i],'\n')
    else:print('                   ',standard[i])


#(x-E(x))*(y-E(y))的和
for i in range(1,len(data)):
    for y in range(5,14):
        if(data[i][y]!='NULL'):
            c=float(data[i][y])-float(mean[y-5])
        else:
            c=0-float(mean[y-5])
        if(data[i][15]=='excellent'):
            g=95-float(mean[9])
        elif(data[i][15]=='good'):
            g=85-float(mean[9])
        elif(data[i][15]=='general'):
            g=75-float(mean[9])
        elif(data[i][15]=='bad'):
            g=65-float(mean[9])
        else:
            g=0-float(mean[9])
        X_EX[y-5]+=c*g

 #计算9对数据的相关系数
for i in range(len(X_EX)):
    p[i]='%.6f' % (X_EX[i]/((len(data)-1)*float(standard[i])*float(standard[9])))
    if(i==0):
        print('9门课程学习成绩跟体能测试成绩的相关系数:',p[i])
    else:print('                                      ',p[i])
    if(i==8):
        print('\n')
print('综上，学习成绩与体能测试成绩相关性如下：')
for i in range(len(p)):
    if(float(p[i])>0):
        if(float(p[i])<0.3):
            print('课程【',i+1,'】的学习成绩跟体能测试成绩无直线相关！')
        elif(0.3<=float(p[i])<0.5):
            print('课程【',i+1,'】的学习成绩跟体能测试成绩低度相关！')
        elif(0.5<float(p[i])<0.8):
            print('课程【',i+1,'】的学习成绩跟体能测试成绩中等程序相关！')
        else:
            print('课程【',i+1,'】的学习成绩跟体能测试成绩高度相关！')
    else:
        if(-float(p[i])<0.3):
            print('课程【',i+1,'】的学习成绩跟体能测试成绩无直线相关！')
        elif(0.3<=-float(p[i])<0.5):
            print('课程【',i+1,'】的学习成绩跟体能测试成绩低度相关！')
        elif(0.5<-float(p[i])<0.8):
            print('课程【',i+1,'】的学习成绩跟体能测试成绩中等程序相关！')
        else:
            print('课程【',i+1,'】的学习成绩跟体能测试成绩高度相关！')
