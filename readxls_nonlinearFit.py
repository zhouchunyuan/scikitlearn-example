import xlrd
import matplotlib.pyplot as plt
import numpy as np
from sklearn import datasets, linear_model
from sklearn.metrics import mean_squared_error, r2_score

#打开xls数据(用于线性拟合的数据)
book4fit = xlrd.open_workbook('HDJY-2013-select-SPSS.xls')
#打开表单0
table4fit = book4fit.sheets()[0]
#读取表头
head = table4fit.row_values(0)[2:16]
#准备y数据，PE在第16列
PE = table4fit.col_values(16)[1:50]
#准备X数据
data4fit = []
for i in range(1,table4fit.nrows):
    data4fit.append(table4fit.row_values(i)[2:16])
#转换成numpy array    
X = np.array(data4fit)
y = np.array(PE)

#打开xls数据(用于预测的数据)
book4predict = xlrd.open_workbook('HDJY 解释参数.xls')
#打开表单0
table14predict = book4predict.sheets()[0]
#读取预测数据
data4predict = []
for i in range(3,5):#只有两行（3：3+2）
    data4predict.append(table14predict.row_values(i)[2:16])
    
#T 表示 test data
T = np.array(data4predict)
lm = linear_model.LinearRegression()

#对X进行非线性操作
for i in range(14):
    X[:,i]=X[:,i]*X[:,i]
    T[:,i]=T[:,i]*T[:,i]
    
print("敲掉的列  ","[   预测值  ]","  R2分数")
print("---------全数据拟合---------")
lm.fit(X,y)
print("All Data:",lm.predict(T),lm.score(X,y))

print("\n--------- 敲掉一列 ---------")
for i in range(14):
    X_new = np.delete(X,i,1)
    T_new = np.delete(T,i,1)
    lm.fit(X_new,y)
    print("{:<9}".format(head[i]),lm.predict(T_new),lm.score(X_new,y))
    
print("\n--------- 敲掉二列 ---------")
for i in range(14):
    for j in range((i+1),14):
        X_new = np.delete(X,[i,j],1)
        T_new = np.delete(T,[i,j],1)
        lm.fit(X_new,y)
        print("{:<7}".format(head[i]),"{:<7}".format(head[j]),lm.predict(T_new),lm.score(X_new,y))

print("\n--------- 敲掉三列 ---------")
for i in range(14):
    for j in range((i+1),14):
        for k in range((j+1),14):
            X_new = np.delete(X,[i,j,k],1)
            T_new = np.delete(T,[i,j,k],1)
            lm.fit(X_new,y)
            print("{:<7}".format(head[i]),"{:<7}".format(head[j]),"{:<7}".format(head[k]),lm.predict(T_new),lm.score(X_new,y))
