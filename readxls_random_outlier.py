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
#读取企业名称
NAME = table4fit.col_values(1)[1:50]
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

print("敲掉的列  ","[   预测值  ]","  R2分数")
print("---------全数据拟合---------")
lm.fit(X,y)
print("All Data:",lm.predict(T),lm.score(X,y))

datalength = X.shape[0]
bestR2 = -1.0 #
bestIndexString = ""
displayon = False#True

print("\n--------- 敲掉一行 ---------")
for i in range(datalength):
    X_new = np.delete(X,i,0)
    y_new = np.delete(y,i,0)
    lm.fit(X_new,y_new)
    currentscore =lm.score(X_new,y_new)
    currentIndex = "("+str(i)+":"+NAME[i]+")"
    if(currentscore > bestR2):
        bestR2 = currentscore
        bestIndexString = currentIndex
    if(displayon):print(currentIndex,lm.predict(T),currentscore)
    
print("\n--------- 敲掉二行 ---------")
for i in range(datalength):
    for j in range((i+1),datalength):
        X_new = np.delete(X,[i,j],0)
        y_new = np.delete(y,[i,j],0)
        lm.fit(X_new,y_new)
        currentscore =lm.score(X_new,y_new)
        currentIndex = "("+str(i)+":"+NAME[i]+","+str(j)+":"+NAME[j]+")"
        if(currentscore > bestR2):
            bestR2 = currentscore
            bestIndexString = currentIndex
        if(displayon):print(currentIndex,lm.predict(T),currentscore)

print("\n--------- 敲掉三行 ---------")
for i in range(datalength):
    for j in range((i+1),datalength):
        for k in range((j+1),datalength):
            X_new = np.delete(X,[i,j,k],0)
            y_new = np.delete(y,[i,j,k],0)
            lm.fit(X_new,y_new)
            currentscore =lm.score(X_new,y_new)
            currentIndex = "("+str(i)+":"+NAME[i]+","+str(j)+":"+NAME[j]+","+str(k)+":"+NAME[k]+")"
            if(currentscore > bestR2):
                bestR2 = currentscore
                bestIndexString = currentIndex
            if(displayon):print(currentIndex,lm.predict(T),currentscore)

print("\n 最好R2："+str(bestR2)+" （敲除企业）@"+bestIndexString)
#最好R2：0.72669916063 （敲除企业）@(36:通化东宝,45:新日恒力,48:中源协和)
