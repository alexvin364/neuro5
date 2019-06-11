from random import uniform
from termcolor import colored
import colorama
import xlrd
colorama.init()

rb = xlrd.open_workbook('.\data-7422-2019-05-10.xlsx',formatting_info=False)
sheet = rb.sheet_by_index(0)

row = []
for i in range(sheet.nrows):
    r = sheet.row_values(i)
    try:
        row.append([r[1],float(r[16]),0])
    except:
        pass

#for i in range(200):
#    row.append(["Музей {}".format(i+1),round(uniform(0,1000000),1),0])
x = []
s = []
for i in range(len(row)):
    x.append(float(row[i][1]))
    s.append(0)

norma = 0.2
colors = ("red","yellow","blue","green")

r = []
w = []
for i in range(len(colors)):
    w.append(0)
    r.append(0)

t = True
while (t==True):
    t = False
    for i in range(len(row)):
        s = abs(x[i]-w[0])
        k = 0
        for j in range(len(colors)):
            r[j]=abs(x[i]-w[j])
            if (r[j]<s):
                s = r[j]
                k = j
        if (row[i][2]!=k):
            t = True
        row[i][2] = k
        w[k] += norma*(x[i]-w[k])
        
for i in range(len(row)):    
    print(colored("{:<10}{}".format(row[i][1],row[i][0]),colors[row[i][2]]))
print()
for i in range(len(w)):
    print(colored(round(w[i],2),colors[i]))
