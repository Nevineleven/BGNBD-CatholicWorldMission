#!/usr/bin/env python
# coding: utf-8

# In[104]:


import csv
import pandas as pd
import datetime as dt
from datetime import datetime
from itertools import islice
import numpy as np
from openpyxl import load_workbook
import collections
import openpyxl as op
import math
import xlsxwriter


# In[105]:


exportFile = 'Model Running.xlsx'
writer = pd.ExcelWriter(exportFile, engine = 'xlsxwriter')


# In[106]:


def datevalue(date1):
    temp = dt.datetime(1899, 12, 30)    # Note, not 31st Dec but 30th!
    delta = date1 - temp
    return float(delta.days) + (float(delta.seconds) / 86400)

def valuedate(fh):
    fh = int(fh)
    return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + fh - 2)


# In[107]:


df_input = pd.read_csv("Input Running.csv")
inputList = df_input["Values"].tolist()
df_input


# In[108]:


predict_thru = dt.datetime(int (inputList[1]), int (inputList[2]), int (inputList[3]))
predict_thru = datevalue(predict_thru)


# In[109]:


df_query3 = pd.read_excel(inputList[0])
df_query3


# In[110]:


dfraw1 = df_query3.copy(deep=True)
removePledge0 = []
for index, rows in dfraw1.iterrows():
    if(dfraw1.iloc[index, 6] == 0 or dfraw1.iloc[index, 7] == ("Pledge")):
        removePledge0.append(True)
    else:
        removePledge0.append("")
dfraw1['remove $0 & pledges'] = (removePledge0)
dfraw1


# In[111]:


dfdp1 = dfraw1.copy(deep=True)
pledgeor0 = dfdp1[dfdp1['remove $0 & pledges'] == True].index
dfdp1.drop(pledgeor0, inplace=True)
dfdp1 = dfdp1.reset_index(drop=True)
dfdp1


# In[112]:


dfraw2 = dfdp1.copy(deep=True)
dfraw2 = dfraw2.drop(columns = ["Gift Amount", "Gift Type", "remove $0 & pledges"])
dfraw2


# In[113]:


adGiftList = [1]
for index, row in islice(dfraw2.iterrows(), 1, None):
    if(dfraw2.iloc[index-1, 0] != dfraw2.iloc[index, 0]):
        adGiftList.append(1)
    else:
        if(dfraw2.iloc[index-1, 0] == dfraw2.iloc[index, 0] and dfraw2.iloc[index-1, 3] != dfraw2.iloc[index, 3]):
            adGiftList.append(adGiftList[index-1] + 1)
        else:
            adGiftList.append(adGiftList[index-1])

dfraw2['adj. num gift'] = (adGiftList)


# In[114]:


len(dfraw2)


# In[115]:


keepList = []
for index, row in islice(dfraw2.iterrows(), 0, None):
    try:
        if (index+1 >= len(dfraw2)):
            keepList.append("")
        elif (dfraw2.iloc[index-1,0] == dfraw2.iloc[index,0] and dfraw2.iloc[index,0] != dfraw2.iloc[index+1,0]):
            keepList.append("")
        else:
            if (dfraw2.iloc[index-1,0] != dfraw2.iloc[index,0] and dfraw2.iloc[index,0] != dfraw2.iloc[index+1,0]):
                keepList.append("")
            else:
                keepList.append(False) 
    except(IndexError):
        print(index)
        
dfraw2['KEEP'] = (keepList)
dfraw2


# In[116]:


dfdp2 = dfraw2.copy(deep=True)
falseKeep = dfdp2[dfdp2['KEEP'] == False].index
dfdp2.drop(falseKeep, inplace=True)
dfdp2 = dfdp2.reset_index(drop=True)
dfdp2


# In[117]:


grgFile = "GRG Running.xlsx"
df_GRG = pd.read_excel(grgFile, header=None, nrows = 5, usecols = [0,1])
r = df_GRG.iloc[0,1]
alpha = df_GRG.iloc[1,1]
a = df_GRG.iloc[2,1]
b = df_GRG.iloc[3,1]
print(r, alpha, a, b)


# In[118]:


dfmd = dfdp2[["Constituent ID", "Name"]]
dfmd["x (#donations)"] = dfdp2["adj. num gift"]-1
dfmd['t_x (last gift)'] = (dfdp2['Gift Date'] - dfdp2['First Gift Date'])/365


# In[119]:


today = pd.to_datetime("today")
today = round(datevalue(today),0)
tempDate = (today)+1
print(tempDate)
dfmd['T (total time span)'] = (tempDate - dfdp2['First Gift Date'])/365
dfmd


# In[120]:


df_allCondExp = dfmd.copy(deep=True)
print(predict_thru, today)
constT = (predict_thru-today)/365
print(constT)
df_allCondExp["T (total time span)"] = df_allCondExp["T (total time span)"]-3
df_allCondExp['t'] = constT
df_allCondExp["E(Y(t)|X=x,t_x,T)"] = ""
df_allCondExp["2F1"] = 0
df_allCondExp["a"] = df_allCondExp.iloc[:,2]+r
df_allCondExp["b"] = df_allCondExp.iloc[:,2] + b
df_allCondExp["c"] = df_allCondExp.iloc[:,2] + b + a - 1
df_allCondExp["z"] = constT/(alpha + df_allCondExp.iloc[:,4] + constT)
df_allCondExp['0'] = 1
df_allCondExp


# In[121]:


aceCount = 1
aceArr = []
while (aceCount <= 200):
    df_allCondExp[str(aceCount)]=df_allCondExp.iloc[:,aceCount+11]*(df_allCondExp.iloc[:,8]+aceCount-1)*(df_allCondExp.iloc[:,9]+aceCount-1)/((df_allCondExp.iloc[:,10]+aceCount-1)*aceCount)*df_allCondExp.iloc[:,11]
    aceCount = aceCount + 1
df_allCondExp


# In[122]:


aceCount1 = 0
while (aceCount1 <= 200):
    df_allCondExp["2F1"] = df_allCondExp["2F1"]+df_allCondExp.iloc[:,aceCount1+12]
    aceCount1 = aceCount1 + 1
df_allCondExp


# In[123]:


longNameArr = []
#((a+b+df_allCondExp.iloc[:,2]-1)/(a-1)*(1-pow(((alpha+df_allCondExp.iloc[:,4])/(alpha+df_allCondExp.iloc[:,4]+df_allCondExp.iloc[:,5])),(r+df_allCondExp.iloc[:,2]))*df_allCondExp.iloc[:,7])/(1+(df_allCondExp.iloc[:,2]>0)*a/(b+df_allCondExp.iloc[:,2]-1)*pow(((alpha+df_allCondExp.iloc[:,4])/(alpha+df_allCondExp.iloc[:,3])),(r+df_allCondExp.iloc[:,2]))))
for index, row in islice(df_allCondExp.iterrows(), 0, None):
    longNameArr.append((a+b+df_allCondExp.iloc[index,2]-1)/(a-1)*(1-pow(((alpha+df_allCondExp.iloc[index,4])/(alpha+df_allCondExp.iloc[index,4]+df_allCondExp.iloc[index,5])),(r+df_allCondExp.iloc[index,2]))*df_allCondExp.iloc[index,7])/(1+(df_allCondExp.iloc[index,2]>0)*a/(b+df_allCondExp.iloc[index,2]-1)*pow(((alpha+df_allCondExp.iloc[index,4])/(alpha+df_allCondExp.iloc[index,3])),(r+df_allCondExp.iloc[index,2]))))
df_allCondExp["E(Y(t)|X=x,t_x,T)"] = longNameArr
df_allCondExp


# In[124]:


df_pAlive = df_allCondExp.iloc[:,0:5].copy(deep=True)
paliveinfo = []
for index, row in islice(df_allCondExp.iterrows(), 0, None):
    paliveinfo.append(1/(1+(df_pAlive.iloc[index,2]>0)*(a/(b+df_pAlive.iloc[index,2]-1))*(pow(((alpha+df_pAlive.iloc[index,4])/(alpha+df_pAlive.iloc[index,3])),(r+df_pAlive.iloc[index,2])))))
df_pAlive["P(Alive) Info"] = paliveinfo
df_pAlive


# In[125]:


dfraw2.to_excel(writer, 'Raw data2', index = False)
dfraw1.to_excel(writer, 'Raw data1', index = False)
dfraw1.to_excel(writer, 'Data prep1', index = False)
dfdp2.to_excel(writer, 'Data prep2', index = False)
dfmd.to_excel(writer, 'Model data', index = False)
df_allCondExp.to_excel(writer, "All Cond. Exp.", index=False)
df_pAlive.to_excel(writer, "P(alive)", index=False)
writer.save()

