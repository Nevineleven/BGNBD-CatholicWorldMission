#!/usr/bin/env python
# coding: utf-8

# In[248]:


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


# In[249]:


r = 0.147642304763359
alpha = 0.351575403514246
a = 0.327857205096383
b = 1.36694414288691


# In[250]:


exportFile = 'Model construction CWM bgnbd2.xlsx'
writer = pd.ExcelWriter(exportFile, engine = 'openpyxl')
grgFile = "GRG Running.xlsx"
writer2 = pd.ExcelWriter(grgFile, engine = 'openpyxl')


# In[251]:


df_input = pd.read_csv("Input.csv")
inputList = df_input["Values"].tolist()
df_input


# In[252]:


beginning_date = dt.datetime(int (inputList[1]), int (inputList[2]), int (inputList[3]))
predict_thru = dt.datetime(int (inputList[4]), int (inputList[5]), int (inputList[6]))


# In[253]:


def datevalue(date1):
    temp = dt.datetime(1899, 12, 30)    # Note, not 31st Dec but 30th!
    delta = date1 - temp
    return float(delta.days) + (float(delta.seconds) / 86400)


# In[254]:


def valuedate(fh):
    fh = int(fh)
    return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + fh - 2)


# In[255]:


beginning_date = datevalue(beginning_date)
predict_thru = datevalue(predict_thru)


# In[256]:


endDate = beginning_date + .5*(predict_thru-beginning_date)
firstTimeTo = beginning_date + .31 * .5 * (predict_thru-beginning_date)
print(endDate)


# In[257]:


df_query2 = pd.read_excel(inputList[0]);
df_query2


# In[258]:


dfraw1 = df_query2.copy(deep=True)
removePledge0 = []
for index, rows in dfraw1.iterrows():
    if(dfraw1.iloc[index, 5] == 0 or dfraw1.iloc[index, 6] == ("Pledge")):
        removePledge0.append(True)
    else:
        removePledge0.append("")
dfraw1['remove $0 & pledges'] = (removePledge0)
dfraw1


# In[259]:


dfdp1 = dfraw1.copy(deep=True)
pledgeor0 = dfdp1[dfdp1['remove $0 & pledges'] == True].index
dfdp1.drop(pledgeor0, inplace=True)
dfdp1 = dfdp1.reset_index(drop=True)
dfdp1


# In[260]:


dfraw2 = dfdp1.copy(deep=True)
dfraw2 = dfraw2.drop(columns = ["Gift Amount", "Gift Type", "remove $0 & pledges"])
dfraw2


# In[261]:


adGiftList = [1]
print(endDate)
for index, row in islice(dfraw2.iterrows(), 1, None):
    if(dfraw2.iloc[index-1, 0] != dfraw2.iloc[index, 0]):
        adGiftList.append(1)
    elif(dfraw2.iloc[index,4]>endDate):
            adGiftList.append(0)
    else:
        if(dfraw2.iloc[index-1, 0] == dfraw2.iloc[index, 0] and dfraw2.iloc[index-1, 4] != dfraw2.iloc[index, 4]):
            adGiftList.append(adGiftList[index-1] + 1)
        else:
            adGiftList.append(adGiftList[index-1])

dfraw2['adj. num gift'] = (adGiftList)


# In[262]:


keepList = []
for index, row in islice(dfraw2.iterrows(), 1, None):
    if ((dfraw2.iloc[index-1, 3]<=endDate and dfraw2.iloc[index-1,0] != dfraw2.iloc[index,0]) or (dfraw2.iloc[index-1, 3]>endDate and dfraw2.iloc[index-1, 4]<=endDate and dfraw2.iloc[index, 4]>endDate)):
        keepList.append("")
    else:
        keepList.append(False)

        
if(dfraw2.iloc[-1, 3]<=endDate):
    keepList.append("")
else:
    keepList.append(False)
        
dfraw2['KEEP'] = (keepList)


# In[263]:


rptList = [""]
for index, row in islice(dfraw2.iterrows(), 1, None):
    if (dfraw2.iloc[index-1, 0] == dfraw2.iloc[index, 0] and dfraw2.iloc[index, 4]>endDate and dfraw2.iloc[index-1, 4] == dfraw2.iloc[index, 4]):
        rptList.append("blah")
    else:
        rptList.append("")
        
dfraw2['rpt gift after cali'] = (rptList)


# In[264]:


dfdp2 = dfraw2.drop(columns = 'rpt gift after cali')
falseKeep = dfdp2[dfdp2['KEEP'] == False].index
dfdp2.drop(falseKeep, inplace=True)
dfdp2 = dfdp2.reset_index(drop=True)
dfdp2


# In[265]:


dfmd = dfdp2[["Constituent ID", "Name"]]
dfmd["x (#donations)"] = dfdp2["adj. num gift"]-1
dfmd['t_x (last gift)'] = (dfdp2['Gift Date'] - dfdp2['First Gift Date'])/365


# In[266]:


tempDate = endDate+1
dfmd['T (total time span)'] = (tempDate - dfdp2['First Gift Date'])/365
dfmd


# In[267]:


df_BGNBD = dfmd.copy(deep=True)
donationNum = df_BGNBD['x (#donations)'].tolist()
lastGift = df_BGNBD['t_x (last gift)'].tolist()
T = df_BGNBD['T (total time span)'].tolist()


# In[268]:


A4 = []
for index, rows in islice(df_BGNBD.iterrows(), 0, None):
    if donationNum[index]>0:
        A4.append(math.log(a)-math.log(b+donationNum[index]-1)-(r+donationNum[index])*math.log(alpha+lastGift[index]))
    else:
        A4.append(0)
df_BGNBD['ln(A_4)'] = (A4)
A_4 = df_BGNBD['ln(A_4)'].tolist()


# In[269]:


A3 = []
for index, rows in islice(df_BGNBD.iterrows(), 0, None):
    A3.append(-(r+donationNum[index])*math.log(alpha+T[index]))
df_BGNBD['ln(A_3)'] = (A3)
A_3 = df_BGNBD['ln(A_3)'].tolist()


# In[270]:


A2 = []
for index, rows in islice(df_BGNBD.iterrows(), 0, None):
    A2.append(math.lgamma(a+b)+math.lgamma(b+donationNum[index])-math.lgamma(b)-math.lgamma(a+b+donationNum[index]))
df_BGNBD['ln(A_2)'] = (A2)
A_2 = df_BGNBD['ln(A_2)'].tolist()


# In[271]:


A1 = []
for index, rows in islice(df_BGNBD.iterrows(), 0, None):
    A1.append(math.lgamma(r+donationNum[index])-math.lgamma(r)+r*math.log(alpha))
df_BGNBD['ln(A_1)'] = (A1)
A_1 = df_BGNBD['ln(A_1)'].tolist()
df_BGNBD


# In[272]:


ln = []
for index, rows in islice(df_BGNBD.iterrows(), 0, None):
    if donationNum[index]>0:
        ln.append(A_1[index]+A_2[index]+math.log(math.exp(A_3[index])+1*math.exp(A_4[index])))
    else:
        ln.append(A_1[index]+A_2[index]+math.log(math.exp(A_3[index])+0*math.exp(A_4[index])))
df_BGNBD['ln(.)'] = (ln)


# In[273]:


# BGNBDcols = list(df_BGNBD.columns.values)
# df_BGNBD = df_BGNBD[BGNBDcols[0:5] + [BGNBDcols[-1]] + [BGNBDcols[-2]] + [BGNBDcols[-3]] + [BGNBDcols[-4]] + [BGNBDcols[-5]]]
# df_BGNBD


# In[274]:


dfraw1.to_excel(writer, 'Raw data1', index = False)
dfdp1.to_excel(writer, 'Data Prep1', index = False)
dfraw2.to_excel(writer, 'Raw data2', index = False)
dfdp2.to_excel(writer, 'Data Prep2', index = False)
dfmd.to_excel(writer, 'Model Data', index = False)
df_BGNBD.to_excel(writer, "BGNBD Estimation", index = False)
df_BGNBD.to_excel(writer2, "BGNBD Estimation", index = False)
writer.save()
writer2.save()

