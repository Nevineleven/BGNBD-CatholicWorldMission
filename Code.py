#!/usr/bin/env python
# coding: utf-8

# In[50]:


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


# In[51]:


beginning_date = dt.datetime(2011, 1, 1)
predict_thru = dt.datetime(2018, 8, 6)
fileName = "Small Data Set.xlsx"
r = 0.147642304763359
alpha = 0.351575403514246
a = 0.327857205096383
b = 1.36694414288691


# In[52]:


exportFile = 'Model construction CWM bgnbd1.xlsx'
writer = pd.ExcelWriter(exportFile, engine = 'openpyxl')


# In[53]:


def datevalue(date1):
    temp = dt.datetime(1899, 12, 30)    # Note, not 31st Dec but 30th!
    delta = date1 - temp
    return float(delta.days) + (float(delta.seconds) / 86400)


# In[54]:


def valuedate(fh):
    fh = int(fh)
    return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + fh - 2)


# In[55]:


beginning_date = datevalue(beginning_date)
predict_thru = datevalue(predict_thru)


# In[56]:


endDate = beginning_date + .5*(predict_thru-beginning_date)
firstTimeTo = beginning_date + .31 * .5 * (predict_thru-beginning_date)
print(endDate)


# In[57]:


# read file and drop the empty column
dfraw2 = pd.read_excel(fileName, sheet_name='Raw data2')
dfraw2 = dfraw2.drop(columns = ['Unnamed: 2'])


# In[58]:


#create First Gift Column and reorder
allGiftDates = dfraw2.groupby(['Constituent ID']).describe()['Gift Date']['min']
allIDs = dfraw2.groupby(['Constituent ID']).describe()
my_list = []
for index, rows in dfraw2.iterrows():
    my_list.append(allGiftDates[(rows['Constituent ID'])])
    
dfraw2['First Gift Date'] = my_list

cols = list(dfraw2.columns.values)
dfraw2 = dfraw2[cols[0:2] + [cols[-1]] + cols[2:-1]]


# In[59]:


adGiftList = [1]
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


# In[60]:


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


# In[61]:


rptList = [""]
for index, row in islice(dfraw2.iterrows(), 1, None):
    if (dfraw2.iloc[index-1, 0] == dfraw2.iloc[index, 0] and dfraw2.iloc[index, 4]>endDate and dfraw2.iloc[index-1, 4] == dfraw2.iloc[index, 4]):
        rptList.append("blah")
    else:
        rptList.append("")
        
dfraw2['rpt gift after cali'] = (rptList)


# In[62]:


dfraw2
#df.to_excel(writer, 'Raw data2', index = False)
#writer.save()


# In[63]:


#figure out Query2


# In[64]:


dfraw1 = pd.read_excel(writer, 'Raw data1')
dfraw1


# In[65]:


removePledge0 = []
for index, rows in dfraw1.iterrows():
    if(dfraw1.iloc[index, 5] == 0 or dfraw1.iloc[index, 6] == ("Pledge")):
        removePledge0.append(True)
    else:
        removePledge0.append("")
dfraw1['remove $0 & pledges'] = (removePledge0)
dfraw1


# In[66]:


dfdp1 = dfraw1.copy(deep=True)
pledgeor0 = dfdp1[dfdp1['remove $0 & pledges'] == True].index
dfdp1.drop(pledgeor0, inplace=True)
dfdp1 = dfdp1.reset_index(drop=True)
dfdp1


# In[67]:


dfdp2 = dfraw2.drop(columns = 'rpt gift after cali')
falseKeep = dfdp2[dfdp2['KEEP'] == False].index
dfdp2.drop(falseKeep, inplace=True)
dfdp2 = dfdp2.reset_index(drop=True)
dfdp2


# In[68]:


dfmd = dfdp2[["Constituent ID", "Name"]]
dfmd["x (#donations)"] = dfdp2["adj. num gift"]-1
dfmd['t_x (last gift)'] = (dfdp2['Gift Date'] - dfdp2['First Gift Date'])/365


# In[69]:


tempDate = endDate+1
dfmd['T (total time span)'] = (tempDate - dfdp2['First Gift Date'])/365
dfmd


# In[201]:


df_BGNBD = dfmd.copy(deep=True)
donationNum = df_BGNBD['x (#donations)'].tolist()
lastGift = df_BGNBD['t_x (last gift)'].tolist()
T = df_BGNBD['T (total time span)'].tolist()


# In[202]:


A4 = []
for index, rows in islice(df_BGNBD.iterrows(), 0, None):
    if donationNum[index]>0:
        A4.append(math.log(a)-math.log(b+donationNum[index]-1)-(r+donationNum[index])*math.log(alpha+lastGift[index]))
    else:
        A4.append(0)
df_BGNBD['ln(A_4)'] = (A4)
A_4 = df_BGNBD['ln(A_4)'].tolist()


# In[203]:


A3 = []
for index, rows in islice(df_BGNBD.iterrows(), 0, None):
    A3.append(-(r+donationNum[index])*math.log(alpha+T[index]))
df_BGNBD['ln(A_3)'] = (A3)
A_3 = df_BGNBD['ln(A_3)'].tolist()


# In[204]:


A2 = []
for index, rows in islice(df_BGNBD.iterrows(), 0, None):
    A2.append(math.lgamma(a+b)+math.lgamma(b+donationNum[index])-math.lgamma(b)-math.lgamma(a+b+donationNum[index]))
df_BGNBD['ln(A_2)'] = (A2)
A_2 = df_BGNBD['ln(A_2)'].tolist()


# In[205]:


A1 = []
for index, rows in islice(df_BGNBD.iterrows(), 0, None):
    A1.append(math.lgamma(r+donationNum[index])-math.lgamma(r)+r*math.log(alpha))
df_BGNBD['ln(A_1)'] = (A1)
A_1 = df_BGNBD['ln(A_1)'].tolist()
df_BGNBD


# In[199]:


ln = []
for index, rows in islice(df_BGNBD.iterrows(), 0, None):
    if donationNum[index]>0:
        ln.append(A_1[index]+A_2[index]+math.log(math.exp(A_3[index])+1*math.exp(A_4[index])))
    else:
        ln.append(A_1[index]+A_2[index]+math.log(math.exp(A_3[index])+0*math.exp(A_4[index])))
df_BGNBD['ln(.)'] = (ln)


# In[200]:


BGNBDcols = list(df_BGNBD.columns.values)
df_BGNBD = df_BGNBD[BGNBDcols[0:5] + [BGNBDcols[-1]] + [BGNBDcols[-2]] + [BGNBDcols[-3]] + [BGNBDcols[-4]] + [BGNBDcols[-5]]]
df_BGNBD


# In[169]:


df_ext = pd.DataFrame()
tuntil = (predict_thru - beginning_date + 1)/365
print(tuntil)
n = 1/365
t = [n]
while n <= tuntil:
    n = n + (1/365)
    t.append(n)
df_ext['t'] = t
df_ext['E(X(t))'] = 0
df_ext['2F1'] = 0
df_ext


# In[170]:


z=[]
for index, rows in islice(df_ext.iterrows(), 0, None):
    z.append(df_ext.iloc[index,0] / (alpha + df_ext.iloc[index,0]))
df_ext['z']=z
df_ext['0'] = 1
df_ext


# In[171]:


count = 1
ext_cols = []
extCond = True
while (extCond):
    for index, rows in islice(df_ext.iterrows(), 0, None):
        ext_cols.append(df_ext.iloc[index, count+3]*df_ext.iloc[index, 3]*(r+count-1)*(b+count-1)/((a+b-1+count-1)*count))
    if (ext_cols[-1] < 9e-11):
        extCond = False
    else:
        df_ext[str(count)] = ext_cols
        count = count + 1
        ext_cols = [];
df_ext


# In[172]:


df_save = df_ext
df_save


# In[181]:


df_ext = df_save
ext = []
df_ext['2F1'] = df_ext.iloc[:,4:-1].sum(axis=1)
extConst = (a+b-1)/(a-1)
for index, rows in islice(df_ext.iterrows(), 0, None):
    ext.append(extConst*(1-pow((alpha/(alpha+df_ext.iloc[index,0])),r) * df_ext.iloc[index,2]))
df_ext['E(X(t))'] = ext
df_ext


# In[206]:


calibLength = (endDate-beginning_date+1)/365
df_ns = pd.DataFrame()
df_ns["Constituent ID"] = df_BGNBD["Constituent ID"]
df_ns["T (total time span)"] = df_BGNBD["T (total time span)"]


# In[211]:


firstPurch = []
T = df_ns['T (total time span)'].to_list()
for index, row in islice(df_ns.iterrows(), 0, None):
    firstPurch.append(calibLength - T[index])
df_ns['time of 1st purchase'] = (firstPurch)
maxTime = df_ns['time of 1st purchase'].max()
df_ns


# In[344]:


weekUntil = (predict_thru - beginning_date)/7
df_sls_top = pd.DataFrame()
trialDon = []
numDon = []
i = 1/365
while (i<=maxTime):
    trialDon.append(i)
    numDon.append(df_ns.loc[(df_ns['time of 1st purchase'] >= i-.001) & (df_ns['time of 1st purchase'] < i+.001)].count()[0])
    i = i + 1/365
df_sls_top["Time of Trial Donation"] = trialDon
df_sls_top["Number of Donors"] = numDon
df_sls_top


# In[345]:


df_sls = pd.DataFrame()
df_sls['t'] = t
df_sls['Cum. Rpt.'] = ""
df_sls['E(X(t))'] = ext
df_sls


# In[252]:


slsCount = 0
slsCols = []
while (slsCount < df_sls_top.count()[0]):
    for index, rows in islice(df_sls.iterrows(), 0, None):
        if (df_sls.iloc[index, 0] <= df_sls_top["Time of Trial Donation"][slsCount]):
            slsCols.append(0)
        else:
            indx = 365*(df_sls.iloc[index, 0]-df_sls_top["Time of Trial Donation"][slsCount])
            slsCols.append(ext[round(indx)-1])
    df_sls[str(slsCount)] = slsCols
    slsCount = slsCount+1
    slsCols = []
df_sls


# In[269]:


df_save1 = df_sls


# In[366]:


# cumRpt = []
# cumColCnt = 0
# #for c in range(df_sls_top.count()[0]):
# print(df_sls.iloc[0, 3:-1])
# print(df_sls_top["Time of Trial Donation"])
# print(df_sls.iloc[0, 3:-1] * df_sls_top["Time of Trial Donation"])
# #for r in range (df_sls.count()[0]):
#     #print((df_sls.iloc[r, 3:-1] * df_sls_top["Time of Trial Donation"]))
#     #cumRpt.append(sum(df_sls.iloc[r, 3:-1] * df_sls_top["Time of Trial Donation"]))
# df_sls["Cum. Rpt"] = cumRpt
# df_sls
cumRpt = []
cumRptSum = 0
#print(df_sls_top["Number of Donors"])
#print(df_sls.iloc[1,3:-1].astype('float64'))
#print (df_sls_top["Number of Donors"].multiply((df_sls.iloc[0,3:-1].astype('float64')), level = 0))
for index, rows in islice(df_sls.iterrows(), 0, None):
    i = 0
    cumRptSum = 0
    #print(cumRptSum + df_sls.iloc[index, i+3])
    #print(numDon[i])
    while (i<df_sls_top.count()[0]): 
        cumRptSum = cumRptSum + df_sls.iloc[index, i+3] * numDon[i]
        i = i+1
    cumRpt.append(cumRptSum)
    #print(cumRptSum)
df_sls["Cum. Rpt."] = cumRpt
df_sls


# In[389]:


df_checkCumRpt = pd.DataFrame()
week = list(range(1, int(weekUntil)+1))
df_checkCumRpt["week"] = week
df_checkCumRpt


# In[390]:


checkDon = [cumRpt[6]]
checkWeekDon = [checkDon[0]]
weekCount = 2
while (weekCount <= weekUntil):
    checkDon.append(cumRpt[7*weekCount-1])
    checkWeekDon.append(checkDon[weekCount-1] - checkDon[weekCount-2])
    weekCount = weekCount + 1
df_checkCumRpt["Cum. Rpt. Dons."] = checkDon
df_checkCumRpt["Weekly Rpt. Dons."] = checkWeekDon
df_checkCumRpt


# In[391]:


startDate = [beginning_date]
endDate = [startDate[0] + 6]
for index, row in islice(df_checkCumRpt.iterrows(), 1, None):
    startDate.append(startDate[index-1]+7)
    endDate.append(endDate[index-1]+7)
df_checkCumRpt['week start'] = (startDate)
df_checkCumRpt['end'] = (endDate)
df_checkCumRpt


# In[392]:


t = [0]
t2 = [6/365]
for index, row in islice(df_checkCumRpt.iterrows(), 1, None):
    t.append(t[index-1]+7/365)
    t2.append(t2[index-1]+7/365)
df_checkCumRpt["num rpt don"] = ""
df_checkCumRpt["cumul."] = ""
df_checkCumRpt['t'] = (t)
df_checkCumRpt['t2'] = (t2)
df_checkCumRpt


# In[394]:


numRpt = []
cummul = []
weekStart = df_checkCumRpt['week start'].tolist()
weekEnd = df_checkCumRpt['end'].tolist()
t = df_checkCumRpt['t'].tolist()
t2 = df_checkCumRpt['t2'].tolist()
count99 = 0
countTotal99 = 0
for index, row in islice(df_checkCumRpt.iterrows(), 0, None):
    count99 = count99 + len(dfraw2.loc[(dfraw2['Gift Date']>=weekStart[index]) & (dfraw2['Gift Date']<=weekEnd[index])])
    count99 = count99 - len(df_ns.loc[df_ns['time of 1st purchase']>=(t[index]-.001)])
    count99 = count99 + len(df_ns.loc[df_ns['time of 1st purchase']>(t2[index]+.001)])
    countTotal99 = countTotal99+count99
    numRpt.append(count99)
    cummul.append(countTotal99)
    count99 = 0
df_checkCumRpt["num rpt don"] = numRpt
df_checkCumRpt["cumul."] = cummul
df_checkCumRpt


# In[364]:


dfraw2.to_excel(writer, 'Raw data2', index = False)
dfraw1.to_excel(writer, 'Raw data1', index = False)
dfraw1.to_excel(writer, 'Data prep1', index = False)
dfdp2.to_excel(writer, 'Data prep2', index = False)
dfmd.to_excel(writer, 'Model data', index = False)
df_BGNBD.to_excel(writer, "BGNBD Estimation", index = False)
df_ext.to_excel(writer, "E(X(t))", index = False)
df_ns.to_excel(writer, "n_s", index=False)
df_sls.to_excel(writer, "CumRptSls", index=False)
df_checkCumRpt.to_excel(writer, "Check CumRpt", index=False)
writer.save()

