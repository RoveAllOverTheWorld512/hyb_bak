# -*- coding: utf-8 -*-
"""
Created on Tue Dec  5 16:46:47 2017

@author: lenovo
"""
import struct
import pandas as pd
import datetime

##########################################################################
#将字符串转换为时间戳，不成功返回None
##########################################################################
def str2datetime(s):
    try:
        dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
    except :
        dt = None
    return dt


dayfn=r'C:\new_hxzq_hc\vipdoc\sh\lday\sh600356.day'
columns = ['rq','date','open', 'high', 'low','close','amout','volume','rate','pre_close','adj_rate','adj_close']

with open(dayfn,"rb") as f:
    data = f.read()
    f.close()
days = int(len(data)/32)
records = []
qsp = 0
for i in range(days):
    dat = data[i*32:(i+1)*32]
    rq,kp,zg,zd,sp,cje,cjl,tmp = struct.unpack("iiiiifii", dat)
    rq1 = str2datetime(str(rq))
    rq2 = rq1.strftime("%Y-%m-%d")
    kp = kp/100.00
    zg = zg/100.00
    zd = zd/100.00
    sp = sp/100.00
    cje = cje/100000000.00     #亿元
    cjl = cjl/10000.00         #万股
    zf = sp/qsp-1 if (i>0 and qsp>0) else 0.0
    records.append([rq1,rq2,kp,zg,zd,sp,cje,cjl,zf,qsp,zf,sp])
    qsp = sp

df = pd.DataFrame(records,columns=columns)
df = df.set_index('rq')
