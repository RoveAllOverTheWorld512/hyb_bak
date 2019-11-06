# -*- coding: utf-8 -*-
"""
本程序用通达信数据对股价前复权，将数据保存为excel文件

通达信本地数据格式：
每32个字节为一个5分钟数据，每字段内低字节在前
00 ~ 01 字节：日期，整型，设其值为num，则日期计算方法为：
                        year=floor(num/2048)+2004;
                        month=floor(mod(num,2048)/100);
                        day=mod(mod(num,2048),100);
02 ~ 03 字节： 从0点开始至目前的分钟数，整型
04 ~ 07 字节：开盘价*100，整型
08 ~ 11 字节：最高价*100，整型
12 ~ 15 字节：最低价*100，整型
16 ~ 19 字节：收盘价*100，整型
20 ~ 23 字节：成交额*100，float型
24 ~ 27 字节：成交量（股），整型
28 ~ 31 字节：（保留）

每32个字节为一天数据
每4个字节为一个字段，每个字段内低字节在前
00 ~ 03 字节：年月日, 整型
04 ~ 07 字节：开盘价*100， 整型
08 ~ 11 字节：最高价*100,  整型
12 ~ 15 字节：最低价*100,  整型
16 ~ 19 字节：收盘价*100,  整型
20 ~ 23 字节：成交额（元），float型
24 ~ 27 字节：成交量（股），整型
28 ~ 31 字节：（保留）

读取需要加载struct模块，unpack之后得到一个元组。
日线读取：
fn=r"code.day";
fid=open(fn,"rb");
list=fid.read(32)
ulist=struct.unpack("iiiiifii", list)
5分钟线读取也是一样。

"""

import os
import sys
import struct
import datetime
import numpy as np
import pandas as pd
import winreg
import sqlite3

def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    '''
    GPDM股票代码
    RQ交易日期
    OPEN开盘价
    HIGH最高价
    LOW最低价
    CLOSE收盘价
    AMOUT成交额（亿元）
    VOLUME成交量（万股）
    RATE涨跌幅
    PRE_CLOSE前收盘价
    ADJ_RATE前复权调整涨跌幅
    ADJ_CLOS前复权E调整收盘价
    
    GPDM与RQ构成为唯一索引
    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS GJ
           (GPDM TEXT NOT NULL,
            RQ TEXT NOT NULL,
            OPEN REAL NOT NULL DEFAULT (0.00),
            HIGH REAL NOT NULL DEFAULT (0.00),
            LOW REAL NOT NULL DEFAULT (0.00),
            CLOSE REAL NOT NULL DEFAULT (0.00),
            AMOUT REAL NOT NULL DEFAULT (0.00),
            VOLUME REAL NOT NULL DEFAULT (0.00),
            RATE REAL NOT NULL DEFAULT (0.00),
            PRE_CLOSE REAL NOT NULL DEFAULT (0.00),
            ADJ_RATE REAL NOT NULL DEFAULT (0.00),
            ADJ_CLOSE REAL NOT NULL DEFAULT (0.00));''')
    
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ_GJ ON GJ(GPDM,RQ);''')

##########################################################################
#将字符串转换为时间戳，不成功返回None
##########################################################################
def str2datetime(s):
    try:
        dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
    except :
        dt = None
    return dt


###############################################################################
#将通达信.day读入pands
###############################################################################
def day2pd(dayfn,start=None,end=None):
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
    start = str2datetime(start)
    end = str2datetime(end)


    if start == None or end==None :
        return df
    else :
        return df[start:end]

########################################################################
#获取本机通达信安装目录，生成自定义板块保存目录
########################################################################
def gettdxdir():

    try :
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人生")
        value, type = winreg.QueryValueEx(key, "InstallLocation")
    except :
        print("本机未安装【华西证券华彩人生】软件系统。")
        sys.exit()
    return value

###############################################################################
#将通达信.day读入pands
###############################################################################
def tdxday2pd(gpdm,start=None,end=None):
    
    gpdm=sgpdm(gpdm)

    sc = 'sh' if gpdm[0]=='6' else 'sz'
#    dayfn =getdisk()+'\\tdx\\'+sc+'lday\\'+sc+gpdm+'.day'
    dayfn =gettdxdir()+'\\vipdoc\\'+sc+'\\lday\\'+sc+gpdm+'.day'

    if os.path.exists(dayfn) :
        return day2pd(dayfn,start,end)
    else :
        return []

###############################################################################
#将分红数据读入pands
###############################################################################
def fhsql2pd(gpdm):
    
    gpdm=lgpdm(gpdm)
    conn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    sql="select rq as gqdjr,fh,szg from fh where gpdm=='" +gpdm+"';"
    df=pd.read_sql_query(sql, con=conn)
    conn.close()

    df['gqdjr']=df['gqdjr'].map(lambda x:x.replace('-',''))    
    df['date']=df['gqdjr'].map(str2datetime)

    return df.set_index('date')

###############################################################################
#将配股数据读入pands
###############################################################################
def pgsql2pd(gpdm):
    
    gpdm=lgpdm(gpdm)
    conn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    sql="select rq as gqdjr,pgj,pgbl from pg where gpdm=='" +gpdm+"';"
    df=pd.read_sql_query(sql, con=conn)
    conn.close()

    df['gqdjr']=df['gqdjr'].map(lambda x:x.replace('-',''))    
    df['date']=df['gqdjr'].map(str2datetime)

    return df.set_index('date')

###############################################################################
#合并分红配股数据，并按股权登记日降序排列
###############################################################################
def getfhpg(gpdm):
    
    gpdm=lgpdm(gpdm)
    fh=fhsql2pd(gpdm)
    pg=pgsql2pd(gpdm)
    fhpg=pd.merge(fh,pg,how='outer',on='gqdjr')
    fhpg=fhpg.sort_values(by='gqdjr', ascending=False)
    return fhpg.fillna(0)


###############################################################################
#长股票代码
###############################################################################
def lgpdm(dm):
    return dm[:6]+('.SH' if dm[0]=='6' else '.SZ')

###############################################################################
#短股票代码
###############################################################################
def sgpdm(dm):
    return dm[:6]

################################################################################
#提取DataFrame时间索引，返回日期
################################################################################
def df_timeindex_to_datelist(df):
    dfti = df.index
    dftia = np.vectorize(lambda s: s.strftime('%Y%m%d'))(dfti.to_pydatetime())
    return dftia.tolist()

###############################################################################
#分红日期为股权登记日前复权收盘价
###############################################################################
def adj_close(df,fhpg):
    for i in range(len(fhpg)):
        date, fh, szg, pgj, pgbl = fhpg.iloc[i]
#        date=nextdtstr(date,-1) #如果是除权除息日则将除权基准日推前一天变为股权登记日

        fqyes = False     #如果股权登记日不在数据范围内则不能进行复权处理

        if len(df.loc[date:date])==1 :      #股权登记日存在交易
            fqyes = True
        else :
            date = df_next_date(df,date,-1) #股权登记日不存在交易则前找交易日
            if len(df.loc[date:date])==1 :   #股权登记日前有交易则进行复权
                fqyes = True

        if fqyes :
            oldclose = df.loc[date,'adj_close']
            newclose = (oldclose - fh + pgj*pgbl)/(1+szg+pgbl)
            newclose = round(newclose,2)    #四舍五入，保留2位小数
            df.loc[date,'adj_close'] = newclose
            nextdate = df_next_date(df,date,1)
            if nextdate == None :
                break
            df.loc[nextdate,'pre_close'] = newclose
            df.loc[nextdate,'adj_rate'] =  df.loc[nextdate,'close']/df.loc[nextdate,'pre_close']- 1

    ti = df_timeindex_to_datelist(df)
    ti.reverse()
    for i in range(len(ti)):
        date = ti[i]
        if i== 0 :
            df.loc[date,'adj_close'] = df.loc[date,'adj_close']
        else :
            df.loc[date,'adj_close'] = next_close /(1+next_rate)

        next_close = df.loc[date,'adj_close']
        next_rate = df.loc[date,'adj_rate']

    return df

################################################################################
#提取DataFrame时间索引指定日期date前n个日期，返回日期
################################################################################
def df_next_date(df,date,n=0):
    dftilst = df_timeindex_to_datelist(df)
    dftilst.sort()
    tmin = str2datetime(dftilst[0])
    tmax = str2datetime(dftilst[len(dftilst)-1])
    t = str2datetime(date)
    if t< tmin or t>tmax :
        return None

    try :
        i = dftilst.index(date)
        if i+n<0 or i+n>=len(dftilst) :
            return None
        else :
            return dftilst[dftilst.index(date)+n]
    except :
        while True :
            date = nextdtstr(date,-1)
            if date in dftilst :
                return date

##########################################################################
#n天后日期串，不成功返回None
##########################################################################
def nextdtstr(s,n):
    dt = str2datetime(s)
    if dt :
        dt += datetime.timedelta(n)
        return dt.strftime("%Y%m%d")
    else :
        return None

###############################################################################
# 复权股价存入Sqlite3
###############################################################################
def qgjfq(gpdm):
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    
    gj=tdxday2pd(gpdm)
    
    fhpg = getfhpg(gpdm)
    
    if len(fhpg)>0 :
        fqgj=adj_close(gj,fhpg)
        
    fqgj['gpdm']=lgpdm(gpdm)
    fqgj=fqgj.loc[:,['gpdm','date','open','high','low','close','amout','volume',
                     'rate','pre_close','adj_rate','adj_close']]

    data=np.array(fqgj).tolist()
    
    dbcn.executemany('''INSERT OR REPLACE INTO GJ (GPDM,RQ,OPEN,HIGH,LOW,CLOSE,
                     AMOUT,VOLUME,RATE,PRE_CLOSE,ADJ_RATE,ADJ_CLOSE) 
                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''', data)
    dbcn.commit()
    dbcn.close()


if __name__ == '__main__': 
    createDataBase()
    dm='002673'
    

    