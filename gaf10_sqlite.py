# -*- coding: utf-8 -*-
"""
本程序从港澳资讯http://www.gaf10.com网提取股东户数、历年股本变动、历年分红扩股数据导入Sqlite数据库
"""
from pyquery import PyQuery as pq
import pandas as pd
import datetime
import time
import sqlite3
import re
from urllib.error import HTTPError

########################################################################
#建立数据库
########################################################################
def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')

    cn.execute('''CREATE TABLE IF NOT EXISTS GDHS
           (GPDM TEXT NOT NULL,
           RQ TEXT NOT NULL,
           GDHS INTEGER NOT NULL);''')
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ ON GDHS(GPDM,RQ);''')

    cn.execute('''CREATE TABLE IF NOT EXISTS LNGBBD
           (GPDM TEXT NOT NULL,
           RQ TEXT NOT NULL,
           ZGB REAL NOT NULL,
           LTGB REAL NOT NULL,
           SJLTGB REAL);''')
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ ON LNGBBD(GPDM,RQ);''')

    cn.execute('''CREATE TABLE IF NOT EXISTS LNFHKG
           (GPDM TEXT NOT NULL,
           RQ TEXT NOT NULL,
           FH REAL NOT NULL DEFAULT 0.00,
           SZG REAL NOT NULL DEFAULT 0.00,
           FHKG TEXT NOT NULL,
           BJ TEXT);''')
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ ON LNFHKG(GPDM,RQ);''')

###############################################################################
#从通达信系统读取股票代码表
###############################################################################
def get_gpdm():
    datacode = []
    for sc in ('h','z'):
        fn = r'D:\new_hxzq_hc\T0002\hq_cache\s'+sc+'m.tnf'
        f = open(fn,'rb')
        f.seek(50)
        ss = f.read(314)
        while len(ss)>0:
            gpdm=ss[0:6].decode('GBK')
            gpmc=ss[23:31].strip(b'\x00').decode('GBK').replace(' ','').replace('*','')
            gppy=ss[285:289].strip(b'\x00').decode('GBK')
            #剔除非A股代码
            if (sc=="h" and gpdm[0]=='6') :
                gpdm=gpdm+'.SH'
                datacode.append([gpdm,gpmc,gppy])
            if (sc=='z' and (gpdm[0]=='0' or gpdm[0:2]=='30')) :
                gpdm=gpdm+'.SZ'
                datacode.append([gpdm,gpmc,gppy])
            ss = f.read(314)
        f.close()
    gpdmb=pd.DataFrame(datacode,columns=['gpdm','gpmc','gppy'])
    return gpdmb


########################################################################
#检测是不是可以转换成整数
########################################################################
def str2int(num):
    try:
        return int(num)
    except :
        return num

########################################################################
#检测是不是可以转换成整数
########################################################################
def strisint(num):
    try:
        num = int(num)
        return True
    except :
        return False

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def str2float(num):
    try:
        return float(num)
    except :
        return num

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def strisfloat(num):
    try:
        num = float(num)
        return True
    except :
        return False

##########################################################################
#判断字符串日期"2017-01-01"是否有效
##########################################################################
def isVaildDate(date):
    try:
        time.strptime(date, "%Y-%m-%d")
        return True
    except:
        return False
########################################################################
#提取股东户数
########################################################################
def gdhs(gpdm):
    
    dm=gpdm[0:6]
    data = []

    url  = "http://web-f10.gaotime.com/stock/"+dm+"/gdyj/gdhs.html"

    try :
        html = pq(url,encoding="utf-8")
    except HTTPError as e: 
        print("出错退出")
        return data
    #第2个表
    tb = pq(html('table').eq(1).html())
    #行数
    tr = tb('tr')
    if len(tr)>2:
        for i in range(1,len(tr)) :
            row=pq(tr.eq(i).html())
            #前两列
            rq=row.find('td').eq(0).text()
            hs=str2int(row.find('td').eq(1).text())
            if isVaildDate(rq):
                rowdat=[gpdm,rq,hs]
    
            data.append(rowdat)

    return data       

########################################################################
#提取历年股本变动数据
########################################################################
def lngbbd(gpdm): 
    '''
    CREATE TABLE [LNGBBD](
      [GPDM] TEXT NOT NULL, 
      [RQ] TEXT NOT NULL, 
      [ZGB] REAL NOT NULL, 
      [LTGB] REAL NOT NULL, 
      [SJLTGB] REAL);
    
    CREATE INDEX [GPDM_RQ_LNGBBD]
    ON [LNGBBD](
      [GPDM], 
      [RQ]);
    
    
    '''
    
    dm=gpdm[0:6]
    
    data=[]

    url  = "http://web-f10.gaotime.com/stock/"+dm+"/gbjg/lngbbd.html"

    try :
        html = pq(url,encoding="utf-8")
    except HTTPError as e: 
        print("出错退出")
        return data

    tb = html('tr')
    #该页面在一个表中有一个表头、两个表体，提取时去掉表头行和最后一个表体空行
    for i in range(1,len(tb)-1) :
        row=pq(tb.eq(i).html())
        rq = row.find('td').eq(0).text()
        zgb = row.find('td').eq(1).text()
        ltgb = row.find('td').eq(2).text()
        sjltgb = row.find('td').eq(3).text()
        if isVaildDate(rq):
            rowdat=[gpdm,rq,zgb,ltgb,sjltgb]
            rowdat=[e if e!='-' else None for e in rowdat]
            data.append(rowdat)
 
    return data       

########################################################################
#分析分红扩股数据
########################################################################
def fxfhkg(fhkg):

    fhstr=fhkg.replace('股','').replace('元','').replace(" ","")
    i=fhstr.find('(含税)')
    if i!=-1 :
        fhstr=fhstr[0:i]
              
    fh = 0     #每股分红
    sg = 0     #每股送股和转增股数
    fas = 0    #方案数
    bj = "1"

    fhs = re.findall ('([\d\.]+)派([\d\.]+)',fhstr)
    if len(fhs) >1 :
        print(fhstr)
    elif len(fhs) == 1:
        fh = float(fhs[0][1])/float(fhs[0][0])
        sg = 0
        fas += 1
        bj = ""

    fhs = re.findall ('([\d\.]+)送([\d\.]+)',fhstr )
    if len(fhs) >1 :
        print(fhstr)
    elif len(fhs) == 1:
        fh = 0
        sg = float(fhs[0][1])/float(fhs[0][0])
        fas += 1
        bj = ""

    fhs = re.findall ('([\d\.]+)转([\d\.]+)',fhstr )
    if len(fhs) >1 :
        print(fhstr)
    elif len(fhs) == 1:
        fh = 0
        sg = float(fhs[0][1])/float(fhs[0][0])
        fas += 1
        bj = ""

    fhs = re.findall ('([\d\.]+)送([\d\.]+)派([\d\.]+)',fhstr )
    if len(fhs) >1 :
        print(fhstr)
    elif len(fhs) == 1:
        fh = float(fhs[0][2])/float(fhs[0][0])
        sg = float(fhs[0][1])/float(fhs[0][0])
        fas += 1
        bj = ""

    fhs = re.findall ('([\d\.]+)转([\d\.]+)派([\d\.]+)',fhstr )
    if len(fhs) >1 :
        print(fhstr)
    elif len(fhs) == 1:
        fh = float(fhs[0][2])/float(fhs[0][0])
        sg = float(fhs[0][1])/float(fhs[0][0])
        fas += 1
        bj = ""

    fhs = re.findall ('([\d\.]+)送([\d\.]+)转([\d\.]+)',fhstr )
    if len(fhs) >1 :
        print(fhstr)
    elif len(fhs) == 1:
        fh = 0
        sg = (float(fhs[0][1])+float(fhs[0][2]))/float(fhs[0][0])
        fas += 1
        bj = ""

    fhs = re.findall ('([\d\.]+)转([\d\.]+)送([\d\.]+)',fhstr )
    if len(fhs) >1 :
        print(fhstr)
    elif len(fhs) == 1:
        fh = 0
        sg = (float(fhs[0][1])+float(fhs[0][2]))/float(fhs[0][0])
        fas += 1
        bj = ""

    fhs = re.findall ('([\d\.]+)送([\d\.]+)转([\d\.]+)派([\d\.]+)',fhstr )
    if len(fhs) >1 :
        print(fhstr)
    elif len(fhs) == 1:
        fh = float(fhs[0][3])/float(fhs[0][0])
        sg = (float(fhs[0][1])+float(fhs[0][2]))/float(fhs[0][0])
        fas += 1
        bj = ""

    fhs = re.findall ('([\d\.]+)转([\d\.]+)送([\d\.]+)派([\d\.]+)',fhstr )
    if len(fhs) >1 :
        print(fhstr)
    elif len(fhs) == 1:
        fh = float(fhs[0][3])/float(fhs[0][0])
        sg = (float(fhs[0][1])+float(fhs[0][2]))/float(fhs[0][0])
        fas += 1
        bj = ""

    if fas>1 :
        bj="1"

    if (fh==0 and sg==0) :
        print(fhstr)
        bj = "2"

    return [fh,sg,fhkg,bj]

########################################################################
#提取历年分红扩股数据
########################################################################
def lnfhkg(gpdm): 
    
    dm=gpdm[0:6]

    data=[]

    url  = "http://web-f10.gaotime.com/stock/"+dm+"/fhkg/lnfhkg.html"

    try :
        html = pq(url,encoding="utf-8")
    except HTTPError as e: 
        print("出错退出")
        return data
    
    #该页面有3个表，分红数据在第1个表，表的前2行为表头
    tb = pq(html('table').eq(0).html())
    #行数
    tr = tb('tr')
    for i in range(2,len(tr)) :
        row=pq(tr.eq(i).html())

        if (row.find('td').eq(1).text()=='是') and (row.find('td').eq(7).text()!='-'):
            
            rq=row.find('td').eq(6).text()  #股权登记日
            fhkg=row.find('td').eq(3).text()

            if isVaildDate(rq):
                fhsg=fxfhkg(fhkg)
                rowdat=[gpdm,rq]
                rowdat.extend(fhsg)

            data.append(rowdat)

    return data       

    
if __name__ == "__main__":  

    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    gpdmb=get_gpdm()
    gpdmb=gpdmb.set_index('gpdm')
#    gpdmb=gpdmb[gpdmb.index.isin(['600007.SH'])]
#    createDataBase()
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')

    for i in range(len(gpdmb)):
        gpdm = gpdmb.index[i]
        gpmc = gpdmb.loc[gpdm,'gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(gpdmb),i+1,gpdm,gpmc)) 

#        data = gdhs(gpdm)
#        dbcn.executemany('INSERT OR IGNORE INTO GDHS (GPDM,RQ,GDHS) VALUES (?,?,?)', data)

        data = lngbbd(gpdm)
        dbcn.executemany('INSERT OR IGNORE INTO LNGBBD (GPDM,RQ,ZGB,LTGB,SJLTGB) VALUES (?,?,?,?,?)', data)

#        data = lnfhkg(gpdm)
#        dbcn.executemany('INSERT OR IGNORE INTO LNFHKG (GPDM,RQ,FH,SZG,FHKG,BJ) VALUES (?,?,?,?,?,?)', data)

#    dbcn.execute("update gdhs set gdhs=19943 where rq='2017-09-30' and gpdm='600302.SH';")
    dbcn.commit()
    dbcn.close()

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

