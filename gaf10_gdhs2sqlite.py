# -*- coding: utf-8 -*-
"""
提取股东户数导入Sqlite数据库
"""
from pyquery import PyQuery as pq
import pandas as pd
import datetime
import sqlite3
import re
from urllib.error import HTTPError

########################################################################
#建立数据库
########################################################################
def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')

    cn.execute('''CREATE TABLE IF NOT EXISTS GDHS
           (ID integer PRIMARY KEY AUTOINCREMENT,
           CODE TEXT,
           DATE TEXT,
           GDHS INTEGER);''')

    cn.execute('''CREATE TABLE IF NOT EXISTS LNGBBD
           (ID integer PRIMARY KEY AUTOINCREMENT,
           CODE TEXT,
           DATE TEXT,
           ZGB REAL,
           LTGB REAL,
           SJLTGB REAL);''')

    cn.execute('''CREATE TABLE IF NOT EXISTS LNFHKG
           (ID integer PRIMARY KEY AUTOINCREMENT,
           CODE TEXT,
           DATE TEXT,
           FH REAL,
           SZG REAL,
           FHKG TEXT,
           BJ TEXT);''')

###############################################################################
#从通达信系统读取股票代码表
###############################################################################
def getcode():
    datacode = []
    for sc in ('h','z'):
        fn = r'C:\new_hxzq_hc\T0002\hq_cache\s'+sc+'m.tnf'
        f = open(fn,'rb')
        f.seek(50)
        ss = f.read(314)
        while len(ss)>0:
            gpdm=ss[0:6].decode('GBK')
            gpmc=ss[23:31].strip(b'\x00').decode('GBK').replace(' ','').replace('*','')
            gppy=ss[285:289].strip(b'\x00').decode('GBK')
            #剔除非A股代码
            if (sc=="h" and gpdm[0]=='6') or (sc=='z' and (gpdm[0]=='0' or gpdm[0:2]=='30')) :
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
    except ValueError:
        return num


########################################################################
#检测是不是可以转换成浮点数
########################################################################
def str2float(num):
    try:
        return float(num)
    except ValueError:
        return num


########################################################################
#提取股东户数
########################################################################
def gdhs(gpdm):

    data = []

    url  = "http://web-f10.gaotime.com/stock/"+gpdm+"/gdyj/gdhs.html"

    html = pq(url,encoding="utf-8")
    #第2个表
    tb = pq(html('table').eq(1).html())
    #行数
    tr = tb('tr')
    for i in range(1,len(tr)-1) :
        row=pq(tr.eq(i).html())
        rowdat=[gpdm]
        #前两列
        for j in range(0,2):
            col=row.find('td').eq(j).text()
            rowdat.append(str2int(col))

        data.append(rowdat)

    return data       

########################################################################
#提取历年股本变动数据
########################################################################
def lngbbd(gpdm): 
    
    data=[]

    url  = "http://web-f10.gaotime.com/stock/"+gpdm+"/gbjg/lngbbd.html"

    html = pq(url,encoding="utf-8")

    tb = html('tr')

    for i in range(1,len(tb)-1) :
        row=pq(tb.eq(i).html())
        rowdat=[gpdm]
        for j in range(0,4):
            col=row.find('td').eq(j).text()
            rowdat.append(str2float(col))

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

    data=[]

    url  = "http://web-f10.gaotime.com/stock/"+gpdm+"/fhkg/lnfhkg.html"

    try :
        html = pq(url,encoding="utf-8")
    except HTTPError as e: 
        print("出错退出")
        return data
    
    #第2个表
    tb = pq(html('table').eq(0).html())
    #行数
    tr = tb('tr')
    for i in range(2,len(tr)) :
        row=pq(tr.eq(i).html())

        if (row.find('td').eq(1).text()=='是') and (row.find('td').eq(7).text()!='-'):
            rq=row.find('td').eq(7).text()
            fhkg=row.find('td').eq(3).text()

            fhsg=fxfhkg(fhkg)

            rowdat=[gpdm,rq]
            rowdat.extend(fhsg)

            data.append(rowdat)

    return data       

    
if __name__ == "__main__":  

    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    gpdmb=getcode()
    
    createDataBase()
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')

    for i in range(0,len(gpdmb)):
        gpdm = gpdmb.loc[i,'gpdm']
        gpmc = gpdmb.loc[i,'gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(gpdmb),i+1,gpdm,gpmc)) 
        data = lnfhkg(gpdm)
#        print(data) 
#        data = gdhs(gpdm)
#        data = lngbbd(gpdm)
#        dbcn.executemany('INSERT INTO LNGBBD (CODE,DATE,ZGB,LTGB,SJLTGB) VALUES (?,?,?,?,?)', data)
        dbcn.executemany('INSERT INTO LNFHKG (CODE,DATE,FH,SZG,FHKG,BJ) VALUES (?,?,?,?,?,?)', data)

    dbcn.commit()
    dbcn.close()

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)
