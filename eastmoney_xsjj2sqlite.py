# -*- coding: utf-8 -*-
"""
从东方财富网提取限售解禁数据导入Sqlite数据库
"""
from pyquery import PyQuery as pq
import datetime
import sqlite3
import sys
import re
import numpy as np
import pandas as pd
import winreg
import time
from selenium import webdriver

########################################################################
#获取驱动器
########################################################################
def getdrive():
    return sys.argv[0][:2]



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

###############################################################################
#长股票代码
###############################################################################
def lgpdm(dm):
    dm=re.findall('(\d{6})',dm)
    
    if len(dm)==0 :
        return None

    dm=dm[0] 

    return dm+('.SH' if dm[0]=='6' else '.SZ')

###############################################################################
#中股票代码
###############################################################################
def mgpdm(dm):
    dm=re.findall('(\d{6})',dm)
    
    if len(dm)==0 :
        return None
    dm=dm[0]
    return ('SH' if dm[0]=='6' else 'SZ')+dm

###############################################################################
#短股票代码
###############################################################################
def sgpdm(dm):
    dm=re.findall('(\d{6})',dm)
    
    if len(dm)==0 :
        return None

    return dm[0]

###############################################################################
#市场代码
###############################################################################
def scdm(gpdm):
    dm=re.findall('(\d{6})',gpdm)
    
    if len(dm)==0 :
        return None

    dm = dm[0]
    
    return 'SH' if dm[0]=='6' else 'SZ'


###############################################################################
#市场代码
###############################################################################
def minus2none(s):
    return s if s!='-' else None


###############################################################################
#从通达信系统读取股票代码表
###############################################################################
def get_gpdm():
    datacode = []
    for sc in ('h','z'):
        fn = gettdxdir()+'\\T0002\\hq_cache\\s'+sc+'m.tnf'
        f = open(fn,'rb')
        f.seek(50)
        ss = f.read(314)
        while len(ss)>0:
            gpdm=ss[0:6].decode('GBK')
            gpmc=ss[23:31].strip(b'\x00').decode('GBK').replace(' ','').replace('*','')
            gppy=ss[285:291].strip(b'\x00').decode('GBK')
            #剔除非A股代码
            if (sc=="h" and gpdm[0]=='6') :
                gpdm=gpdm+'.SH'
                datacode.append([gpdm,gpmc,gppy])
            if (sc=='z' and (gpdm[0:2]=='00' or gpdm[0:2]=='30')) :
                gpdm=gpdm+'.SZ'
                datacode.append([gpdm,gpmc,gppy])
            ss = f.read(314)
        f.close()
    gpdmb=pd.DataFrame(datacode,columns=['gpdm','gpmc','gppy'])
    gpdmb=gpdmb.set_index('gpdm')
    return gpdmb

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
#万亿转换
###############################################################################
def wyzh(str):
    wy=re.findall('(.+)亿',str)
    if len(wy)==1 :
        return float(wy[0])*100000000
    wy=re.findall('(.+)万',str)
    if len(wy)==1 :
        return float(wy[0])*10000

    return 0

########################################################################
#从东方财富网获取限售解禁数据
########################################################################
def get_xsjj(gpdm):
     
    gpdm=sgpdm(gpdm)

    data=[]
    
    url='http://data.eastmoney.com/dxf/q/%s.html' % gpdm
    
    browser = webdriver.PhantomJS()
    browser.get(url)
    time.sleep(5)
    try:
        html = browser.find_element_by_id("td_1")      # 不要用 browser.page_source，那样得到的页面源码不标准
        dbtbl = html.find_element_by_tag_name("tbody").get_attribute("innerHTML") 
        html = pq(dbtbl)
    
        html.find("script").remove()    # 清理 <script>...</script>
        html.find("style").remove()     # 清理 <style>...</style>
        
        rows=html('tr')
    
        for i in range(len(rows)):
            row=rows.eq(i).text().split(' ')
            jjrq=row[2]
            bcjj=round(wyzh(row[4])/100000000,4)
            hlt=round(wyzh(row[5])/100000000,4)
            wlt=round(wyzh(row[6])/100000000,4)
            qlt=round(hlt-bcjj,4)
            qltbl=round(bcjj/qlt*100,4)
            hltbl=round(bcjj/hlt*100,4)
            qzd=np.nan if row[9]=='-' else float(row[9])
            hzd=np.nan if row[10]=='-' else float(row[10])
            
            data.append([lgpdm(gpdm),jjrq,bcjj,qlt,qltbl,hlt,hltbl,qzd,hzd,wlt])

    except:
        pass
    
    browser.quit()
    
    return data   

########################################################################
#从数据库中提取已有股票
########################################################################
def del_gpdm():
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)
    curs = dbcn.cursor()
    sql='select distinct gpdm from xsjj;'
    curs.execute(sql)        
    data = curs.fetchall()

    cols = ['gpdm']
    
    df=pd.DataFrame(data,columns=cols)
    df=df.set_index('gpdm')
    
    return df
    
if __name__ == "__main__":  
#def temp():
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    #股票代码表
    gpdmb=get_gpdm()
    
    #已有股票代码
    delgpdm=del_gpdm()
    #去掉已有股票代码
    gpdmb=gpdmb[~gpdmb.index.isin(delgpdm.index)]    

    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    for i in range(len(gpdmb)):
        gpdm=gpdmb.index[i]
        gpmc = gpdmb.iloc[i]['gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(gpdmb),i+1,gpdm,gpmc)) 
        data = get_xsjj(gpdm)
        
        if len(data)>0 :
            dbcn.executemany('''INSERT OR REPLACE INTO XSJJ (GPDM,JJRQ,JJSL,QLTGB,QLTBL,HLTGB,HLTBL,QZD,HZD,WLT)
            VALUES (?,?,?,?,?,?,?,?,?,?)''', data)

        if (i % 10 ==0) or i==len(gpdmb) :
            dbcn.commit()
    
    dbcn.close()


    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

#'''
#python使用pyquery库总结 
#https://blog.csdn.net/baidu_21833433/article/details/70313839
#
#'''
#'''
#CREATE TABLE [XSJJ](
#  [GPDM] TEXT NOT NULL, 
#  [JJRQ] TEXT NOT NULL, 
#  [JJSL] REAL NOT NULL, 
#  [QLTGB] REAL NOT NULL, 
#  [QLTBL] REAL NOT NULL, 
#  [HLTGB] REAL NOT NULL, 
#  [HLTBL] REAL NOT NULL, 
#  [QZD] REAL, 
#  [HZD] REAL,
#  [WLT] REAL
#  );
#
#CREATE UNIQUE INDEX [GPDM_JJRQ_XSJJ]
#ON [XSJJ](
#  [GPDM], 
#  [JJRQ]);
#'''

#if __name__ == "__main__":  
#    gpdm='600114'
#    url='http://data.eastmoney.com/dxf/q/%s.html' % gpdm
#    
#    browser = webdriver.PhantomJS()
#    browser.get(url)
#    time.sleep(5)
#
#    html = browser.find_element_by_id("td_1")      # 不要用 browser.page_source，那样得到的页面源码不标准
#    dbtbl = html.find_element_by_tag_name("tbody").get_attribute("innerHTML") 
#    html = pq(dbtbl)
#
#    html.find("script").remove()    # 清理 <script>...</script>
#    html.find("style").remove()     # 清理 <style>...</style>
#    
#    rows=html('tr')
#    data=[]
#    for i in range(len(rows)):
#        row=rows.eq(i).text().split(' ')
#        jjrq=row[2]
#        bcjj=round(wyzh(row[4])/100000000,4)
#        hlt=round(wyzh(row[5])/100000000,4)
#        qlt=round(hlt-bcjj,4)
#        qltbl=round(bcjj/qlt*100,4)
#        hltbl=round(bcjj/hlt*100,4)
#        qzd=np.nan if row[9]=='-' else float(row[9])
#        hzd=np.nan if row[10]=='-' else float(row[10])
#        
#        data.append([lgpdm(gpdm),jjrq,bcjj,qlt,qltbl,hlt,hltbl,qzd,hzd])
#        
        