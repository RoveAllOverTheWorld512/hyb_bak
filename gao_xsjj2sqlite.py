# -*- coding: utf-8 -*-
"""
从港澳资讯网F10提取限售解禁数据导入Sqlite数据库
"""
from pyquery import PyQuery as pq
import datetime
import sqlite3
import sys
import re
import pandas as pd
import winreg


########################################################################
#建立数据库
########################################################################
def createDataBase():
    dbfn=getdrive()+'\\hyb\\STOCKEPS.db'
    cn = sqlite3.connect(dbfn)
    '''
    股票代码,日期，分类（行业，产品，地区），经营业务，营业收入（万元），营业利润（万元），毛利率（%）
    '''
    try :
                
        cn.execute('''CREATE TABLE IF NOT EXISTS YYGC
               (GPDM TEXT NOT NULL,
               RQ TEXT NOT NULL,
               FL TEXT NOT NULL,
               JYYW TEXT NOT NULL,
               YYSR REAL,
               YYLR REAL,
               MLL REAL
               );''')
        cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS YYGC_GPDM_RQ_FL ON YYGC(GPDM,RQ,FL,JYYW);''')
    except:
        cn.close()

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

########################################################################
#从港澳资讯网获取限售解禁明细
########################################################################
def get_xsjjmx(gpdm):

    gpdm=sgpdm(gpdm)
    
    data=[]
    url = 'http://web-f10.gaotime.com/stock/'+gpdm+'/gbjg/xsjj.html'


    try :
        tr = pq(url,encoding="utf-8")
    except : 
        print("出错退出")
        return data


    for i in range(1,len(tr('tr'))-1):
        tds=pq(tr('tr').eq(i))

        if len(tds('td'))==5:
            jjgd=tds('td').eq(0).text()
            jjrq=tds('td').eq(1).text()
            bcjj=tds('td').eq(2).text()
 
        if len(tds('td'))==4:
            jjrq=tds('td').eq(0).text()
            bcjj=tds('td').eq(1).text()
            
        if len(tds('td')) in (4,5) and jjrq!='-':
            data.append([lgpdm(gpdm),jjrq,bcjj,jjgd])
        
    return data

    
if __name__ == "__main__":  
#def temp():
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    
    createDataBase()
    gpdmb=get_gpdm()
    
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)
    
    for i in range(len(gpdmb)):
        gpdm=gpdmb.index[i]
        gpmc = gpdmb.iloc[i]['gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(gpdmb),i+1,gpdm,gpmc)) 
        data = get_xsjjmx(gpdm)
        
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO XSJJ_GAO (GPDM,JJRQ,JJSL,JJGD) VALUES (?,?,?,?)', data)

        #每100个股票提交一次    
        if i%10==0 :           
            dbcn.commit()

    dbcn.commit()
    dbcn.close()


    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

    '''
    python使用pyquery库总结 
    https://blog.csdn.net/baidu_21833433/article/details/70313839
    
    '''
