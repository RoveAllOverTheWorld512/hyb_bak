# -*- coding: utf-8 -*-
"""
从大智慧F10提取股东户数导入Sqlite数据库
"""
from pyquery import PyQuery as pq
import datetime
import sqlite3
import sys
import re
import pandas as pd
import winreg

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
#从大智慧网F10获取股东户数
########################################################################
def get_gdhs(gpdm):

    sc=scdm(gpdm)
    gpdm=sgpdm(gpdm)
    
    data=[]
    url = 'http://webf10.gw.com.cn/'+sc+'/B10/'+sc+gpdm+'_B10.html'

    try :
        html = pq(url,encoding="utf-8")
        #第3个区块
        #sect = pq(html('section').eq(2).html())
        #提取预测明细
        sect=html('section').filter('#股东人数').html()
        tr=pq(sect)
    except : 
        print("出错退出")
        return data

    for i in range(1,len(tr('ul'))):
        
        il=tr('ul').eq(i).text().split(' ')
        rq=il[0]
        gdhs=il[1]

        data.append([lgpdm(gpdm),rq,gdhs])

    return data
    
########################################################################
#从大智慧网F10获取限售解禁
########################################################################
def get_xsjj(gpdm):

    sc=scdm(gpdm)
    gpdm=sgpdm(gpdm)
    
    data=[]
    url = 'http://webf10.gw.com.cn/'+sc+'/B11/'+sc+gpdm+'_B11.html'

    try :
        html = pq(url,encoding="utf-8")
        #第3个区块
        #sect = pq(html('section').eq(2).html())
        #提取预测明细
        sect=html('section').filter('#解禁流通').html()
                 
        tbl=pq(sect)
        tr=pq(tbl('table').eq(1).html())
        
    except : 
        print("出错退出")
        return data

    for i in range(len(tr('tr'))):
        
        tds=pq(tr('tr').eq(i))
        td=[]
        for j in range(len(tds('td'))):
            td.append(tds('td').eq(j).text())
            
        
        jjrq=td[0].replace('/','-')
        bcjj=round(float(td[1])/10000,4)
        wlt=float(td[6])/10000
        try:
            qltbl=float(td[3].replace('%',''))
            qlt=round(bcjj/qltbl*100,4)
            hlt=round(qlt+bcjj,4)
            hltbl=round(bcjj/hlt*100,4)
        except:            
            qltbl=None
            qlt=None
            hlt=None
            hltbl=None
            
        data.append([lgpdm(gpdm),jjrq,bcjj,qlt,qltbl,hlt,hltbl,None,None,wlt])

    return data
    
'''
CREATE TABLE [XSJJ](
  [GPDM] TEXT NOT NULL, 
  [JJRQ] TEXT NOT NULL, 
  [JJSL] REAL NOT NULL, 
  [QLTGB] REAL, 
  [QLTBL] REAL, 
  [HLTGB] REAL, 
  [HLTBL] REAL, 
  [QZD] REAL, 
  [HZD] REAL, 
  [WLT] REAL);

CREATE UNIQUE INDEX [GPDM_JJRQ_XSJJ]
ON [XSJJ](
  [GPDM], 
  [JJRQ]);

'''    

    
if __name__ == "__main__":  
#def temp():
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    gpdmb=get_gpdm()
    
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)
    j=0
    for i in range(j,len(gpdmb)):
        gpdm=gpdmb.index[i]
        gpmc = gpdmb.iloc[i]['gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(gpdmb),i+1,gpdm,gpmc)) 
        data = get_xsjj(gpdm)
        
        if len(data)>0 :
            dbcn.executemany('''INSERT OR REPLACE INTO XSJJ_DZH (GPDM,JJRQ,JJSL,QLTGB,QLTBL,HLTGB,HLTBL,QZD,HZD,WLT)
            VALUES (?,?,?,?,?,?,?,?,?,?)''', data)

        if (i % 10 ==0) or i==len(gpdmb) :
            dbcn.commit()

    dbcn.close()


    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

'''
python使用pyquery库总结 
https://blog.csdn.net/baidu_21833433/article/details/70313839

'''

