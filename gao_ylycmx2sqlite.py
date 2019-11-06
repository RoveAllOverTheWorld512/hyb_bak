# -*- coding: utf-8 -*-
"""
从港澳资讯网提取盈利预测明细导入Sqlite数据库
"""
from pyquery import PyQuery as pq
import datetime
import sqlite3
import sys
import re
import pandas as pd
import winreg

def jgdic():
    return {
        '高盛':'高盛高华',
        '国泰君安国际':'国泰君安',
        '群益证券(香港)':'群益证券',
        '申万宏源研究':'申万宏源',
        '新时代证券':'新时代',
        '银河国际':'银河证券',
        '银河国际(香港)':'银河证券',
        '元大证券(香港)':'元大证券',
        '元大证券股份有限公司':'元大证券',
        '中国银河':'银河证券',
        '中国银河国际':'银河证券',
        '中国银河国际证券':'银河证券',
        '中信建投(国际)':'中信建投',
        '中信建投证券':'中信建投'
        }


########################################################################
#建立数据库
########################################################################
def createDataBase():
    dbfn=getdrive()+'\\hyb\\STOCKEPS.db'
    cn = sqlite3.connect(dbfn)
    '''
    股票代码,日期(发布日期)，机构名称，预测年份，EPS
    '''
    try :
        cn.execute('''CREATE TABLE IF NOT EXISTS GAORPT
               (GPDM TEXT NOT NULL,
               RQ TEXT NOT NULL,
               PJJG TEXT NOT NULL,
               NF TEXT,
               EPS REAL
               );''')
        cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GAORPT_GPDM_RQ_PJJG ON GAORPT(GPDM,RQ,PJJG);''')

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
#从港澳资讯网F10获取盈利预测明细
########################################################################
def get_ylycmx(gpdm):
    
    pjjgdic=jgdic()
    gpdm=sgpdm(gpdm)
    
    data=[]

    url = 'http://web-f10.gaotime.com/stock/'+gpdm+'/jzfx/jgycmx.html'

    try :
        tr = pq(url,encoding="utf-8")

    except : 
        print("出错退出")
        return data
        
    rs=len(tr('tr'))
    if rs>4 :   
        for i in range(2,len(tr('tr'))-1):
            #分析行
            td=tr('tr').eq(i).text().split(' ')
    
            rq=td[0]
            nf=td[1][:4]
            pjjg=td[2]
            if pjjg in pjjgdic.keys():
                pjjg=pjjgdic[pjjg]
            
            eps=td[4]
            if eps != '-' :
                data.append([lgpdm(gpdm),rq,pjjg,nf,eps])

    return data
    

    
if __name__ == "__main__":  
#def temp():
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    
    createDataBase()
    gpdmb=get_gpdm()
    
    dbfn=getdrive()+'\\hyb\\STOCKEPS.db'
    dbcn = sqlite3.connect(dbfn)
#    curs = dbcn.cursor()
#
#    curs.execute('''select distinct gpdm from yygc;''')
#    
#    data = curs.fetchall()
#    data = [e[0] for e in data]    
#    gpdmb=gpdmb[~gpdmb.index.isin(data)]
    
    for i in range(0,len(gpdmb)):
        gpdm=gpdmb.index[i]
        gpmc = gpdmb.iloc[i]['gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(gpdmb),i+1,gpdm,gpmc)) 
        data = get_ylycmx(gpdm)
        
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO GAORPT (GPDM,RQ,PJJG,NF,EPS) VALUES (?,?,?,?,?)', data)

        #每100个股票提交一次    
        if i%50==0 :           
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

