# -*- coding: utf-8 -*-
"""
从大智慧提取股东进出数据导入Sqlite数据库
"""
from pyquery import PyQuery as pq
import datetime
import sqlite3
import sys
import re
import pandas as pd
import winreg
from selenium import webdriver


########################################################################
#建立数据库
########################################################################
def createDataBase():
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    cn = sqlite3.connect(dbfn)
    '''
    股东进出：
    股东代码，股东名称，股票代码,股票名称，日期，持股数量，持股变化

    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS DZHGDJC
           (
           GDDM TEXT NOT NULL,
           GDMC TEXT NOT NULL,
           GPDM TEXT NOT NULL,
           GPMC TEXT NOT NULL,
           RQ TEXT NOT NULL,
           CGSL REAL NOT NULL,
           CGBH TEXT
           );''')
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS DZHGDJC_GDDM_GPDM_RQ ON DZHGDJC(GDDM,GPDM,RQ);''')

    '''
    股东信息：
    股东代码，股东名称

    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS DZHGDDM
               (GDDM TEXT PRIMARY KEY NOT NULL, 
                GDMC TEXT NOT NULL);
                ''')

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
#从大智慧网F10获取股东代码
########################################################################
def get_gddm(gpdm):

    sc=scdm(gpdm)
    gpdm=sgpdm(gpdm)
    
    data=[]
    url = 'http://webf10.gw.com.cn/'+sc+'/B10/'+sc+gpdm+'_B10.html'

    try :
        html = pq(url,encoding="utf-8")
        #第3个区块
        #sect = pq(html('section').eq(2).html())
        #提取预测明细
        sect=html('section').filter('#十大流通股东').html()
        
        div=pq(sect)
        tbl=pq(div('div').eq(2).html())
        tr=pq(tbl('table').eq(0).html())
    except : 
        print("出错退出")
        return data

    trs=tr('tr')
    
    for i in range(len(trs)):
        row=pq(trs.eq(i).html())
        td=row('td').eq(0).html()

        if not td is None:
            gddm=re.findall('stockgddm=(.*)&amp;stockgdmc',td)[0]
            gdmc=row('td').eq(0).text()[3:]
    
            data.append([gddm,gdmc,'普通股东'])


    return data
        
        
########################################################################
#从大智慧网F10获取股东代码保存
########################################################################
def gddm2sqlite():
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    gpdmb=get_gpdm()
    
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    for i in range(len(gpdmb)):
        gpdm=gpdmb.index[i]
        gpmc = gpdmb.iloc[i]['gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(gpdmb),i+1,gpdm,gpmc)) 
        data = get_gddm(gpdm)
        
        if len(data)>0 :
            dbcn.executemany('INSERT OR REPLACE INTO DZHGDDM (GDDM,GDMC,GDLX) VALUES (?,?,?)', data)

    dbcn.commit()
    dbcn.close()


    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)


########################################################################
#从大智慧网F10获取机构股东代码
########################################################################
def get_jgdm(gpdm):
    sc=scdm(gpdm)
    gpdm=sgpdm(gpdm)
    
    data=[]

    url = 'http://webf10.gw.com.cn/'+sc+'/B5/'+sc+gpdm+'_B5.html'

    browser.get(url)

    html = browser.find_element_by_xpath("//*").get_attribute("outerHTML")  # 不要用 browser.page_source，那样得到的页面源码不标准
     
    html = pq(html)
    html.find("script").remove()    # 清理 <script>...</script>
    html.find("style").remove()     # 清理 <style>...</style>

    try :
        sect=html('section').filter('#机构持仓明细').html()
        
        div=pq(sect)
        tbl=pq(div('div').filter('.jgccmx_tabel_scroll').html())
               
        tr=pq(tbl('table').eq(0).html())
    
    except :
        print("出错退出")
        return data
    
    trs=tr('tr')
    
    for i in range(len(trs)):
#        print(i)
        row=pq(trs.eq(i).html())
        td=row('td').eq(1).html()

        if not td is None:
            gddm=re.findall('stockgddm=(.*)&amp;stockgdmc',td)[0]
            gdmc=row('td').eq(1).text()
            gdlx=row('td').eq(2).text()
    
            data.append([gddm,gdmc,gdlx])

    return data
    
    

'''
python使用pyquery库总结 
https://blog.csdn.net/baidu_21833433/article/details/70313839

'''

########################################################################
#从大智慧网F10获取机构股东代码保存
########################################################################
def jgdm2sqlite():
    
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    browser = webdriver.PhantomJS()

    gpdmb=get_gpdm()
    
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    for i in range(1122,len(gpdmb)):
        gpdm=gpdmb.index[i]
        gpmc = gpdmb.iloc[i]['gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(gpdmb),i+1,gpdm,gpmc)) 
        data = get_jgdm(gpdm)
        
        if len(data)>0 :
            dbcn.executemany('INSERT OR REPLACE INTO DZHGDDM (GDDM,GDMC,GDLX) VALUES (?,?,?)', data)

        dbcn.commit()

    dbcn.close()

    browser.quit()

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)


if __name__ == "__main__":  
    
    '''
    
    select * from dzhgdjc where gddm in ('80064460','80188285','80553146') and rq>='2018-06-30'  and (cgbh='新进' or cast(cgbh as int)>0)
    '''

    createDataBase()
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)
    
    browser = webdriver.PhantomJS()

    data=[]
    gddic={'80188285':'中国证券金融股份有限公司(证金)',
           '80553146':'中央汇金资产管理有限责任公司',
           '80064460':'中央汇金投资有限责任公司(汇金)',
           '80568462':'香港中央结算有限公司',
           '244642432019996':'徐开东',
           '80010104':'香港中央结算(代理人)有限公司(H股)',
           '362132431424179':'赵建平',
           '2644627427':'李欣',
           '2957929756':'王琼',
           '2160826216':'周晨',
           }
#    gddic={
#           '80188285':'中国证券金融股份有限公司',
#           }
    
#    for gddm in gddic:
    gddm='244642432019996'
    gdmc=gddic[gddm]

    url = 'http://webf10.gw.com.cn/gdmcDefault.html?stockgddm=%s&stockgdmc=%s' % (gddm,gdmc)
    browser.get(url)

    '''
    innerHTML、outerHTML、innerText、outerText的区别及兼容性问题
    https://blog.csdn.net/html5_/article/details/23619103
    
    '''
    html = browser.find_element_by_xpath("//*").get_attribute("outerHTML")  # 不要用 browser.page_source，那样得到的页面源码不标准
     
    html = pq(html)
    html.find("script").remove()    # 清理 <script>...</script>
    html.find("style").remove()     # 清理 <style>...</style>
    div=pq(html('div#DIVcontent'))
             
    tbls=div('table')
    
    for i in range(len(tbls)):
        tbl=pq(tbls.eq(i))
        trs=tbl('tr')
        for j in range(len(trs)):
            tr=pq(trs.eq(j))
            tds=tr('td')
            rq=tds.eq(2).text()
            gpdm=tds.eq(3).text()
            gpmc=tds.eq(4).text()
            cgsl=tds.eq(5).text()
            
            cgbh=tds.eq(7).text()
            
            data.append([gddm,gdmc,lgpdm(gpdm),gpmc,rq,cgsl,cgbh])
#            print(gpdm,gpmc,rq,cgsl,cgbh)
                
    if len(data)>0 :
        dbcn.executemany('INSERT OR REPLACE INTO DZHGDJC (GDDM,GDMC,GPDM,GPMC,RQ,CGSL,CGBH) VALUES (?,?,?,?,?,?,?)', data)
    
    dbcn.commit()
    dbcn.close()
    browser.quit()
            

