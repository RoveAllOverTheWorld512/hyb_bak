# -*- coding: utf-8 -*-
"""
功能：本程序从同花顺网提取自选个股解禁数据，保存sqlite
用法：
"""
import datetime
import time
from selenium import webdriver
import sqlite3
import sys
from pyquery import PyQuery as pq
import struct
import os
import sys
import re
import pandas as pd
import winreg

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

########################################################################
#建立数据库
########################################################################
def createDataBase():
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    cn = sqlite3.connect(dbfn)

    cn.execute('''CREATE TABLE IF NOT EXISTS [XSJJ_THS](
                  [GPDM] TEXT NOT NULL, 
                  [JJRQ] TEXT NOT NULL, 
                  [JJSL] TEXT NOT NULL, 
                  [JJSZ] TEXT,
                  [ZGBBL] TEXT);''')
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS [GPDM_JJRQ_XSJJ_THS]
                ON [XSJJ_THS]([GPDM], [JJRQ]);''')



def getdrive():
    return sys.argv[0][:2]

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
#获取本机通达信安装目录，生成自定义板块保存目录
########################################################################
def gettdxblkdir():
    try :
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人生")
        value, type = winreg.QueryValueEx(key, "InstallLocation")
        return value + '\\T0002\\blocknew'
    except :
        print("本机未安装【华西证券华彩人生】软件系统。")
        sys.exit()

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
def gettdxblk(lb):

    try :
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人生")
        value, type = winreg.QueryValueEx(key, "InstallLocation")
    except :
        print("本机未安装【华西证券华彩人生】软件系统。")
        sys.exit()

    blkfn = value + '\\T0002\\hq_cache\\block_'+lb+'.dat'
    blk = {}
    with open(blkfn,'rb') as f :
        blknum, = struct.unpack('384xH', f.read(386))
        for i in range(blknum) :
            stk = []
            blkname = f.read(9).strip(b'\x00').decode('GBK')
            stnum, = struct.unpack('H2x', f.read(4))
            for j in range(stnum) :
                stkid = f.read(7).strip(b'\x00').decode('GBK')
                stk.append(stkid)
            blk[blkname] = [blkname,stnum,stk]

            f.read((400-stnum)*7)
        f.close()


    return blk


#############################################################################
#股票列表,通达信板块文件调用时wjtype="tdxbk"
#############################################################################
def zxglist(zxgfn,wjtype=""):
    zxglst = []
    p = "(\d{6})"
    if wjtype == "tdxblk" :
        p ="\d(\d{6})"
    if os.path.exists(zxgfn) :
        #用二进制方式打开再转成字符串，可以避免直接打开转换出错
        with open(zxgfn,'rb') as dtf:
            zxg = dtf.read()
            if zxg[:3] == b'\xef\xbb\xbf' :
                zxg = zxg.decode('UTF8','ignore')   #UTF-8
            elif zxg[:2] == b'\xfe\xff' :
                zxg = zxg.decode('UTF-16','ignore')  #Unicode big endian
            elif zxg[:2] == b'\xff\xfe' :
                zxg = zxg.decode('UTF-16','ignore')  #Unicode
            else :
                zxg = zxg.decode('GBK','ignore')      #ansi编码
        zxglst =re.findall(p,zxg)
    else:
        print("文件%s不存在！" % zxgfn)
    if len(zxglst)==0:
        print("股票列表为空,请检查%s文件。" % zxgfn)

    zxg = list(set(zxglst))
    zxg.sort(key=zxglst.index)

    return zxg


#############################################################################
#通达信自选股A股列表，去掉了指数代码
#依赖关系
#zxglb---> gettdxblkdir
#     ---> zxglist
#     ---> get_gpdm
#     ---> os    
#############################################################################    
def zxglst():
    zxgfile="zxg.blk"
    tdxblkdir = gettdxblkdir()
    zxgfile = os.path.join(tdxblkdir,zxgfile)
    zxg = zxglist(zxgfile,"tdxblk")
    
    gpdmb=get_gpdm()
    
    #去掉指数代码只保留A股代码
    zxglb=[]
    for e in zxg:
        dm=lgpdm(e)
        if dm in gpdmb.index:
            zxglb.append(dm)
            
    return zxglb



if __name__ == "__main__": 

    print('%s Running' % sys.argv[0])
    gpdmb=get_gpdm()
    gplb=zxglst()            

    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

#    fireFoxOptions = webdriver.FirefoxOptions()
#    fireFoxOptions.set_headless()
#    browser = webdriver.Firefox(firefox_options=fireFoxOptions)

#    browser = webdriver.PhantomJS() #用本句不成功
    browser.get("http://data.10jqka.com.cn/market/xsjj/")
    time.sleep(5)

    for i in range(len(gplb)):

        data = []
        gpdm=gplb[i]
        gpmc = gpdmb.loc[gpdm]['gpmc']
        now = datetime.datetime.now().strftime('%H:%M:%S')

        print("%s  共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (now,len(gplb),i+1,gpdm,gpmc)) 
        
        elem = browser.find_element_by_id("search-center")
        elem.clear()
        elem.send_keys(sgpdm(gpdm))
        time.sleep(5)
        '''
        下面的js代码非常重要，用来关闭弹窗
        '''
        js="document.getElementById('search-center').setAttribute('autocomplete', 'off')"
        browser.execute_script(js) 
        time.sleep(2)
            
        browser.find_element_by_id("search-center-submit").click()
        
        time.sleep(3)
        
        html = browser.find_element_by_class_name("page-table").get_attribute("outerHTML")
    #    html = browser.find_element_by_xpath("//*").get_attribute("outerHTML")
        html = pq(html)
        html.find("script").remove()    # 清理 <script>...</script>
        html.find("style").remove()     # 清理 <style>...</style>
        html=pq(html('tbody'))
     
        rows=html('tr')
     
    
        for j in range(1,len(rows)):
    
            row=rows.eq(j).text().split(' ')
            jjrq=row[1]
            bcjj=row[2]+'万'
            jjsz=row[4]+'万'
            zgbzb=row[5]

            rowdat = [gpdm,jjrq,bcjj,jjsz,zgbzb]
            data.append(rowdat)

        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO XSJJ_THS (GPDM,JJRQ,JJSL,JJSZ,ZGBBL) VALUES (?,?,?,?,?)', data)
            dbcn.commit()

        
    browser.quit()
    dbcn.commit()
    dbcn.close()

