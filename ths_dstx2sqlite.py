# -*- coding: utf-8 -*-
"""
功能：本程序从同花顺网提取大宗交易数据、业绩预告、业绩快报、业绩公告、公告速递等
生成提示信息，保存sqlite
用法：每天运行
"""
import time
import datetime
from selenium import webdriver
import sqlite3
import os
import sys
from pyquery import PyQuery as pq
from configobj import ConfigObj
import struct
import re
import pandas as pd
import winreg

########################################################################
#初始化本程序配置文件
########################################################################
def iniconfig():
    inifile = os.path.splitext(sys.argv[0])[0]+'.ini'  #设置缺省配置文件
    return ConfigObj(inifile,encoding='GBK')


#########################################################################
#读取键值,如果键值不存在，就设置为defvl
#########################################################################
def readkey(config,key,defvl=None):
    keys = config.keys()
    if defvl==None :
        if keys.count(key) :
            return config[key]
        else :
            return ""
    else :
        if not keys.count(key) :
            config[key] = defvl
            config.write()
            return defvl
        else:
            return config[key]


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
    gpdmb['dm']=gpdmb['gpdm'].map(lambda x:x[:6])
    gpdmb=gpdmb.set_index('gpdm',drop=False)
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
#############################################################################    
def zxglst(zxgfile=None):

    if zxgfile==None:
        zxgfile="zxg.blk"
    else:
        if '.blk' not in zxgfile:
            zxgfile=zxgfile+'.blk'
            
    tdxblkdir = gettdxblkdir()
    zxgfile = os.path.join(tdxblkdir,zxgfile)
    if not os.path.exists(zxgfile):
        print("板块不存在，请检查！")
        return pd.DataFrame()
    
    zxg = zxglist(zxgfile,"tdxblk")
    
    gpdmb=get_gpdm()
    
    #去掉指数代码只保留A股代码
    zxglb=gpdmb.loc[gpdmb['dm'].isin(zxg),:]
    #增加一列
    #http://pandas.pydata.org/pandas-docs/stable/generated/pandas.DataFrame.assign.html
    zxglb=zxglb.assign(no=zxglb['dm'].map(lambda x:zxg.index(x)+1))

    zxglb=zxglb.set_index('no') 
    zxglb=zxglb.sort_index()       
    return zxglb

##########################################################################
#获取运行程序所在驱动器
##########################################################################
def getdrive():
    if sys.argv[0]=='' :
        return os.path.splitdrive(os.getcwd())[0]
    else:
        return os.path.splitdrive(sys.argv[0])[0]

########################################################################
#建立数据库
########################################################################
def createDataBase():
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    cn = sqlite3.connect(dbfn)

    cn.execute('''CREATE TABLE IF NOT EXISTS [DZJY_THS](
                  [GPDM] TEXT NOT NULL, 
                  [RQ] TEXT NOT NULL, 
                  [CJJ] REAL NOT NULL, 
                  [CJL] REAL NOT NULL, 
                  [CJE] REAL NOT NULL, 
                  [ZYL] REAL NOT NULL, 
                  [MRF] TEXT NOT NULL, 
                  [MCF] TEXT NOT NULL);
                ''')

'''
CREATE TABLE [THS](
  [GPDM] TEXT NOT NULL, 
  [RQ] TEXT NOT NULL, 
  [TS1] TEXT, 
  [TS2] TEXT, 
  [TSLX] TEXT NOT NULL);

CREATE UNIQUE INDEX [GPDM_RQ_TS1_TS2_THS]
ON [THS](
  [GPDM], 
  [RQ], 
  [TS1], 
  [TS2]);
'''
    

def dstx_dzjy(lastdate):
    
    print('正在处理大宗交易……')

    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/market/dzjy/'
    browser.get(url)
    time.sleep(5)

    elem = browser.find_element_by_class_name("page_info")
    pgs=int(1/eval(elem.text))
    ts1='大宗交易'
    while True:
        pages=browser.find_element_by_class_name('m-page')
        cur=eval(pages.find_element_by_class_name('cur').text)
        dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
        html = pq(dbtbl)

        print("正在处理第%d/%d页，请等待。" % (cur,pgs))

        data = []
        rows=html('tr')
        
        mrf0=None
        zyl0=None
        zje=0
        for i in range(1,len(rows)):
            
            rowdat=[]    
            row=pq(rows('tr').eq(i))
            
            rq=row('td').eq(1).text()
            dm=row('td').eq(2).text()            
            cjj=float(row('td').eq(5).text())            
            cjl=float(row('td').eq(6).text())            
            cje=round(cjj*cjl,2)
            zyl=row('td').eq(7).text()
            mrf=row('td').eq(8).text()       

            if zyl==zyl0 and mrf==mrf0 :
                zje=zje+cje
                continue
            else:
                zje=cje
                
            if float(zyl.replace('%',''))<0:
                ts1='大宗交易折价,折溢率%s%%' % zyl
            else:
                ts1='大宗交易溢价,折溢率%s%%' % zyl

                
            ts2='买方：%s,折溢率%s,成交额%d万元' % (mrf,zyl,zje)
            dm = lgpdm(dm)       

            rowdat = [dm,rq,ts1,ts2,'0']
            data.append(rowdat)
            mrf0=mrf
            zyl0=zyl
            zje=0
    
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data)
            dbcn.commit()

        if cur<pgs:
            elem = browser.find_element_by_xpath("//a[text()='下一页']")
            elem.click()
            time.sleep(5)
        else:
            break
        
        if rq<lastdate:
            break
        
    browser.quit()
    dbcn.commit()
    dbcn.close()

    return    

def dstx_yjyg(lastdate):
    
    print('正在处理业绩预告……')
    
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/financial/yjyg/'
    browser.get(url)
    time.sleep(5)

    elem = browser.find_element_by_class_name("page_info")
    pgs=int(1/eval(elem.text))

    while True:
        pages=browser.find_element_by_class_name('m-page')
        cur=eval(pages.find_element_by_class_name('cur').text)
        dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
        html = pq(dbtbl)

        print("正在处理第%d/%d页，请等待。" % (cur,pgs))

        data = []
        rows=html('tr')
 
        for i in range(1,len(rows)):

            rowdat=[]    
            row=pq(rows('tr').eq(i))
            
            rq=row('td').eq(7).text()
            dm=row('td').eq(1).text()            
            ts1=row('td').eq(3).text()            
            ts2=row('td').eq(4).text()

            dm = lgpdm(dm)       

            rowdat = [dm,rq,ts1,ts2,'0']
            data.append(rowdat)
    
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data)
            dbcn.commit()

        if cur<pgs:
            elem = browser.find_element_by_xpath("//a[text()='下一页']")
            elem.click()
            time.sleep(5)
        else:
            break

        if rq<lastdate:
            break
        
    browser.quit()
    dbcn.commit()
    dbcn.close()

    return


def dstx_ggsd(lastdate):    
    print('正在处理公告速递……')
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/market/ggsd/'
    browser.get(url)
    time.sleep(5)

    bbs = browser.find_elements_by_class_name('J-board-item')
    for n in range(len(bbs)):
        bbs[n].click()
        time.sleep(5)
        
        print('正在处理公告速递:[%s]……' % bbs[n].text )

        elem = browser.find_element_by_class_name("page_info")
        pgs=int(1/eval(elem.text))
        
        while True:
            pages=browser.find_element_by_class_name('m-page')
            cur=eval(pages.find_element_by_class_name('cur').text)
            dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
            html = pq(dbtbl)
    
            print("正在处理第%d/%d页，请等待。" % (cur,pgs))
    
            data = []
            rows=html('tr')
     
            for i in range(1,len(rows)):
    
                rowdat=[]    
                row=pq(rows('tr').eq(i))
                tds=row('td')
                if len(tds)==2:
                    ggbt=tds.eq(0).text()
                    gglx=tds.eq(1).text()
                else:
                    rq=tds.eq(1).text()
                    dm=tds.eq(2).text() 
                    dm = lgpdm(dm)    
                    ggbt=None
                    gglx=None
    
    #            if gglx in ('持股变动公告','股权激励','股票质押公告','资产购买公告','增发事项公告',):
                if gglx!=None and gglx!='--':
                    rowdat = [dm,rq,gglx,ggbt,'0']
                    data.append(rowdat)
                    
            #由于这个网页存在嵌套表，pyQuery分析时行数会被递归多次计算，出现数据重复，下面是对数据去重        
            data1=[]
            for dt in data:
                if dt not in data1:
                    data1.append(dt)                
                    
            if len(data1)>0:
                dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data1)
                dbcn.commit()
    
            if cur<pgs:
                elem = browser.find_element_by_xpath("//a[text()='下一页']")
                elem.click()
                time.sleep(5)
            else:
                break

            if rq<lastdate:
                break
            
    browser.quit()
    dbcn.commit()
    dbcn.close()

    return

def dstx_yjkb(lastdate):
    print('正在处理业绩快报……')
    
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/financial/yjkb/'
    browser.get(url)
    time.sleep(5)
    try:
        elem = browser.find_element_by_class_name("page_info")
        pgs=int(1/eval(elem.text))
    except:
        pgs=1

    while True:
        try:
            pages=browser.find_element_by_class_name('m-page')
            cur=eval(pages.find_element_by_class_name('cur').text)
        except:
            cur=1
            
        dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
        html = pq(dbtbl)

        print("正在处理第%d/%d页，请等待。" % (cur,pgs))

        data = []
        rows=html('tr')
 
        for i in range(2,len(rows)):

            rowdat=[]    
            row=pq(rows('tr').eq(i))
            
            rq=row('td').eq(3).text()
            dm=row('td').eq(1).text()    
            yysr=row('td').eq(4).text()
            yysr_g = row('td').eq(6).text()
            jlr=row('td').eq(8).text()
            jlr_g = row('td').eq(10).text()
            eps = row('td').eq(12).text()
            roe = row('td').eq(12).text()
            
            if float(jlr_g)>0 :
                ts1='业绩快报:净利润增长,净利润同比%s%%' % jlr_g
            else :
                ts1='业绩快报:净利润减少,净利润同比%s%%' % jlr_g

            ts2='业绩快报,营业收入%s,同比%s%%,净利润%s,同比%s%%,EPS%s元,ROE%s%%' % (yysr,yysr_g,jlr,jlr_g,eps,roe)
                
            dm = lgpdm(dm)       

            rowdat = [dm,rq,ts1,ts2,'0']
            data.append(rowdat)
    
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data)
            dbcn.commit()

        if cur<pgs:
            elem = browser.find_element_by_xpath("//a[text()='下一页']")
            elem.click()
            time.sleep(5)
        else:
            break

        if rq<lastdate:
            break
        
    browser.quit()
    dbcn.commit()
    dbcn.close()
    return

def dstx_yjgg(lastdate):
    print('正在处理业绩公告……')
    
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/financial/yjgg/'

    browser.get(url)
    time.sleep(5)
    try:
        elem = browser.find_element_by_class_name("page_info")
        pgs=int(1/eval(elem.text))
    except:
        pgs=1

    while True:
        try:
            pages=browser.find_element_by_class_name('m-page')
            cur=eval(pages.find_element_by_class_name('cur').text)
        except:
            cur=1
            
        dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
        html = pq(dbtbl)

        print("正在处理第%d/%d页，请等待。" % (cur,pgs))

        data = []
        rows=html('tr')
 
        for i in range(2,len(rows)):

            rowdat=[]    
            row=pq(rows('tr').eq(i))
            
            rq=row('td').eq(3).text()
            dm=row('td').eq(1).text()  
            
            yysr=row('td').eq(4).text()
            yysr_g = row('td').eq(5).text()
            jlr=row('td').eq(7).text()
            jlr_g = row('td').eq(8).text()
            eps = row('td').eq(10).text()
            roe = row('td').eq(12).text()
            
            if float(jlr_g)>0 :
                ts1='业绩公告:净利润增长,净利润同比%s%%' % jlr_g
            else :
                ts1='业绩公告:净利润减少,净利润同比%s%%' % jlr_g

            ts2='业绩公告,营业收入%s,同比%s%%,净利润%s,同比%s%%,EPS%s元,ROE%s%%' % (yysr,yysr_g,jlr,jlr_g,eps,roe)
                
            dm = lgpdm(dm)       

            rowdat = [dm,rq,ts1,ts2,'0']
            data.append(rowdat)
    
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data)
            dbcn.commit()

        if cur<pgs:
            elem = browser.find_element_by_xpath("//a[text()='下一页']")
            elem.click()
            time.sleep(5)
        else:
            break

        if rq<lastdate:
            break
        
    browser.quit()
    dbcn.commit()
    dbcn.close()
    
    return

def ths_dstx(gpdm,rq):
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)
    curs = dbcn.cursor()

    sql='select gpdm,rq,ts1,ts2 from ths where gpdm=? and rq>=? order by rq desc;'
    curs.execute(sql,(gpdm,rq))        
    data = curs.fetchall()
    cols = ['gpdm','rq','ts1','ts2']
    
    df=pd.DataFrame(data,columns=cols)
    
    df['rq']=pd.to_datetime(df['rq'],format='%Y-%m-%d')

    return df
        

if __name__ == "__main__": 
    print('%s Running' % sys.argv[0])

    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    today=datetime.datetime.now().strftime('%Y-%m-%d')
#    config = iniconfig()
#    lastdate = readkey(config,'lastdate')
#
#    dstx_dzjy(lastdate)    #大宗交易  
#    
#    dstx_yjyg(lastdate)    #业绩预告
#    
#    dstx_yjkb(lastdate)    #业绩快报
#    
#    dstx_yjgg(lastdate)    #业绩公告
#    
#    dstx_ggsd(lastdate)    #公告速递
#     
#    config['lastdate'] = today
#    config.write()
    rq='2018-01-01'
    xlsfn='%s\\selestock\\dstx%s.xlsx' % (getdrive(),today)
    writer=pd.ExcelWriter(xlsfn,engine='xlsxwriter')
    workbook = writer.book
    cell_format1 = workbook.add_format({#一种方法可以直接在字典里 设置属性
                                        'font_size':  9,   #字体大小
                                        'align':    'center',
                                        'valign':   'vcenter',
                                        })
    cell_format2 = workbook.add_format({#一种方法可以直接在字典里 设置属性
                                        'font_size':  9,   #字体大小
                                        'align':    'left',
                                        'valign':   'vcenter',
                                        'num_format': 'yyyy-mm-dd'
                                        })

    cell_format3 = workbook.add_format({#一种方法可以直接在字典里 设置属性
                                        'font_size':  9,   #字体大小
                                        'align':    'left',
                                        'valign':   'vcenter',
                                        })
    cell_format4 = workbook.add_format({
                                        'border':1,       #单元格边框宽度
                                        })
    cell_format5 = workbook.add_format({'bg_color': '#FFC7CE',
                                       'font_color': '#9C0006'
                                       })
    

    '''
    http://xlsxwriter.readthedocs.io/working_with_conditional_formats.html#working-with-conditional-formats
    '''
    zxg=zxglst('cg')
    for i in range(len(zxg)):
        gpdm=zxg.iloc[i]['gpdm']
        gpmc=zxg.iloc[i]['gpmc']
        shtname='%d.%s' % (i+1,gpmc)
        df= ths_dstx(gpdm,rq)       

        df.to_excel(writer, sheet_name=shtname,index=False)   

        worksheet = writer.sheets[shtname]
        shtdic=worksheet.__dict__
        rows=shtdic['dim_rowmax']
        cols=shtdic['dim_colmax']
        worksheet.set_column('A:A', 10,cell_format1)
        worksheet.set_column('B:B', 10,cell_format2)
        worksheet.set_column('C:C', 15,cell_format3)
        worksheet.set_column('D:D', 100,cell_format3)

        worksheet.conditional_format(0,0,rows,cols, {'type': 'cell',
                                         'criteria': '!=',
                                         'value': 0,
                                         'format': cell_format4})


        worksheet.conditional_format(0,0,rows,cols, {'type':     'formula',
                                        'criteria': '=mod(row(),2)=0',
                                        'format':   cell_format5})
    writer.save()



    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

