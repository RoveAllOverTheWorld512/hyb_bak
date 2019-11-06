# -*- coding: utf-8 -*-
"""
股票所属行业、风格、概念等
"""

import os
import sys
import re
import datetime
from configobj import ConfigObj
import sqlite3
import numpy as np
import pandas as pd
import xlwings as xw
import struct
import winreg
from selenium import webdriver
import time


########################################################################
#建立数据库
########################################################################
def createDataBase():
    dbfn=getdrive()+'\\hyb\\STOCKEPS.db'
    cn = sqlite3.connect(dbfn)

    """
    股票代码表：股票代码，股票名称
    """

    cn.execute('''CREATE TABLE IF NOT EXISTS GPDM
           (GPDM TEXT PRIMARY KEY,
           GPMC TEXT);''')

    """
    股票EPS：股票代码，日期，基本EPS0，稀释EPS1，基本EPS0同比增长率，稀释EPS1同比增长率
    """

    cn.execute('''CREATE TABLE IF NOT EXISTS GPEPS
           (GPDM TEXT,
           RQ TEXT,
           EPS0 REAL,
           EPS1 REAL,
           EPS0_G REAL,
           EPS1_G REAL);''')

    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPEPS_GPDM_RQ ON GPEPS(GPDM,RQ);''')
 
    """
    股票成长性：股票代码，日期，营业总收入同比增长率(%)，营业收入同比增长率(%)，净利润同比增长率(%)，
    扣除非经常性损益后的净利润同比增长率(%)，营业收入(元)，净利润(元)，非经常性损益(元)，上年同期净利润(元)
    
    """

    cn.execute('''CREATE TABLE IF NOT EXISTS GPGROWTH
           (GPDM TEXT,
           RQ TEXT,
           YYZSR_G REAL,
           YYSR_G REAL,
           JLR_G REAL,
           KFJLR_G REAL,
           YYSR REAL,
           JLR REAL,
           FSY REAL,
           SNJLR REAL);''')

    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPGROWTH_GPDM_RQ ON GPGROWTH(GPDM,RQ);''')

    """
    股票年报关键数据：股票代码，日期，经营现金流量净额，营业收入，净利润，资产总计，流动资产，净资产，
    商誉，带息债务，应收账款，预付账款，，应付款项，预收款项，加权净资产收益率，杠杆倍数，带息负债率，商誉净资产占比，商誉总资产占比，
    收入经营现金含量，净利润经营现金含量，应收账款收入占比，被占资金，被占资金与总资产比，被占资金与总收入比，被占资金与净利润比
    算法：
    杠杆倍数=净资产/资产总计
    带息负债率=带息债务/资产总计
    商誉净资产占比=商誉/净资产
    商誉总资产占比=商誉/总资产
    收入经营现金含量=经营现金流量净额/营业收入
    净利润经营现金含量=经营现金流量净额/净利润
    应收账款收入占比=应收账款/营业收入
    被占资金=应收账款+预付款项-应付账款-预收款项
    被占资金与总资产比=被占资金/总资产
    被占资金与营业收入比=被占资金/营业收入
    被占资金与净利润比=被占资金/净利润
    
    """

    cn.execute('''CREATE TABLE IF NOT EXISTS GPNB
           (GPDM TEXT,
           RQ TEXT,
           JYXJL REAL,
           YYSR REAL,
           JLR REAL,
           ZCZJ REAL,
           LDZC REAL,
           JZC REAL,
           SY REAL,
           DXZW REAL,
           YSZK REAL,
           YFZK REAL,
           YSKX REAL,
           YFKX REAL,
           ROE REAL,
           GGBS REAL,
           DXFZL REAL,
           SYJZCZB REAL,
           SYZZCZB REAL,
           SRXJLHL REAL,
           LRXJLHL REAL,
           YSZKSRZB REAL,
           BZZJ REAL,
           BZZJZCZB REAL,
           BZZJSRZB REAL,
           BZZJLRZB REAL
           );''')

    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPNB_GPDM_RQ ON GPNB(GPDM,RQ);''')

    cn.commit()

########################################################################
#初始化本程序配置文件
########################################################################
def iniconfig():
    inifile = os.path.splitext(sys.argv[0])[0]+'.ini'  #设置缺省配置文件
    return ConfigObj(inifile,encoding='GBK')


#############################################################################
#读取dbf文件
#############################################################################
def dbfreader(f):

    numrec, lenheader = struct.unpack('<xxxxLH22x', f.read(32))
    numfields = (lenheader - 33) // 32

    fields = []
    for fieldno in range(numfields):
        name, typ, size, deci = struct.unpack('<11sc4xBB14x', f.read(32))
        name = name.decode().replace('\x00', '')
        typ  = typ.decode()
        fields.append((name, typ, size, deci))
    yield [field[0] for field in fields]
    yield [tuple(field[1:]) for field in fields]

    terminator = f.read(1)

    fields.insert(0, ('DeletionFlag', 'C', 1, 0))
    fmt = ''.join(['%ds' % fieldinfo[2] for fieldinfo in fields])
    fmtsiz = struct.calcsize(fmt)

    for i in range(numrec):
        record = struct.unpack(fmt, f.read(fmtsiz))
        if record[0].decode() != ' ':
            continue                        # deleted record
        result = []
        for (name, typ, size, deci), value in list(zip(fields, record)):
            if name == 'DeletionFlag':
                continue
            if typ == "C":
                value = value.strip(b'\x00').decode('GBK')
            if typ == "N":
                value = value.strip(b'\x00').strip(b'\x20').decode('GBK')
                if value == '':
                    value = 0
                elif deci:
                    value = float(value)
                else:
                    value = int(value)
            elif typ == 'D':
                value = value.decode('GBK')
            elif typ == 'L':
                value = value.decode('GBK')
                value = (value in 'YyTt' and 'T') or (value in 'NnFf' and 'F')

            result.append(value)

        yield result

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
            if (sc=='z' and (gpdm[0]=='0' or gpdm[0:2]=='30')) :
                gpdm=gpdm+'.SZ'
                datacode.append([gpdm,gpmc,gppy])
            ss = f.read(314)
        f.close()
    gpdmb=pd.DataFrame(datacode,columns=['gpdm','gpmc','gppy'])
    return gpdmb

###############################################################################
#读取dbf到pandas
###############################################################################
def dbf2pandas(dbffn,cols):
    with open(dbffn,"rb") as f:
        data = list(dbfreader(f))
        f.close()
    columns = data[0]
    columns=[e.lower() for e in columns]
    data = data[2:]
    df = pd.DataFrame(data,columns=columns)
    if len(cols) == 0 :
        return df
    else :
        return df[cols]


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
#同花顺行业分类
###############################################################################
def thshy():
    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)
    thshyfn=ths_dl_xls()
    wb = xw.Book(thshyfn)
    wb.sheets[0].range('A1').value="gpdm"
    wb.sheets[0].range('B1').value="gpmc"
    wb.sheets[0].range('C1').value="thshy"
    data=wb.sheets[0].range('A1').options(pd.DataFrame, expand='table').value
    xw.apps[0].quit()
    #将索引股票代码加为一列
    data['gpdm']=data.index
    #提取行业、去重、排序、编码
    hydm=data.loc[:,'thshy']
    hydm=hydm.drop_duplicates()
    hydm=hydm.sort_values()
    hydmdf=hydm.to_frame()
    #编码    
    hydmdf['zldm']="THS"
    hydmdf['hydm']=""
    i=1
    for index, row in hydmdf.iterrows():
        hydmdf.loc[index,'hydm']=i
        i += 1
    hydmdf=hydmdf.set_index('hydm',drop=False)

    hydmdf=hydmdf.loc[:,['zldm','hydm','thshy']]    
    hydmlst=hydmdf.values.tolist()
    
    cn.executemany('INSERT OR IGNORE INTO FLDM (ZLDM,FLDM,FLMC) VALUES (?,?,?)', hydmlst)
    cn.commit()
    
    #将股票行业编码化    
    gphydf=pd.merge(data,hydmdf,on='thshy')  
    gphydf=gphydf.loc[:,['gpdm','zldm','hydm']]    
    gphylst=gphydf.values.tolist()
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', gphylst)
    cn.commit()

    cn.close()

###############################################################################
#下载文件名，参数1表示如果文件存在则将原有文件名用其创建时间命名
###############################################################################
def dlfn(dldir):

    today=datetime.datetime.now().strftime("%Y-%m-%d")
    dlfn = today+'.xls'
    fn = os.path.join(dldir,dlfn)

    if os.path.exists(fn):
        ctime=os.path.getctime(fn)  #文件建立时间
        ltime=time.localtime(ctime)
        newfn = time.strftime("%Y%m%d%H%M%S",ltime)+'.xls'
        os.rename(fn,os.path.join(os.path.dirname(fn),newfn))

    return fn

###############################################################################
#从同花顺i问财下载,prefix:eps、growth、rpt
###############################################################################
def dl_ths_xls(prefix):

    config = iniconfig()
    
    ddir=os.path.join(getdrive(),readkey(config,'dldir'))
    dafn = dlfn(ddir)

    nf1 = int(readkey(config,prefix + 'nf1'))
    nf2 = int(readkey(config,prefix + 'nf2'))
    nb = readkey(config, prefix + 'rq')
    kw0 = readkey(config, prefix + 'kw')
    sele = readkey(config, prefix + 'sl')
    newfn0 = readkey(config, prefix + 'fn')

    username = readkey(config,'iwencaiusername')
    pwd = readkey(config,'iwencaipwd')
    
    profile = webdriver.FirefoxProfile()
    profile.set_preference('browser.download.dir', ddir)
    profile.set_preference('browser.download.folderList', 2)
    profile.set_preference('browser.download.manager.showWhenStarting', False)
    
    #http://www.w3school.com.cn/media/media_mimeref.asp
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/vnd.ms-excel')
    
    browser = webdriver.Firefox(firefox_profile=profile)

    #浏览器窗口最大化
    browser.maximize_window()
    #登录同花顺
    browser.get("http://upass.10jqka.com.cn/login")
    #time.sleep(1)
    elem = browser.find_element_by_id("username")
    elem.clear()
    elem.send_keys(username)
    
    elem = browser.find_element_by_class_name("pwd")
    elem.clear()
    elem.send_keys(pwd)
    
    browser.find_element_by_id("loginBtn").click()
    time.sleep(2)
    
    for j in range(nf1,nf2+1):

        kw = str(j) + nb + kw0
        newfn = newfn0 + str(j) + '.xls'
        newfn = os.path.join(ddir,newfn)
        
        if os.path.exists(newfn):
            os.remove(newfn)

        browser.get("http://www.iwencai.com/")
        time.sleep(5)
        browser.find_element_by_id("auto").clear()
        browser.find_element_by_id("auto").send_keys(kw)
        browser.find_element_by_id("qs-enter").click()
        time.sleep(10)
        
        #打开查询项目选单
        trigger = browser.find_element_by_class_name("showListTrigger")
        trigger.click()
        time.sleep(1)
        
        #获取查询项目选单
        checkboxes = browser.find_elements_by_class_name("showListCheckbox")
        indexstrs = browser.find_elements_by_class_name("index_str")
        
        #去掉选项前的“√”
        #涨幅、股价保留
        for i in range(0,len(checkboxes)):
            checkbox=checkboxes[i]
            
            #对于“pe,ttm”之类中间的逗号用下划线代替，注意配置文件也需要这样
            indexstr=indexstrs[i].text.replace(",","_")
            
            if checkbox.is_selected() and not indexstr in sele :
                checkbox.click()
            if not checkbox.is_selected() and indexstr in sele :
                checkbox.click()

        #向上滚屏
        js="var q=document.documentElement.scrollTop=0"  
        browser.execute_script(js)  
        time.sleep(3) 
        
        #关闭查询项目选单
        trigger = browser.find_element_by_class_name("showListTrigger")
        trigger.click()
        time.sleep(3)
        
        #导出数据
        elem = browser.find_element_by_class_name("export.actionBtn.do") 
        #在html中类名包含空格
        elem.click() 
        time.sleep(10)

        if os.path.exists(dafn):
            os.rename(dafn,newfn)
            
    browser.quit()

#########################################################################
#读取键值
#########################################################################
def readkey(config,key):
    keys = config.keys()
    if keys.count(key) :
        return config[key]
    else :
        return ""

##########################################################################
#获取运行程序所在驱动器
##########################################################################
def getpath():
    return os.path.dirname(sys.argv[0])


#############################################################################
#获取市盈率文件交易日列表
#############################################################################
def jyrlist(syldir):
    files = os.listdir(syldir)
    fs = [re.findall('csi(\d{8})\.xls',e) for e in files]
    jyrqlist =[]
    for e in fs:
        if len(e)>0:
            jyrqlist.append(e[0])

    return sorted(jyrqlist,reverse=1)


##########################################################################
#获取运行程序所在驱动器
##########################################################################
def getdrive():
    return os.path.splitdrive(sys.argv[0])[0]
#def getdrive():
#    return sys.argv[0][:2]


##########################################################################
#删除所有表
##########################################################################
def DeleTables(dbfn):    
#    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)
    curs = cn.cursor()

    curs.execute('''SELECT name FROM sqlite_master WHERE type ='table' 
                 AND name != 'sqlite_sequence';''')
    
    data = curs.fetchall()
    for tbn in data:
        print("DROP TABLE "+ tbn[0])
        curs.execute("DROP TABLE "+ tbn[0])

    cn.commit()
    cn.close()
        
##########################################################################
#查询股票
##########################################################################
def Query(gpdm):
    dbfn=getdrive()+'\\hyb\\STOCKEPS.db'
    cn = sqlite3.connect(dbfn)
    curs = cn.cursor()
    sql='''select gpfldm.gpdm,gpdm.gpmc,gpfldm.zldm,zldm.ZLMC,gpfldm.fldm,fldm.flmc 
            from gpfldm,gpdm,zldm,fldm where gpfldm.gpdm=="'''+gpdm.upper()+'''" 
            and gpfldm.gpdm==gpdm.gpdm and gpfldm.ZLDM==zldm.ZLDM 
            and gpfldm.fldm=fldm.fldm and fldm.zldm=zldm.zldm;'''
    curs.execute(sql)        
    data = curs.fetchall()
    
    return data

##########################################################################
#股票代码表
##########################################################################
def gpdmtbl():
    
    dbfn=getdrive()+'\\hyb\\STOCKEPS.db'
    cn = sqlite3.connect(dbfn)
    gpdmb=get_gpdm()
    data=[[a[0],a[1]] for a in gpdmb.values.tolist()]
    cn.executemany('INSERT OR IGNORE INTO GPDM (GPDM,GPMC) VALUES (?,?)', data)

    cn.commit()
    cn.close()    
    

##########################################################################
#股票质押文件名
##########################################################################
def gpzyfn():
    config = iniconfig()
    ddir=os.path.join(getdrive(),readkey(config,'dldir'))
    today=datetime.datetime.now()
    n = datetime.datetime.weekday(today)
    t1=(today+datetime.timedelta(-8-n)).strftime("%Y%m%d")
    t2=(today+datetime.timedelta(-2-n)).strftime("%Y%m%d")
    fn="gpzyhgmx_" + t1 +"_" +t2 + ".xls"

    return os.path.join(ddir,fn)
    
##########################################################################
#获取最新市盈率文件名
##########################################################################
def syl_pe_fn(pedir):
    
    pedir = 'pedir' if pedir=='pedir' else 'syldir'
    pref = '' if pedir=='pedir' else 'csi'
    
    config = iniconfig()
    pedir=os.path.join(getdrive(),readkey(config,pedir))
    files = os.listdir(pedir)
    fs = [re.findall('.*(\d{8})\.xls',e) for e in files]
    jyrqlist =[]
    for e in fs:
        if len(e)>0:
            jyrqlist.append(e[0])

    jyr= sorted(jyrqlist,reverse=1)

    return os.path.join(pedir,pref+jyr[0]+'.xls')


##########################################################################
#生成字段名字典
##########################################################################
def flddic(nf):
    
    fld_dic={'股票代码':'gpdm','股票简称':'gpmc'}
    
    fld = [['基本每股收益(元)','eps0'],
            ['稀释每股收益(元)','eps1'],
            ['基本每股收益(同比增长率)(%)','eps0_g'],
            ['稀释每股收益(同比增长率)(%)','eps1_g'],
            ['营业总收入(同比增长率)(%)','yyzsr_g'],
            ['营业收入(同比增长率)(%)','yysr_g'],
            ['净利润同比增长率(%)','jlr_g'],
            ['扣除非经常性损益后的净利润同比增长率(%)','kfjlr_g'],
            ['营业收入(元)','yysr'],
            ['净利润(元)','jlr'],
            ['非经常性损益(元)','fsy'],
            ['上年同期净利润(元)','snjlr'],
            ['经营现金流量净额(元)','jyxjl'],
            ['资产总计(元)','zczj'],
            ['流动资产(元)','ldzc'],
            ['净资产(元)','jzc'],
            ['商誉(元)','sy'],
            ['带息债务(元)','dxzw'],
            ['应收账款(元)','yszk'],
            ['预付账款(元)','yfzk'],
            ['应付款项(元)','yfkx'],
            ['预收款项(元)','yskx'],
            ['加权净资产收益率(%)','roe']
            ]
    
    for key,value in fld:
        key = key + nf + '.12.31'
        fld_dic[key] = value
        
    return fld_dic
    
##########################################################################
#读取eps数据
##########################################################################
def read_eps(xlsfn):
    wb = xw.Book(xlsfn)
        
    nf=re.findall('.+_(\d{4})\.xls',xlsfn)[0]
    
    
    #生成字段名字典
    fld_dic=flddic(nf)
    
    c = len(xw.Range('A1').expand('right').columns)
    #修改字段名
    for i in range(1,c+1):
        fldn = xw.Range((1,i)).value
        
        if fldn in fld_dic.keys() :
            xw.Range((1,i)).value = fld_dic[fldn]
    
    #读取数据
    data = wb.sheets[0].range('A1').options(pd.DataFrame, expand='table').value

    '''下面的语句很重要，MultiIndex转换成Index'''
    data.columns=[e[0] for e in data.columns]
    
    ''' 注意：数据列的元素数据类型有两种：str、float，运行下条语句后都变成了numpy.float64'''
    '''下面的语句很重要，运行后面的保留小数位数就不会出错'''
    data=data.replace('--',np.nan)   

    data=data.drop('gpmc',axis=1)

    '''保留2位小数必须在data=data.replace(np.nan,'--') 前执行
    注意：执行round(2)必须保证同一列各元素的数据类型是一致的,float和numpy.float64是两种不同的类型
    '''
    data=data.round(2)

    '''只保留至少有2项有效数字的行'''
    data=data.dropna(thresh=2)  
    
    data['rq']=nf+'.12.31'
    data['dm']=data.index

#    wb.close()
    xw.apps[0].quit()
    
    return data[['dm','rq','eps0','eps1','eps0_g','eps1_g']]
    
##########################################################################
#读取growth数据
##########################################################################
def read_growth(xlsfn):
    wb = xw.Book(xlsfn)
        
    nf=re.findall('.+_(\d{4})\.xls',xlsfn)[0]
    
    #生成字段名字典
    fld_dic=flddic(nf)
    
    c = len(xw.Range('A1').expand('right').columns)
    #修改字段名
    for i in range(1,c+1):
        fldn = xw.Range((1,i)).value
        
        if fldn in fld_dic.keys() :
            xw.Range((1,i)).value = fld_dic[fldn]
    
    #读取数据
    data = wb.sheets[0].range('A1').options(pd.DataFrame, expand='table').value

    '''下面的语句很重要，MultiIndex转换成Index'''
    data.columns=[e[0] for e in data.columns]

    #删除有效数据少于2项的股票
    data=data.replace('--',np.nan) 
    data=data.drop('gpmc',axis=1)
    data=data.dropna(thresh=2)  

    #单位换算
    data['yysr'] = data['yysr'].map(y2yy)
    data['jlr'] = data['jlr'].map(y2wy)
    data['fsy'] = data['fsy'].map(y2wy)
    data['snjlr'] = data['snjlr'].map(y2wy)


    '''保留2位小数'''
    data=data.round(2)

    data['rq'] = nf + '.12.31'
    data['dm'] = data.index

#    wb.close()
    xw.apps[0].quit()
    
    return data[['dm','rq','yyzsr_g','yysr_g','jlr_g','kfjlr_g','yysr','jlr','fsy','snjlr']]

##########################################################################
#读取rpt年报数据
##########################################################################
def read_rpt(xlsfn):
    wb = xw.Book(xlsfn)
        
    nf=re.findall('.+_(\d{4})\.xls',xlsfn)[0]
    
    #生成字段名字典
    fld_dic=flddic(nf)
    
    c = len(xw.Range('A1').expand('right').columns)
    #修改字段名
    for i in range(1,c+1):
        fldn = xw.Range((1,i)).value
        
        if fldn in fld_dic.keys() :
            xw.Range((1,i)).value = fld_dic[fldn]
    
    #读取数据
    data = wb.sheets[0].range('A1').options(pd.DataFrame, expand='table').value
    
    '''下面的语句很重要，MultiIndex转换成Index'''
    data.columns=[e[0] for e in data.columns]
    
    #删除有效数据少于2项的股票
    data=data.replace('--',np.nan) 
    
#    data=data.drop('gpmc',axis=1)
    data=data.dropna(thresh=2)  
    
    #用应收账款、预付账款、预收款项、应付款项用0代替缺失值
    values = {'yszk': 0, 'yfzk': 0, 'yskx': 0, 'yfkx': 0}
    data=data.fillna(value=values)
    
    #带息债务为0的用np.nan
    data=data.replace({'dxzw': 0},np.nan)
    """
    股票年报关键数据：股票代码，日期，经营现金流量净额，营业收入，净利润，资产总计，流动资产，净资产，
    商誉，带息债务，应收账款，预付账款，，应付款项，预收款项，加权净资产收益率，杠杆倍数，带息负债率，商誉净资产占比，商誉总资产占比，
    收入经营现金含量，净利润经营现金含量，应收账款收入占比，被占资金，被占资金与总资产比，被占资金与总收入比，被占资金与净利润比
    算法：
    杠杆倍数=净资产/资产总计
    带息负债率=带息债务/资产总计
    商誉净资产占比=商誉/净资产
    商誉总资产占比=商誉/总资产
    收入经营现金含量=经营现金流量净额/营业收入
    净利润经营现金含量=经营现金流量净额/净利润
    应收账款收入占比=应收账款/营业收入
    被占资金=应收账款+预付款项-应付账款-预收款项
    被占资金与总资产比=被占资金/总资产
    被占资金与营业收入比=被占资金/营业收入
    被占资金与净利润比=被占资金/净利润
    
    """

    #单位换算JYXJL,YYSR,JLR,ZCZJ,LDZC,JZC,SY,DXZW,YSZK,ROE,GGBS,DXFZL,SYJZCZB,SYZZCZB,SRXJLHL,LRXJLHL,YSZKSRZB
    data['jyxjl'] = data['jyxjl'].map(y2wy)
    data['yysr'] = data['yysr'].map(y2wy)
    data['jlr'] = data['jlr'].map(y2wy)
    
    data['zczj'] = data['zczj'].map(y2yy)
    data['ldzc'] = data['ldzc'].map(y2yy)
    data['jzc'] = data['jzc'].map(y2yy)
    data['sy'] = data['sy'].map(y2yy)
    
    data['dxzw'] = data['dxzw'].map(y2yy)
    
    data['yszk'] = data['yszk'].map(y2wy)
    data['yfzk'] = data['yfzk'].map(y2wy)
    data['yskx'] = data['yskx'].map(y2wy)
    data['yfkx'] = data['yfkx'].map(y2wy)


    data.eval('ggbs = zczj / jzc',inplace=True)

    data.eval('dxfzl = dxzw / zczj * 100',inplace=True)
    
    data.eval('syjzczb = sy / jzc * 100',inplace=True)

    data.eval('syzzczb = sy / zczj * 100',inplace=True)

    data.eval('srxjlhl = jyxjl / yysr * 100',inplace=True)

    data.eval('lrxjlhl = jyxjl / jlr * 100',inplace=True)

    data.eval('yszksrzb = yszk / yysr * 100',inplace=True)
    
    data.eval('bzzj = yszk+yfzk-yfkx-yskx',inplace=True)
    
    data.eval('bzzjzczb = bzzj / zczj / 100',inplace=True) #资产总计的单位是亿元

    data.eval('bzzjsrzb = bzzj / yysr * 100',inplace=True)
    
    data.eval('bzzjlrzb = bzzj / jlr * 100',inplace=True)

    data['rq'] = nf + '.12.31'
    data['dm'] = data.index

    '''保留2位小数必须在data=data.replace(np.nan,'--') 前执行'''
    data=data.round(2)

    #将应收账款、预付账款、预收款项、应付款项、被占资金及相关比例为0的设为np.nan
    values = {'yszk': 0, 'yfzk': 0, 'yskx': 0, 'yfkx': 0,'bzzj': 0, 
              'yszksrzb':0,'bzzjzczb': 0,'bzzjsrzb': 0,'bzzjlrzb': 0}

    data=data.replace(values,np.nan)

    xw.apps[0].quit()

    return data[['dm','rq','jyxjl','yysr','jlr','zczj','ldzc','jzc','sy','dxzw',
               'yszk','yfzk','yskx','yfkx','roe','ggbs','dxfzl','syjzczb','syzzczb',
               'srxjlhl','lrxjlhl','yszksrzb','bzzj','bzzjzczb','bzzjsrzb','bzzjlrzb']]

##########################################################################
#写入eps数据
##########################################################################
def write_eps():
    
    createDataBase()
    prefix = 'eps'
    
    dbfn=getdrive()+'\\hyb\\STOCKEPS.db'
    cn = sqlite3.connect(dbfn)

    config = iniconfig()
    nf1 = int(readkey(config,prefix + 'nf1'))
    nf2 = int(readkey(config,prefix + 'nf2'))

    newfn0 = readkey(config, prefix + 'fn')
    
    ddir=os.path.join(getdrive(),readkey(config,'dldir'))

    for j in range(nf1,nf2+1):

        newfn = newfn0+str(j)+'.xls'        
        xlsfn = os.path.join(ddir,newfn)

        epsdf=read_eps(xlsfn)

        data=epsdf.values.tolist()
        
        cn.executemany('INSERT OR REPLACE INTO GPEPS (GPDM,RQ,EPS0,EPS1,EPS0_G,EPS1_G) VALUES (?,?,?,?,?,?)', data)

        cn.commit()
        
    cn.close()    


##########################################################################
#写入财报数据
##########################################################################
def write_nb():

    createDataBase()
    prefix = 'rpt'
    dbfn=getdrive()+'\\hyb\\STOCKEPS.db'
    cn = sqlite3.connect(dbfn)

    config = iniconfig()
    nf1 = int(readkey(config,prefix + 'nf1'))
    nf2 = int(readkey(config,prefix + 'nf2'))

    newfn0 = readkey(config, prefix + 'fn')
    
    ddir=os.path.join(getdrive(),readkey(config,'dldir'))


    for j in range(nf1,nf2+1):
        
        newfn = newfn0+str(j)+'.xls'        
        xlsfn = os.path.join(ddir,newfn)

        df=read_rpt(xlsfn)

        data=df.values.tolist()
        
        cn.executemany('''INSERT OR REPLACE INTO GPNB 
                       (GPDM,RQ,JYXJL,YYSR,JLR,ZCZJ,LDZC,JZC,SY,DXZW,YSZK,YFZK,YSKX,YFKX,ROE,
                       GGBS,DXFZL,SYJZCZB,SYZZCZB,SRXJLHL,LRXJLHL,YSZKSRZB,BZZJ,BZZJZCZB,BZZJSRZB,BZZJLRZB)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', data)

        cn.commit()
        
    cn.close()    

##########################################################################
#元转成亿元
##########################################################################
def y2yy(num):
    try :
        return num/100000000
    except:
        return num

##########################################################################
#元转成万元
##########################################################################
def y2wy(num):
    try :
        return num/10000
    except:
        return num
    
##########################################################################
#写入growth数据
##########################################################################
def write_growth():
    
    createDataBase()
    prefix = 'czx'
    dbfn=getdrive()+'\\hyb\\STOCKEPS.db'
    cn = sqlite3.connect(dbfn)

    config = iniconfig()
    nf1 = int(readkey(config,prefix + 'nf1'))
    nf2 = int(readkey(config,prefix + 'nf2'))

    newfn0 = readkey(config, prefix + 'fn')
    
    ddir=os.path.join(getdrive(),readkey(config,'dldir'))

    for j in range(nf1,nf2+1):

        newfn = newfn0+str(j)+'.xls'        
        xlsfn = os.path.join(ddir,newfn)

        df=read_growth(xlsfn)

        data=df.values.tolist()
        
        cn.executemany('INSERT OR IGNORE INTO GPGROWTH (GPDM,RQ,YYZSR_G,YYSR_G,JLR_G,KFJLR_G,YYSR,JLR,FSY,SNJLR) VALUES (?,?,?,?,?,?,?,?,?,?)', data)

        cn.commit()
        
    cn.close()    

'''
python pandas 组内排序、单组排序、标号     
http://blog.csdn.net/qq_22238533/article/details/72395564    

pandas如何去掉、过滤数据集中的某些值或者某些行？ 
http://blog.csdn.net/qq_22238533/article/details/76127966

基于财务因子的多因子选股模型
https://www.windquant.com/qntcloud/v?3540281b-9a75-4506-adb6-983cf5091e74

选股条件：
2015年营业总收入同比增长率>20% 2016年营业总收入同比增长率>20% 
2016年销售毛利率>2015年销售毛利率 
2016年销售净利率>2015年销售净利率 
2017年业绩预增 上市时间在2016年1月以前 
2016年roe>10 2015年roe>10
'''

if __name__ == '__main__':

   #下载年报关键财务数据
#    dl_ths_xls('czx')
#    dl_ths_xls('eps')
#    dl_ths_xls('rpt')
#    write_eps()
#    write_growth()    
#   write_nb()



###############################################################################
#从同花顺i问财下载,prefix:eps、growth、rpt
###############################################################################
#def dl_ths_xls(prefix):
   
    prefix ='djd'
    config = iniconfig()
    
    ddir=os.path.join(getdrive(),readkey(config,'dldir'))
    dafn = dlfn(ddir)

    nf1 = int(readkey(config,prefix + 'nf1'))
    nf2 = int(readkey(config,prefix + 'nf2'))
    nb = readkey(config, prefix + 'rq')
    kw0 = readkey(config, prefix + 'kw')
    sele = readkey(config, prefix + 'sl')
    newfn0 = readkey(config, prefix + 'fn')

    username = readkey(config,'iwencaiusername')
    pwd = readkey(config,'iwencaipwd')
    
    profile = webdriver.FirefoxProfile()
    profile.set_preference('browser.download.dir', ddir)
    profile.set_preference('browser.download.folderList', 2)
    profile.set_preference('browser.download.manager.showWhenStarting', False)
    
    #http://www.w3school.com.cn/media/media_mimeref.asp
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/vnd.ms-excel')
    
    browser = webdriver.Firefox(firefox_profile=profile)

    #浏览器窗口最大化
    browser.maximize_window()
    #登录同花顺
    browser.get("http://upass.10jqka.com.cn/login")
    #time.sleep(1)
    elem = browser.find_element_by_id("username")
    elem.clear()
    elem.send_keys(username)
    
    elem = browser.find_element_by_class_name("pwd")
    elem.clear()
    elem.send_keys(pwd)
    
    browser.find_element_by_id("loginBtn").click()
    time.sleep(2)
    
    for j in range(nf1,nf2+1):

        kw = str(j) + nb + kw0
        newfn = newfn0 + str(j) + '.xls'
        newfn = os.path.join(ddir,newfn)
        
        if os.path.exists(newfn):
            os.remove(newfn)

        browser.get("http://www.iwencai.com/")
        time.sleep(5)
        browser.find_element_by_id("auto").clear()
        browser.find_element_by_id("auto").send_keys(kw)
        browser.find_element_by_id("qs-enter").click()
        time.sleep(10)
        
        #打开查询项目选单
        trigger = browser.find_element_by_class_name("showListTrigger")
        trigger.click()
        time.sleep(1)
        
        #获取查询项目选单
        checkboxes = browser.find_elements_by_class_name("showListCheckbox")
        indexstrs = browser.find_elements_by_class_name("index_str")
        
        #去掉选项前的“√”
        #涨幅、股价保留
        for i in range(0,len(checkboxes)):
            checkbox=checkboxes[i]
            
            #对于“pe,ttm”之类中间的逗号用下划线代替，注意配置文件也需要这样
            indexstr=indexstrs[i].text.replace(",","_")
            
            if checkbox.is_selected() and not indexstr in sele :
                checkbox.click()
            if not checkbox.is_selected() and indexstr in sele :
                checkbox.click()

        #向上滚屏
        js="var q=document.documentElement.scrollTop=0"  
        browser.execute_script(js)  
        time.sleep(3) 
        
        #关闭查询项目选单
        trigger = browser.find_element_by_class_name("showListTrigger")
        trigger.click()
        time.sleep(3)
        
        #导出数据
        elem = browser.find_element_by_class_name("export.actionBtn.do") 
        #在html中类名包含空格
        elem.click() 
        time.sleep(10)

        if os.path.exists(dafn):
            os.rename(dafn,newfn)
            
    browser.quit()
