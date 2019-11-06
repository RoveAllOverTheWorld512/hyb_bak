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
    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)
    """
    种类代码表：分类种类代码，分类种类名称
    """
    data=[["ZJH1","证监会行业门类代码"],
          ["ZJH2","证监会行业大类代码"],
          ["ZZ1","中证行业分类一级代码"],
          ["ZZ2","中证行业分类二级代码"],
          ["ZZ3","中证行业分类三级代码"],
          ["ZZ4","中证行业分类四级代码"],
          ["SW","申万行业分类代码"],
          ["TDX","通达信细分行业代码"],
          ["FG","通达信所属风格代码"],
          ["GN","通达信所属概念代码"],
          ["ZS","通达信所属指数代码"],
          ["DY","通达信地域分类代码"],
          ["THS","同花顺行业分类代码"]
          ]
    cn.execute('''CREATE TABLE IF NOT EXISTS ZLDM
           (ZLDM TEXT PRIMARY KEY,
           ZLMC TEXT);''')

    cn.executemany('INSERT OR IGNORE INTO ZLDM (ZLDM,ZLMC) VALUES (?,?)', data)
    
    """
    分类代码表：种类代码，分类代码、分类名称、个股数量
    """
    cn.execute('''CREATE TABLE IF NOT EXISTS FLDM
           (ZLDM TEXT,
           FLDM TEXT,
           FLMC TEXT,
           GPSL INTEGER);''')

    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS FLDM_ZLDM_FLDM ON FLDM(ZLDM,FLDM);''')

    """
    股票代码表：股票代码，股票名称
    """

    cn.execute('''CREATE TABLE IF NOT EXISTS GPDM
           (GPDM TEXT PRIMARY KEY,
           GPMC TEXT);''')

    """
    股票分类代码表：股票代码，股票名称，种类代码，分类代码
    """

    cn.execute('''CREATE TABLE IF NOT EXISTS GPFLDM
           (GPDM TEXT,
           ZLDM TEXT,
           FLDM TEXT);''')

    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPFLDM_GPDM_ZLDM_FLDM ON GPFLDM(GPDM,ZLDM,FLDM);''')

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

########################################################################
#获取本机通达信安装目录，生成自定义板块保存目录
#输入参数："gn"，"fg"，"zs"
#输出形如：    
#{'3D打印': ['3D打印',
#  39,
#  ['000928',
#   '000938',
#   '000969',
#   ......
#   '603167']],
# '黄金概念': ['黄金概念',
#  30,
#  ['000587',
#   '000975',
#   '002102',
#   ......
#   '601212',
#   '601899']]}
########################################################################
def gettdxblk(lb):

    blkfn = gettdxdir() + '\\T0002\\hq_cache\\block_'+lb+'.dat'
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

            tmp = f.read((400-stnum)*7)
        f.close()

    return blk


###############################################################################
#通达信概念、风格、指数
###############################################################################
def tdx_gnfgzs(lb):
    
    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)

    gn=gettdxblk(lb)
    
    gndm=[]
    gpgn=[]
    
    i=1
    
    for key in gn:
        gnmc,gpsl,gplst=gn[key]
#        ZLDM,FLDM,FLMC,GPSL
        gndm.append([lb.upper(),i,gnmc,gpsl])
#        GPDM,ZLDM,FLDM
        for gp in gplst:
            gpdm=gp+'.S'+('H' if gp[0]=='6' else 'Z')
            gpgn.append([gpdm,lb.upper(),i])
        
        i += 1
    
    cn.executemany('INSERT OR IGNORE INTO FLDM (ZLDM,FLDM,FLMC,GPSL) VALUES (?,?,?,?)', gndm)
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', gpgn)

    cn.commit()
    cn.close()

###############################################################################
#通达信概念、风格、指数
###############################################################################
def tdxgnfgzs():
    gfz=['gn','fg','zs']
    for lb in gfz:
        tdx_gnfgzs(lb)

###############################################################################
#从通达信系统读取股票公司所在地域
###############################################################################
def tdxdy():

    dydm="""
    黑龙江|1
    新疆|2
    吉林|3
    甘肃|4
    辽宁|5
    青海|6
    北京|7
    陕西|8
    天津|9
    广西|10
    河北|11
    广东|12
    河南|13
    宁夏|14
    山东|15
    上海|16
    山西|17
    深圳|18
    湖北|19
    福建|20
    湖南|21
    江西|22
    四川|23
    安徽|24
    重庆|25
    江苏|26
    云南|27
    浙江|28
    贵州|29
    海南|30
    西藏|31
    内蒙|32
    """
    dydm=dydm.replace('|','\t')
    p = '(.*?)\t(.*?)\n'
    dylst=re.findall(p,dydm)
    dylst=[['DY',a[1].strip(),a[0].strip()] for a in dylst]

    fn=gettdxdir()+"\\T0002\\hq_cache\\base.dbf"
    gpdy = dbf2pandas(fn,['gpdm', 'dy']) 
    #去掉指数和基金，这些代码的dy=0
    gpdy=gpdy[gpdy['dy'].map(lambda x:eval(x)>0)]

    gpdy['gpdm']=gpdy['gpdm'].map(lambda x: x+('.SH' if x[0]=='6' else '.SZ'))

    #先转成list
    gpszdy=[[a[0],'DY',a[1].strip()] for a in gpdy.values.tolist()]

    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)

    cn.executemany('INSERT OR IGNORE INTO FLDM (ZLDM,FLDM,FLMC) VALUES (?,?,?)', dylst)
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', gpszdy)

    cn.commit()
    cn.close()

        
###############################################################################
#同花顺行业分类
###############################################################################
def thshy():
    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)
    thshyfn=ths_hy_xls()
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
#同花顺花行业下载
###############################################################################
def ths_hy_xls():

    config = iniconfig()
    ddir=os.path.join(getdrive(),readkey(config,'dldir'))
    dafn = dlfn(ddir)

    username = readkey(config,'iwencaiusername')
    pwd = readkey(config,'iwencaipwd')
    kw = readkey(config,'iwencaikw')
    sele = readkey(config,'iwencaiselestr')
    
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

    #关闭浏览器
    browser.quit()

    if os.path.exists(dafn):
        newfn = os.path.splitext(dafn)[0]+"_thshy.xls"

        if os.path.exists(newfn):
            os.remove(newfn)

        os.rename(dafn,newfn)
        return newfn

    else:

        return ""

########################################################################
#根据通达信新行业或申万行业代码提取股票列表
########################################################################
def ggtdx_swhy():

    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)

    p = '(\d{6})\t(.+)\t(.*?)\r\n'
    zxgfn = gettdxdir()+r'T0002\hq_cache\tdxhy.cfg'
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
    zxg=zxg.replace('|','\t')
    zxglst =re.findall(p,zxg)

    data=[[gpdm+'.S'+('H' if gpdm[0]=='6' else 'Z'),'TDX',tdxnhy] for gpdm,tdxnhy,swhy in zxglst if (gpdm[0] == '6' or gpdm[0:2] in ('30','00'))]
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', data)

    data=[[gpdm+'.S'+('H' if gpdm[0]=='6' else 'Z'),'SW',swhy] for gpdm,tdxnhy,swhy in zxglst if (gpdm[0] == '6' or gpdm[0:2] in ('30','00'))]
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', data)

    cn.commit()
    cn.close()

########################################################################
#读取通达信、申万行业代码incon.dat
########################################################################
def tdx_swhydm():
    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)

    hyflfn = gettdxdir()+'incon.dat'
    with open(hyflfn,'rb') as dtf:
        hyfl = dtf.read()
        if hyfl[:3] == b'\xef\xbb\xbf' :
            hyfl = hyfl.decode('UTF8','ignore')   #UTF-8
        elif hyfl[:2] == b'\xfe\xff' :
            hyfl = hyfl.decode('UTF-16','ignore')  #Unicode big endian
        elif hyfl[:2] == b'\xff\xfe' :
            hyfl = hyfl.decode('UTF-16','ignore')  #Unicode
        else :
            hyfl = hyfl.decode('GBK','ignore')      #ansi编码


    p='#TDXNHY([\s\S]*?)######'      
    hystr =re.findall(p,hyfl)
    
    hystr=hystr[0].replace('|','\t')
    p = '(.*?)\t(.*?)\r\n'
    tdxhy =re.findall(p,hystr)
    tdxhy=[['TDX',a[0],a[1]] for a in tdxhy]
    cn.executemany('INSERT OR IGNORE INTO FLDM (ZLDM,FLDM,FLMC) VALUES (?,?,?)', tdxhy)

    p='#SWHY([\s\S]*?)######'      
    hystr =re.findall(p,hyfl)
    
    hystr=hystr[0].replace('|','\t')
    p = '(.*?)\t(.*?)\r\n'
    swhy =re.findall(p,hystr)
    swhy=[['SW',a[0],a[1]] for a in swhy]
    cn.executemany('INSERT OR IGNORE INTO FLDM (ZLDM,FLDM,FLMC) VALUES (?,?,?)', swhy)

    cn.commit()
    cn.close()

########################################################################
#读取证监会分类
########################################################################
def zjhfl():
    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)
    sylfn=syl_pe_fn('pedir')
    print(sylfn)
    wb = xw.Book(sylfn)
    sht = wb.sheets['证监会行业静态市盈率']
    data=sht.range("A2:D2").options(expand='down').value

    data=[[("ZJH1" if len(a[0])==1 else "ZJH2"),a[0],a[1],a[3]] for a in data]
    cn.executemany('INSERT OR IGNORE INTO FLDM (ZLDM,FLDM,FLMC,GPSL) VALUES (?,?,?,?)', data)
 
    sht = wb.sheets['个股数据']
    data=sht.range("A2:E2").options(expand='down').value

    xw.apps[0].quit()

    #股票代码表
#    data1=[[a[0]+'.S'+('H' if a[0][0]=='6' else 'Z' ),a[1]] for a in data]
#    cn.executemany('INSERT OR IGNORE INTO GPDM (GPDM,GPMC) VALUES (?,?)', data1)
    #股票门类
    data1=[[a[0]+'.S'+('H' if a[0][0]=='6' else 'Z' ),"ZJH1",a[2]] for a in data]
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', data1)
    #股票大类
    data1=[[a[0]+'.S'+('H' if a[0][0]=='6' else 'Z' ),"ZJH2",a[4]] for a in data]
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', data1)

    cn.commit()
    cn.close()
    
########################################################################
#读取中证分类
########################################################################
def zzfl():
    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)
    sylfn=syl_pe_fn('syldir')
    wb = xw.Book(sylfn)
    print(sylfn)
    sht = wb.sheets['中证行业静态市盈率']
    data=sht.range("A2:D2").options(expand='down').value

    data1=[]
    for a in data:
        if len(a[0])==2 :
            data1.append(['ZZ1',a[0],a[1],a[3]])
        if len(a[0])==4 :
            data1.append(['ZZ2',a[0],a[1],a[3]])
        if len(a[0])==6 :
            data1.append(['ZZ3',a[0],a[1],a[3]])
        if len(a[0])==8 :
            data1.append(['ZZ4',a[0],a[1],a[3]])
        
    cn.executemany('INSERT OR IGNORE INTO FLDM (ZLDM,FLDM,FLMC,GPSL) VALUES (?,?,?,?)', data1)
 
    sht = wb.sheets['个股数据']
    data=sht.range("A2:I2").options(expand='down').value

    xw.apps[0].quit()
    
    data1=[[a[0]+'.S'+('H' if a[0][0]=='6' else 'Z' ),"ZZ1",a[2]] for a in data]
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', data1)
    
    data1=[[a[0]+'.S'+('H' if a[0][0]=='6' else 'Z' ),"ZZ2",a[4]] for a in data]
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', data1)

    data1=[[a[0]+'.S'+('H' if a[0][0]=='6' else 'Z' ),"ZZ3",a[6]] for a in data]
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', data1)

    data1=[[a[0]+'.S'+('H' if a[0][0]=='6' else 'Z' ),"ZZ4",a[8]] for a in data]
    cn.executemany('INSERT OR IGNORE INTO GPFLDM (GPDM,ZLDM,FLDM) VALUES (?,?,?)', data1)

    cn.commit()
    cn.close()
    
#########################################################################
#从配置文件中读取休市日期
#########################################################################
def readclosedate(config):
    keys = config.keys()
    if keys.count('stockclosedate') :
        return eval(config['stockclosedate'])
    else :
        return []

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
def DeleTables():    
    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
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
#查询股票所属行业、概念、风格、地域的信息
##########################################################################
def Query(gpdm):
    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
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
    
    dbfn=getdrive()+'\\hyb\\STOCKHY.db'
    cn = sqlite3.connect(dbfn)
    gpdmb=get_gpdm()
    data=[[a[0],a[1]] for a in gpdmb.values.tolist()]
    cn.executemany('INSERT OR IGNORE INTO GPDM (GPDM,GPMC) VALUES (?,?)', data)

    cn.commit()
    cn.close()    
    
##########################################################################
#下载股票质押数据文件
##########################################################################
def dl_gpzy():
    today=datetime.datetime.now().strftime("%Y.%m.%d")

    config = iniconfig()
    ddir=os.path.join(getdrive(),readkey(config,'dldir'))
    
    profile = webdriver.FirefoxProfile()
    profile.set_preference('browser.download.dir', ddir)
    profile.set_preference('browser.download.folderList', 2)
    profile.set_preference('browser.download.manager.showWhenStarting', False)

    #http://www.w3school.com.cn/media/media_mimeref.asp
    #http://blog.csdn.net/kmanzxbin/article/details/78329751
    #实际Content-Type类型却是application/x-msdownload
    profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 
                           'application/x-msdownload, application/octet-stream, application/vnd.ms-excel, text/csv, application/zip')
        
    browser = webdriver.Firefox(firefox_profile=profile)

    url='http://www.chinaclear.cn/cms-rank/queryPledgeProportion?queryDate='+today+'&secCde='
    browser.get(url)
    time.sleep(3)
    browser.find_element_by_xpath('//div[1]/a').click()
    
    browser.quit()


def thisweek(str_date):
    try:
        #尝试将参数转换成为datetime.date格式，1是方便后面的日期加减，2是验证日期是否有效。
        date_input = datetime.date.fromtimestamp(time.mktime(time.strptime(str_date,"%Y-%m-%d")))
    except:
        raise '参数错误：错误的日期，期待值2016-01-01格式'
        
    n = datetime.datetime.weekday(date_input)
    weeklist = []
    for i in range(7):
        this_day=date_input  + datetime.timedelta(0-n+i)
        weeklist.append([i,this_day])
        
    return weeklist

    


def hy():
#    DeleTables()    
#    createDataBase()
#    gpdmtbl()
#    zjhfl() 
#    zzfl() 
#    tdx_swhydm()
#    ggtdx_swhy() 
#    thshy()
#    tdxgnfgzs()
#    tdxdy()

    gpdm='600410.SH'
    data=Query(gpdm)
    return data

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


if __name__ == '__main__':
    
#    xlsfn=gpzyfn()
#    if not os.path.exists(xlsfn):
#        dl_gpzy()
#        
#    wb = xw.Book(xlsfn)
#    sht = wb.sheets['Sheet1']
#    data=sht.range("B4:H4").options(expand='down').value
#    
#    wb.close()
# 
    data= hy()

#    print(syl_pe_fn('syldir'))