# -*- coding: utf-8 -*-
"""
Created on Thu Feb  9 16:18:24 2017
http://www.jdon.com/idea/matplotlib.html
@author: Lenovo
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Jan 20 11:27:04 2017

@author: lenovo


"""
import re
import os
import sys
import getopt
import struct
import unicodedata
import requests
import zipfile
import datetime
import time
from dateutil import parser
import xlrd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import pandas as pd
from configobj import ConfigObj
import statsmodels.api as sm
import xlwt
import winreg

ezxf= xlwt.easyxf

def writesheet(book_name,sheet_name, headings, data, heading_xf, data_xfs,fore_red_com, width_xfs):
    sheet = book_name.add_sheet(sheet_name)
    rowx = 0
    for colx, value in enumerate(headings):
        sheet.write(rowx, colx, value,heading_xf)

    sheet.set_panes_frozen(True)        # 冻结窗口
    sheet.set_horz_split_pos(1)         # 冻结行数
    sheet.set_vert_split_pos(3)         #冻结列数
    sheet.set_remove_splits(True)       # 使用冻结窗口不能分屏

    coln,num = fore_red_com
    for row in data:
        rowx += 1

        for colx, value in enumerate(row):
            if row[coln] > 0:
                sheet.write(rowx, colx, value,data_xfs[colx][1])
            else :
                sheet.write(rowx, colx, value,data_xfs[colx][0])

    for colx, width in enumerate(width_xfs):
        sheet.col(colx).width = 256*width

def write_xls(xlsfile,gg):

    hdngs = ['序号','股票代码','股票名称','标的代码','标的名称','起始时间','截止时间','beta','alpha','样本数量','R平方','adj_R平方']
    kinds = 'cint ctxt ctxt ctxt ctxt ctxt ctxt flt flt int flt flt'.split()
    widths= 'wd1  wd2  wd2  wd2  wd2  wd2  wd2  wd3 wd3 wd2 wd3 wd3'.split()
    heading_xf = ezxf('font: bold on; align:wrap on, vert centre, horiz center')

    kind_to_xf_map = {
        'cint': [ezxf('align:horiz center',num_format_str='#0'),
                 ezxf('pattern: pattern solid,fore_colour red;align:horiz center',num_format_str='#0')],
        'int': [ezxf(num_format_str='#0'),
                ezxf('pattern: pattern solid,fore_colour red',num_format_str='#0')],
        'flt': [ezxf(num_format_str='#0.00000000'),
                ezxf('pattern: pattern solid,fore_colour red',num_format_str='#0.00000000')],
        'text': [ezxf(),
                 ezxf('pattern: pattern solid,fore_colour red')],
        'ctxt': [ezxf('align:horiz center'),
                 ezxf('pattern: pattern solid,fore_colour red;align:horiz center')],
        }
    data_xfs = [kind_to_xf_map[k] for k in kinds]
    fore_red_con = [8,0]      #第8列大于0

    width_to_xf_map = {
        'wd1':6,
        'wd2':10,
        'wd3':16,
        'wd4':30,
        }
    width_xfs = [width_to_xf_map[k] for k in widths]

    book = xlwt.Workbook()
    writesheet(book,'个股数据', hdngs, gg, heading_xf, data_xfs,fore_red_con, width_xfs)
    book.save(xlsfile)

##########################################################################
#将字符串转换为时间戳，不成功返回None
##########################################################################
def str2datetime(s):
    try:
        dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
    except :
        dt = None
    return dt

##########################################################################
#n天后日期串，不成功返回None
##########################################################################
def nextdtstr(s,n):
    dt = str2datetime(s)
    if dt :
        dt += datetime.timedelta(n)
        return dt.strftime("%Y%m%d")
    else :
        return None

########################################################################
#初始化本程序配置文件
########################################################################
def iniconfig():
    myname=filename(sys.argv[0])
    wkdir = os.getcwd()
    inifile = os.path.join(wkdir,myname+'.ini')  #设置缺省配置文件
    return ConfigObj(inifile,encoding='GBK')

#########################################################################
#读INI文件
#########################################################################
def readini(inifile):
    config = ConfigObj(inifile,encoding='GBK')
    return config

#########################################################################
#读取键值
#########################################################################
def readkey(config,key):
    keys = config.keys()
    if keys.count(key) :
        return config[key]
    else :
        return ""

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def is_float_by_except(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

#############################################################################
#读取EXCEL表单季度EPS
#############################################################################
def read_xls(file,sheet,rowbg):
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name(sheet)
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    nyear = int((ncols-2)/4)

    data =[]
    for rownum in range(rowbg,nrows):
        row = table.row_values(rownum)
        gpdm = row[0][:6]
        gpmc = row[1]
        bj = False
        for yearn in range(1990,1990+nyear):
            dts = ['0331','0630','0930','1231']
            for j in range(4):
                date = str(yearn)+dts[j]
                value = row[(yearn-1990)*4+2+j]
                if isinstance(value,float) :
                    bj = True

                if bj :
                    if isinstance(value,float) :
                        data.append([gpdm,gpmc,date,value])
                    else :
                        data.append([gpdm,gpmc,date,0])

    return data

#############################################################################
#返回包含中文的byte字符串转的长度(一个汉字的长度为2)
#############################################################################
def str_width(s):
    w=0
    for c in s:
        if (unicodedata.east_asian_width(c) in ('F','W')):
            w +=2
        else:
            w +=1
    return(w)

#############################################################################
#将包含中文的byte字符串转变为指定长度（一个汉字为2个宽度,后面用空格补齐)
#############################################################################
def cnstrjust(cnstr,length):
    cnstrw=str_width(cnstr)
    if cnstrw>length :
        i=0
        while i<len(cnstr):
            i += 1
            cutstr = cnstr[:i]
            if str_width(cutstr)>length :
                break
        cnstr = cutstr[:i-1]
        cnstrw=str_width(cnstr)

    return cnstr+" "*(length-cnstrw)

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
                value = value.strip(b'\x00').decode('GBK')
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


#############################################################################
#检查字段类型和宽度，如果N，C型宽度不够，则扩展宽度，如果D，L型宽度不符则改为C型
#############################################################################
def checkdata(fieldnames, fieldspecs, records):
    flds = []
    for name, (typ, size, deci) in list(zip(fieldnames, fieldspecs)):
        i = fieldnames.index(name)
        if typ=="N" :
            p="{:>"+str(size)+"."+str(deci)+"f}"
            maxlen = max([len(p.format(e[i])) for e in records])
            if maxlen>size :
                size = maxlen
        if typ=="C" :
            maxlen = max([len(e[i]) for e in records])
            if maxlen>size :
                size = maxlen
        if typ in ("D","L"):
            maxlen = max([len(e[i]) for e in records])
            minlen = min([len(e[i]) for e in records])
            if maxlen!=8 or minlen!=8 :
                typ = 'C'
                size = maxlen
        flds.append([typ,size,deci])
    return flds


#############################################################################
#写dbf文件
#############################################################################
def dbfwriter(f, fieldnames, fieldspecs, records):
    #对数据与字段的类型和宽度进行检查、优化
    fieldspecs = checkdata(fieldnames, fieldspecs, records)
    # 文件头部信息
    ver = 3
    now = datetime.datetime.now()
    yr, mon, day = now.year-2000, now.month, now.day
    numrec = len(records)
    numfields = len(fieldspecs)
    lenheader = numfields * 32 + 33
    lenrecord = sum(field[1] for field in fieldspecs) + 1
    codepageid = 122
    #Code Pages Supported by Visual FoxPro:936Chinese (PRC, Singapore) Windows
    #https://technet.microsoft.com/zh-cn/learning/aa975345
    hdr = struct.pack('<BBBBLHH17xB2x', ver, yr, mon, day, numrec, lenheader, lenrecord, codepageid)
    f.write(hdr)

    # 字段名信息
    addr = 1
    for name, (typ, size, deci) in list(zip(fieldnames, fieldspecs)):
        name = name.ljust(11, '\x00').encode('GBK')
        typ = typ.encode('GBK')
        fld = struct.pack('<11sciBB14x', name, typ, addr, size, deci)
        addr += size
        f.write(fld)

    # 终止符
    f.write('\r'.encode())

    # 记录
    for record in records:
        f.write(' '.encode())                        # deletion flag
        for (typ, size, deci), value in list(zip(fieldspecs, record)):
            if typ == "C":
                value = cnstrjust(value,size)
            if typ == "N":
                p="{:>"+str(size)+"."+str(deci)+"f}"
                value = p.format(value)

            if typ == 'D':
                value = value.ljust(8, ' ')
            if typ == 'L':
                value = value.upper()

            f.write(value.encode("GBK"))

    # 文件尾
    f.write('\x1A'.encode())

def day2pandas(dayfn,start=None,end=None):
    columns = ['date','open', 'high', 'low','close','amout','volume','rate','pre_close','adj_rate','adj_close']

    with open(dayfn,"rb") as f:
        data = f.read()
        f.close()
    days = int(len(data)/32)
    records = []
    qsp = 0
    for i in range(days):
        dat = data[i*32:(i+1)*32]
        rq,kp,zg,zd,sp,cje,cjl,tmp = struct.unpack("iiiiifii", dat)
        rq = str2datetime(str(rq))
        kp = kp/100.00
        zg = zg/100.00
        zd = zd/100.00
        sp = sp/100.00
        cje = cje/100000000.00     #亿元
        cjl = cjl/10000.00         #万股
        zf = sp/qsp-1 if i>0 else 0.0
        records.append([rq,kp,zg,zd,sp,cje,cjl,zf,qsp,zf,sp])
        qsp = sp

    df = pd.DataFrame(records,columns=columns)
    df = df.set_index('date')
    start = str2datetime(start)
    end = str2datetime(end)

    if start == None or end==None :
        return df
    else :
        return df[start:end]



def day2dbf(dayfn,dbffn):
    fieldnames = ['date','open', 'high', 'low','close','amout','volume','rate','pre_close','adj_rate','adj_close']
    fieldspecs = [('D', 8, 0),('N', 8, 2),('N', 8, 2),('N', 8, 2),('N', 8, 2),
                  ('N', 10, 2),('N', 10, 2),('N', 12, 8),('N', 8, 2),('N', 12, 8),('N', 8, 2)]

    with open(dayfn,"rb") as f:
        data = f.read()
        f.close()
    days = int(len(data)/32)
    records = []
    qsp = 0
    for i in range(days):
        dat = data[i*32:(i+1)*32]
        rq,kp,zg,zd,sp,cje,cjl,tmp = struct.unpack("iiiiifii", dat)
        rq = str(rq)
        kp = kp/100.00
        zg = zg/100.00
        zd = zd/100.00
        sp = sp/100.00
        cje = cje/100000000.00     #亿元
        cjl = cjl/10000.00         #万股
        zf = sp/qsp-1 if i>0 else 0.0
        records.append([rq,kp,zg,zd,sp,cje,cjl,zf,qsp,zf,sp])
        qsp = sp

    with open(dbffn,"wb") as f:
        dbfwriter(f, fieldnames, fieldspecs, records)
        f.close()

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

def fhpg2pandas(dm):
    fhpgfn = getdisk()+r'\tdx\dbf\fhpg.dbf'
    fhpg = dbf2pandas(fhpgfn,['gpdm', 'gqdjr', 'mgfh', 'mgsg'])
    fhpg['date'] = fhpg['gqdjr'].map(str2datetime)
    fhpg = fhpg.set_index(['gpdm','date'])

    try :
        return fhpg.ix[dm,'gqdjr':'mgsg']
    except :
        return []

def getdisk():
    return sys.argv[0][:2]

def makedir(dirname):
    if dirname == None :
        return False

    if not os.path.exists(dirname):
        try :
            os.mkdir(dirname)
            return True
        except(OSError):
            print("创建目录%s出错，请检查！" % dirname)
            return False
    else :
        return True

def filename(pathname):
    wjm = os.path.splitext(os.path.basename(pathname))
    return wjm[0]

###############################################################################
#获取最新交易日，如果当天是交易日，在18:00后用当天，如果当天不是交易日
###############################################################################
def lastday():
    config = iniconfig()
    stockclosedate = readkey(config,'stockclosedate')
    now = datetime.datetime.now()
    td = now.strftime("%Y%m%d") #今天
    hr = now.strftime("%H") #今天
    if hr<'18' :
        td = nextdtstr(td,-1)

    wk = str2datetime(td).weekday()
    if wk<5 and not td in stockclosedate :
        return td
    else :
        while True :
            td = nextdtstr(td,-1)
            wk = str2datetime(td).weekday()
            if wk<5 and not td in stockclosedate :
                return td


def dlday():
    #每天下载一次
    #http://www.tdx.com.cn/products/data/data/vipdoc/shlday.zip
    #http://www.tdx.com.cn/products/data/data/vipdoc/szlday.zip

    url0 = "http://www.tdx.com.cn/products/data/data/vipdoc/"
    fnls = ["shlday.zip","szlday.zip"]
    svdir = getdisk()+"\\tdx"
    if not os.path.exists(svdir) :
        makedir(svdir)

    for fn in fnls:

        dlyes = False    #下载标志，True表示要下载
        zip_file = svdir + "\\" + fn
        url = url0 + fn
        if os.path.exists(zip_file):
            mtime=os.path.getmtime(zip_file)  #文件建立时间
            ltime=time.strftime("%Y%m%d",time.localtime(mtime))
            if ltime >= lastday() :
                dlyes = False
            else :
                dlyes = True
        else :
            dlyes = True


        if dlyes:
            print ("正在下载的文件%s，请等待！" % zip_file)

            r = requests.get(url)
            #如果下载文件不存在 ，r返回 <Response [404]>， r.ok为False
            #如果下载文件存在 ，r返回 <Response [200]>，r.ok为True
            if not r.ok :
                print ("你所下载的文件%s不存在！" % zip_file)

            else :
                os.remove(zip_file)
                with open(zip_file, "wb") as f:
                    f.write(r.content)
                    f.close()


        if dlyes and os.path.exists(zip_file):
            print ("正在解压文件%s，请等待！" % zip_file)
            extdir = svdir + "\\" + fn[:6]
            f_zip = zipfile.ZipFile(zip_file, 'r')
            f_zip.extractall(extdir)
            f_zip.close()

def topandas(gpdm,start=None,end=None):

    sc = 'sh' if gpdm[0]=='6' else 'sz'
    dayfn =getdisk()+'\\tdx\\'+sc+'lday\\'+sc+gpdm+'.day'
    if os.path.exists(dayfn) :
        return day2pandas(dayfn,start,end)
    else :
        return []

################################################################################
#提取DataFrame时间索引指定日期date前n个日期，返回日期
################################################################################
def df_timeindex_to_datelist(df):
    dfti = df.index
    dftia = np.vectorize(lambda s: s.strftime('%Y%m%d'))(dfti.to_pydatetime())
    return dftia.tolist()

################################################################################
#提取DataFrame时间索引指定日期date前n个日期，返回日期
################################################################################
def df_next_date(df,date,n=0):
    dftilst = df_timeindex_to_datelist(df)
    dftilst.sort()
    tmin = str2datetime(dftilst[0])
    tmax = str2datetime(dftilst[len(dftilst)-1])
    t = str2datetime(date)
    if t< tmin or t>tmax :
        return None

    try :
        i = dftilst.index(date)
        if i+n<0 or i+n>=len(dftilst) :
            return None
        else :
            return dftilst[dftilst.index(date)+n]
    except :
        while True :
            date = nextdtstr(date,-1)
            if date in dftilst :
                return date



###############################################################################
#前复权收盘价
###############################################################################
def adj_close(df,fhpg):
    for i in range(len(fhpg)):
        date, mgfh, mgsg = fhpg.iloc[i]

        fqyes = False     #如果股权登记日不在数据范围内则不能进行复权处理

        if len(df.ix[date:date])==1 :      #股权登记日存在交易
            fqyes = True
        else :
            date = df_next_date(df,date,-1) #股权登记日不存在交易则前找交易日
            if len(df.ix[date:date])==1 :   #股权登记日前有交易则进行复权
                fqyes = True

        if fqyes :
            oldclose = df.ix[date,'adj_close']
            newclose = (oldclose - mgfh)/(1+mgsg)
            df.ix[date,'adj_close'] = newclose
            nextdate = df_next_date(df,date,1)
            if nextdate == None :
                break
            df.ix[nextdate,'pre_close'] = newclose
            df.ix[nextdate,'adj_rate'] =  df.ix[nextdate,'close']/df.ix[nextdate,'pre_close']- 1

    ti = df_timeindex_to_datelist(df)
    ti.reverse()
    for i in range(len(ti)):
        date = ti[i]
        if i== 0 :
            df.ix[date,'adj_close'] = df.ix[date,'adj_close']
        else :
            df.ix[date,'adj_close'] = next_close /(1+next_rate)

        next_close = df.ix[date,'adj_close']
        next_rate = df.ix[date,'adj_rate']

    return df

def beta(stocklst,start,end):
    marketlst = ['399300','399101','399102']
    mktm = {'399300':'沪深300','399101':'中小板综','399102':'创业板综'}
    dfxd = {}
    for market in marketlst :
        dfxd[market] = topandas(market,start,end)

    for stock in stocklst :
        print('共有%d只股票，正在处理第%d只股票:%s，请等待。' %
              (len(stocklst),stocklst.index(stock)+1,stock))
        if stock[0] == '3' :
            ml = ['399300','399102']
        elif stock[:3] == "002" :
            ml = ['399300','399101']
        else :
            ml = ['399300']

        dfy=topandas(stock,start,end)
        if len(dfy)==0 :
            continue

        fh = fhpg2pandas(stock)

        if len(fh)>0 :
            adj_close(dfy,fh)

        for market in ml :

            dfx = dfxd[market]

            daily_return = pd.merge(dfx,dfy,left_index = True, right_index = True)
            daily_return = daily_return[['adj_rate_x','adj_rate_y']]
            daily_return["intercept"]=1.0
            model = sm.OLS(daily_return["adj_rate_y"],daily_return[["adj_rate_x","intercept"]])
            results = model.fit()

            rs = list(results.params)
            rs.append(results.nobs)
            rs.append(results.rsquared)
            rs.append(results.rsquared_adj)
            rs.insert(0,end)
            rs.insert(0,start)
            rs.insert(0,mktm[market])
            rs.insert(0,market)
            rs.insert(0,stock)
            rs.insert(0,stocklst.index(stock)+1)
            yield rs

def Usage():
    print ('用法:')
    print ('-h, --help: 显示帮助信息。')
    print ('-v, --version: 显示版本信息。')
    print ('-i, --input: 股票列表文本文件。')

def Version():
    print ('版本 2.0.0')

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

            tmp = f.read((400-stnum)*7)
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

########################################################################
#股票代码表
########################################################################
def gpdmdict():
    fn = getdisk()+'\\tdx\\gpdmb.txt'
    with open(fn) as f:
        gpdmb = f.read()
        f.close()

    dmb = re.findall('(\d{6})\t(.+)\n',gpdmb)
    dm = {}
    for (gpdm,gpmc) in dmb :
        dm[gpdm] = gpmc

    return dm


########################################################################
#根据通达信新行业或申万行业代码提取股票列表
########################################################################
def hy(hydms):
    hylb = re.findall('(.+)_',hydms)[0].lower()
    hydm = re.findall('_(.+)',hydms)[0].upper()
    dmcd = len(hydm)
    if hylb not in ['tdx','sw'] :
        print('参数不对！')
        return None

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

    if hylb == 'tdx' :
        zxglb = [gpdm for gpdm,tdxnhy,swhy in zxglst if tdxnhy[:dmcd] == hydm]
    else :
        zxglb = [gpdm for gpdm,tdxnhy,swhy in zxglst if swhy[:dmcd] == hydm]

    return zxglb


########################################################################
#提取文件名
########################################################################
def flnm(pathname):
    wjm = os.path.splitext(os.path.basename(pathname))
    return wjm[0]

########################################################################
#主程序
########################################################################
def main(argv):

    myname=flnm(argv[0])
    wkdir = os.getcwd()
    inifile = os.path.join(wkdir,myname+'.ini')  #设置缺省配置文件
    config = iniconfig()
    tdxblkdir = readkey(config,'tdxblkdir')

    try:
        opts, args = getopt.getopt(argv[1:], 'hvi:k:d:',
                       ['help','version','kind=','inpute=','dir='])
    except (getopt.GetoptError):
        Usage()
        sys.exit(1)

    bklb = "zd"
    zxgfile = ""
    imgdir = ''
    for o, a in opts:
        if o in ('-h', '--help'):
            Usage()
            sys.exit(0)
        elif o in ('-v', '--version'):
            Version()
            sys.exit(0)
        elif o in ('-i', '--input'):
            zxgfile = a
        elif o in ('-k', '--kind'):
            imgdir = a
            bkxx = a
            bklb = re.findall('(.+)_',bkxx)[0].lower()
        elif o in ('-d', '--dir'):
            imgdir = a
        else:
            print ('无效参数！')
            sys.exit(3)

    if bklb not in ['fg','gn','zs','tdx','sw','zd'] :
        print('板块类别参数不对，请查查。')
        sys.exit(3)

    if not os.path.exists(inifile) :
        print("配置文件%s不存在，无法运行，请检查。" % inifile)
        sys.exit(3)

    if bklb in ['tdx','sw'] :
        zxglb = hy(bkxx)

    if bklb in ['fg','gn','zs'] :
        bklb = re.findall('(.+)_',bkxx)[0].lower()
        bkjc = re.findall('_(.+)',bkxx)[0].upper()
        bkinfo = gettdxblk(bklb)
        try :
            zxglb = bkinfo[bkjc][2]
        except :
            zxglb = []
    if bklb == 'zd' :
        imgdir = 'zd_'+flnm(zxgfile)
        if len(zxgfile)==0 :
            zxgfile = "zxg.blk"          #没有指定股票列表就用通达信自选股板块
        if zxgfile.upper().endswith(".BLK") or zxgfile.upper().endswith(".EBK") :
            tdxblkdir = gettdxblkdir()
            zxgfile = os.path.join(tdxblkdir,zxgfile)
            zxglb = zxglist(zxgfile,"tdxblk")
        else:
            zxglb =  zxglist(zxgfile)

    if len(zxglb)>0 :
        print("正在准备数据，估计需要3分钟请等待……")
        pe = pe2pandas()
    else :
        sys.exit()

    if not makedir(imgdir) :
        sys.exit()

    i=0
    for gpdm in zxglb :
        i += 1
        try :
            gpmc = gpdmdict()[gpdm]
            ggpe = pe.loc[gpdm]
        except:
            continue
        print('共有%d只股票，正在处理第%d只股票：%s%s，请等待。' % (len(zxglb),i,gpdm,gpmc))

        ggpe = ggpe.set_index('date')
        generateimg(ggpe,gpdm,imgdir)

########################################################################
#将滚动市盈率读入pandas
########################################################################
def pe2pandas():
    with open('pettm.dbf',"rb") as f:
        data = list(dbfreader(f))
        f.close()
    columns = data[0]
    columns=[e.lower() for e in columns]
    data = data[2:]
    df = pd.DataFrame(np.array(data),columns=columns)
    df['date'] = df['date'].map(str2datetime)
    df = df.set_index('gpdm')
    return df.ix[:,['date','pe_ttm']]



########################################################################
#读取个股市盈率
########################################################################
def ggsyl(file,sheet,gpdm):
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name(sheet)
    nrows = table.nrows #行数

    data =[]

    for rownum in range(1,nrows):
        row = table.row_values(rownum)
        if row[1]==gpdm:
            if row[12] != '-' :
                data.append([row[16], row[12]])
            else :
                data.append([row[16], 0])

    return data


def generateimg(pe,stock,imgdir):
    pe['pe_ttm'] = pe['pe_ttm'].astype('float')

    dfy=topandas(stock)
    fh = fhpg2pandas(stock)
    if len(fh)>0 :
        adj_close(dfy,fh)

    try :
        xsmd = pd.merge(dfy,pe,left_index = True, right_index = True)
        xsmd.eval('eps=adj_close/pe_ttm')
    except :
        return

    font = FontProperties(fname=r"c:\windows\fonts\simhei.ttf", size=14)

    fig, ax1 = plt.subplots(figsize=(18,6))
    ax2 = ax1.twinx()
    ax3 = ax1.twinx()

    ax1.plot(xsmd.index,xsmd['pe_ttm'],color="blue",linewidth=1.5,label='pe(ttm)')
    ax2.plot(xsmd.index,xsmd['adj_close'],color="red",linewidth=1.5,label='adj_close')
    ax3.plot(xsmd.index,xsmd['eps'],color="green",linewidth=1.5,label='eps(ttm)')

    title = stock+ gpdmdict()[stock] + "股价、市盈率及每股收益走势图"
    fig.suptitle(title, fontproperties=font,fontsize = 14, fontweight='bold')

    ax1.set_xlabel('日期', fontproperties=font,fontsize = 16)
    ax1.set_ylabel('滚动市盈率', color="blue", fontproperties=font, fontsize = 16)
    legend = ax1.legend(loc='upper left',  fontsize=16)
    ax1.set_ylim(0,min(200,max(xsmd['pe_ttm'])))
    ax1.tick_params('y', colors='b')
    ax1.grid(True)

    ax2.set_ylabel('复权股价',color="red", fontproperties=font,fontsize = 16)
    ax2.get_yaxis().set_label_coords(1.04,0.2)
    legend = ax2.legend(loc='upper right', fontsize=16)
    ax2.tick_params('y', colors='red')

    ax3.set_ylabel('每股收益',color="green", fontproperties=font,fontsize = 16)
    ax3.get_yaxis().set_label_coords(1.04,0.8)
    ax3.tick_params('y', colors='green')
    legend = ax3.legend(loc='upper center', fontsize=16)
    imgfn = imgdir+'\\'+gpdmdict()[stock].replace(' ','').replace('*','')+'.png'
    plt.savefig(imgfn)

    plt.show()

###############################################################################
#
###############################################################################
def getttmeps(ggdm,hgall,zzgball,epsall):
    gghg = [[date,hg] for gpdm,gpmc,date,hg in hgall if gpdm == ggdm]
    ggzzgb = [[date,zzgb] for gpdm,gpmc,date,zzgb in zzgball if gpdm == ggdm]
    ggeps = [[date,eps] for gpdm,gpmc,date,eps in epsall if gpdm == ggdm]

    df1 = pd.DataFrame(gghg,columns=['date','hg'])
    df2 = pd.DataFrame(ggzzgb,columns=['date','zzgb'])
    df3 = pd.DataFrame(ggeps,columns=['date','eps'])
    df4 = pd.merge(df1,df2,on='date')
    df5 = pd.merge(df4,df3,on='date')
    df5 = df5.sort_values(by='date',ascending=False)
    df5.eval('zzfh=1+zzgb+hg')
    df5['ratio'] = 1

    r = 1
    for index, row in df5.iterrows():
        df5.ix[index,'ratio'] = row['ratio']*r
        r = row['ratio']/row['zzfh']*r

    df5.eval('adj_eps=eps*ratio')
    df5.eval('ttm_eps=adj_eps')


    df5 = df5.sort_values(by='date')
    df5['ttm_eps'] = pd.rolling_sum(df5['ttm_eps'],4)

    i = 1
    es = []
    for index, row in df5.iterrows():
        if i > 1 :
            t = pd.date_range(dt,row['date'],closed='right',freq ='B')
            s = pd.DataFrame(epsttm,index=t,columns=['eps_ttm'])
            es.append(s)
        dt = row['date']
        epsttm = row['ttm_eps']
        i += 1

    t = pd.date_range(dt,parser.parse(dt)+datetime.timedelta(92),closed='right',freq ='B')
    s = pd.DataFrame(epsttm,index=t,columns=['eps_ttm'])
    es.append(s)
    return pd.concat(es)

if __name__ == '__main__':
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    hgall = read_xls('hg.xlsx','Wind资讯',1)
    zzgball = read_xls('zzgb.xlsx','Wind资讯',1)
    epsall = read_xls('eps.xlsx','Wind资讯',1)
    ggdm = '000651'

    eps_ttm = getttmeps(ggdm,hgall,zzgball,epsall)



#    main(sys.argv)

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)



