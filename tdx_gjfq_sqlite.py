# -*- coding: utf-8 -*-
"""
本程序用通达信数据对股价前复权，将数据保存为excel文件

通达信本地数据格式：
每32个字节为一个5分钟数据，每字段内低字节在前
00 ~ 01 字节：日期，整型，设其值为num，则日期计算方法为：
                        year=floor(num/2048)+2004;
                        month=floor(mod(num,2048)/100);
                        day=mod(mod(num,2048),100);
02 ~ 03 字节： 从0点开始至目前的分钟数，整型
04 ~ 07 字节：开盘价*100，整型
08 ~ 11 字节：最高价*100，整型
12 ~ 15 字节：最低价*100，整型
16 ~ 19 字节：收盘价*100，整型
20 ~ 23 字节：成交额*100，float型
24 ~ 27 字节：成交量（股），整型
28 ~ 31 字节：（保留）

每32个字节为一天数据
每4个字节为一个字段，每个字段内低字节在前
00 ~ 03 字节：年月日, 整型
04 ~ 07 字节：开盘价*100， 整型
08 ~ 11 字节：最高价*100,  整型
12 ~ 15 字节：最低价*100,  整型
16 ~ 19 字节：收盘价*100,  整型
20 ~ 23 字节：成交额（元），float型
24 ~ 27 字节：成交量（股），整型
28 ~ 31 字节：（保留）

读取需要加载struct模块，unpack之后得到一个元组。
日线读取：
fn=r"code.day";
fid=open(fn,"rb");
list=fid.read(32)
ulist=struct.unpack("iiiiifii", list)
5分钟线读取也是一样。

"""

import re
import os
import sys
import struct
import datetime
import numpy as np
import pandas as pd
import winreg
import sqlite3

##########################################################################
#读取当前工作路径盘符
##########################################################################
def getdisk():
    return sys.argv[0][:2]

##########################################################################
#将字符串转换为时间戳，不成功返回None
##########################################################################
def str2datetime(s):
    try:
        dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
    except :
        dt = None
    return dt

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
#将通达信.day读入pands
###############################################################################
def day2pandas(dayfn,start=None,end=None):
    columns = ['rq','date','open', 'high', 'low','close','amout','volume','rate','pre_close','adj_rate','adj_close']

    with open(dayfn,"rb") as f:
        data = f.read()
        f.close()
    days = int(len(data)/32)
    records = []
    qsp = 0
    for i in range(days):
        dat = data[i*32:(i+1)*32]
        rq,kp,zg,zd,sp,cje,cjl,tmp = struct.unpack("iiiiifii", dat)
        rq1 = str2datetime(str(rq))
        rq2 = rq1.strftime("%Y-%m-%d")
        kp = kp/100.00
        zg = zg/100.00
        zd = zd/100.00
        sp = sp/100.00
        cje = cje/100000000.00     #亿元
        cjl = cjl/10000.00         #万股
        zf = sp/qsp-1 if (i>0 and qsp>0) else 0.0
        records.append([rq1,rq2,kp,zg,zd,sp,cje,cjl,zf,qsp,zf,sp])
        qsp = sp

    df = pd.DataFrame(records,columns=columns)
    df = df.set_index('rq')
    start = str2datetime(start)
    end = str2datetime(end)


    if start == None or end==None :
        return df
    else :
        return df[start:end]


###############################################################################
#将通达信.day读入pands
###############################################################################
def topandas(gpdm,start=None,end=None):

    sc = 'sh' if gpdm[0]=='6' else 'sz'
#    dayfn =getdisk()+'\\tdx\\'+sc+'lday\\'+sc+gpdm+'.day'
    dayfn =gettdxdir()+'\\vipdoc\\'+sc+'\\lday\\'+sc+gpdm+'.day'

    if os.path.exists(dayfn) :
        return day2pandas(dayfn,start,end)
    else :
        return []

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


###############################################################################
#将dbf读入pands
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

###############################################################################
#将分红配股数据读入pands
###############################################################################
def fhpg2pandas(fhpgfn):
#    fhpgfn = r'D:\公司研究\生物股份\600201生物股份分红配股.dbf'
    fhpg = dbf2pandas(fhpgfn,['gqdjr', 'mgfh', 'mgsg'])
    fhpg['date'] = fhpg['gqdjr'].map(str2datetime)
    
    return fhpg.set_index('date')


###############################################################################
#将分红配股数据读入pands
###############################################################################
def fhsql2pd(gpdm):
    conn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    gpdm=gpdm+('.SH' if gpdm[0]=='6' else '.SZ')
    sql="select rq as gqdjr,fh as mgfh,szg as mgsg from lnfhkg where gpdm=='" +gpdm+"';"
    df=pd.read_sql_query(sql, con=conn)
    conn.close()
    df['gqdjr']=df['gqdjr'].map(lambda x:x.replace('-',''))    
    df['date']=df['gqdjr'].map(str2datetime)
    return df.set_index('date')

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

################################################################################
#提取DataFrame时间索引，返回日期
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
#分红日期为股权登记日前复权收盘价
###############################################################################
def adj_close(df,fhpg):
    for i in range(len(fhpg)):
        date, mgfh, mgsg = fhpg.iloc[i]
#        date=nextdtstr(date,-1) #将除权基准日推前一天变为股权登记日

        fqyes = False     #如果股权登记日不在数据范围内则不能进行复权处理

        if len(df.loc[date:date])==1 :      #股权登记日存在交易
            fqyes = True
        else :
            date = df_next_date(df,date,-1) #股权登记日不存在交易则前找交易日
            if len(df.loc[date:date])==1 :   #股权登记日前有交易则进行复权
                fqyes = True

        if fqyes :
            oldclose = df.loc[date,'adj_close']
            newclose = (oldclose - mgfh)/(1+mgsg)
            df.loc[date,'adj_close'] = newclose
            nextdate = df_next_date(df,date,1)
            if nextdate == None :
                break
            df.loc[nextdate,'pre_close'] = newclose
            df.loc[nextdate,'adj_rate'] =  df.loc[nextdate,'close']/df.loc[nextdate,'pre_close']- 1

    ti = df_timeindex_to_datelist(df)
    ti.reverse()
    for i in range(len(ti)):
        date = ti[i]
        if i== 0 :
            df.loc[date,'adj_close'] = df.loc[date,'adj_close']
        else :
            df.loc[date,'adj_close'] = next_close /(1+next_rate)

        next_close = df.loc[date,'adj_close']
        next_rate = df.loc[date,'adj_rate']

    return df


###############################################################################
# 复权股价
###############################################################################
def gjfq(gpdm):
    
    dfy=topandas(gpdm)
    
    fh = fhsql2pd(gpdm)
    
    if len(fh)>0 :
        adj_close(dfy,fh)
    
    return dfy

def saveexcel(gjfn,df):
    
    writer = pd.ExcelWriter(gjfn, engine='xlsxwriter')

    df.to_excel(writer, sheet_name='股价与成交量',index=False)

    workbook = writer.book
    worksheet = writer.sheets['股价与成交量']

    format1 = workbook.add_format({'num_format': '0.0000'})
    format2 = workbook.add_format({'num_format': '0.00'})
    format3 = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    
    worksheet.set_column('A:A', 11, format3)
    worksheet.set_column('B:E', 12, format2)
    worksheet.set_column('F:F', 14, format1)
    worksheet.set_column('G:G', 14, format2)
    worksheet.set_column('H:H', 10, format1)
    worksheet.set_column('I:I', 10, format2)
    worksheet.set_column('J:K', 19, format1)
    worksheet.freeze_panes(1, 0)

    writer.save()

######################################################################################
#检测路径是否存在，不存则创建
######################################################################################    
def exsit_path(pth):
    if not os.path.exists(pth) :
        os.makedirs(pth)

########################################################################
#股票代码表
########################################################################
def gpdmdict():
    fn = getdisk()+'\\hyb\\gpdmb.txt'
    with open(fn) as f:
        gpdmb = f.read()
        f.close()

    dmb = re.findall('(\d{6})\t(.+)\n',gpdmb)
    dm = {}
    for (gpdm,gpmc) in dmb :
        dm[gpdm] = gpmc.replace(" ","").replace("*","")

    return dm


########################################################################
#股票代码表
########################################################################
def get_pe(gpdm):
    conn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    gpdm=gpdm+('.SH' if gpdm[0]=='6' else '.SZ')
    sql="select rq,pe_lyr,pe_ttm from pe_pb where gpdm=='" +gpdm+"';"
    df=pd.read_sql_query(sql, con=conn)
    
    df.rename(columns=lambda x:x.lower(), inplace=True)
    
    return df

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

if __name__ == '__main__':        
    
    zxgfile="zxg.blk"
    tdxblkdir = gettdxblkdir()
    zxgfile = os.path.join(tdxblkdir,zxgfile)
    zxglb = zxglist(zxgfile,"tdxblk")
    j=155       #最小值为1
    k=155
    l=k if k<=len(zxglb) else len(zxglb)
    for i in range(j-1,l):
        gpdm=zxglb[i]
        gpmc = gpdmdict()[gpdm]
        gpmc = gpmc.replace(" ","").replace("*","")

        print('共有%d只股票,正在处理第%d只股票%s%s，请等待。' %(len(zxglb),i+1,gpdm,gpmc))

        pth =  'D:/公司研究/'+gpmc
        exsit_path(pth)
    
        gjfn = pth+'/'+gpdm+gpmc+'股价与成交量.xlsx'
    
        fqgj=gjfq(gpdm)
        pe=get_pe(gpdm)
        pe.columns=['date','pe_lyr','pe_ttm']
        gjpe=pd.merge(fqgj,pe,on='date')
        
        columns = ['日期','开盘价(元)', '最高级(元)', '最低价(元)','收盘价(元)','成交额(亿元)','成交量(万股)','涨幅','前收盘','调整涨幅','前复权收盘价(元)','静态市盈率','滚动市盈率']
        gjpe.columns=columns
        saveexcel(gjfn,gjpe)

        