# -*- coding: utf-8 -*-
"""
本程序用通达信数据对股价前复权，将数据保存到SQLite3数据库

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

import os
import sys
import re
import struct
import datetime
import numpy as np
import pandas as pd
import winreg
import sqlite3
import math

def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    '''
    GPDM股票代码
    RQ交易日期
    OPEN开盘价
    HIGH最高价
    LOW最低价
    CLOSE收盘价
    AMOUT成交额（亿元）
    VOLUME成交量（万股）
    RATE涨跌幅%
    PRE_CLOSE前收盘价
    ADJ_RATE调整涨跌幅为自然对数涨幅
    ADJ_CLOS前复权调整收盘价
    
    GPDM与RQ构成为唯一索引
    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS GJ
           (GPDM TEXT NOT NULL,
            RQ TEXT NOT NULL,
            OPEN REAL NOT NULL DEFAULT (0.00),
            HIGH REAL NOT NULL DEFAULT (0.00),
            LOW REAL NOT NULL DEFAULT (0.00),
            CLOSE REAL NOT NULL DEFAULT (0.00),
            AMOUT REAL NOT NULL DEFAULT (0.00),
            VOLUME REAL NOT NULL DEFAULT (0.00),
            RATE REAL NOT NULL DEFAULT (0.00),
            PRE_CLOSE REAL NOT NULL DEFAULT (0.00),
            ADJ_RATE REAL NOT NULL DEFAULT (0.00),
            ADJ_CLOSE REAL NOT NULL DEFAULT (0.00));''')
    
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ_GJ ON GJ(GPDM,RQ);''')
    
    '''
    GJFQ股价复权控制表
    
    GPDM股票代码
    FHPGRQ最近分红配股日期
    GJFQRQ最后计算股价复权日期
    
    GPDM为主键
    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS GJFQ
           (GPDM TEXT PRIMARY KEY NOT NULL,
            FHPGRQ TEXT,
            GJFQRQ TEXT);''')
    

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

###############################################################################
#从通达信系统读取股票上市日期
###############################################################################
def getssdate():
    fn=gettdxdir()+"\\T0002\\hq_cache\\base.dbf"
    ssrq = dbf2pandas(fn,['gpdm', 'ssdate']) 
    ssrq['ssdate'] = ssrq['ssdate'].map(str2datetime)

    ssrq=ssrq[ssrq['gpdm'].map(lambda x:x[0]).isin(['0','3','6'])]
    ssrq['gpdm']=ssrq['gpdm'].map(lambda x: x+('.SH' if x[0]=='6' else '.SZ'))
    
    return ssrq

##########################################################################
#将字符串转换为时间戳，不成功返回None
##########################################################################
def str2datetime(s):
    if s is None:
        return None
    if ('-' in s) or ('/' in s):
        if '-' in s:
            dt=s.split('-')
        if '/' in s:
            dt=s.split('/')        
        try:
            dt = datetime.datetime(int(dt[0]),int(dt[1]),int(dt[2]))
        except :
            dt = None

    if len(s)==8:
        try:
            dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
        except :
            dt = None

    return dt


###############################################################################
#将通达信.day读入pands
###############################################################################
def day2pd_ln(dayfn,start=None,end=None):
    
    if end == None:
        end=datetime.datetime.now().strftime('%Y%m%d')
    if start == None:
        start='20100101'

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
        if rq==0 or rq<int(start):
            continue
#        print(days,i,rq)
        rq1 = str2datetime(str(rq))
        rq2 = rq1.strftime("%Y-%m-%d")
        kp = kp/100.00
        zg = zg/100.00
        zd = zd/100.00
        sp = sp/100.00
        cje = cje/100000000.00     #亿元
        cjl = cjl/10000.00         #万股
        zf_ln=math.log(sp/qsp)  if (i>0 and qsp>0) else 0.0
        zf = sp/qsp-1 if (i>0 and qsp>0) else 0.0
        records.append([rq1,rq2,kp,zg,zd,sp,cje,cjl,zf,qsp,zf_ln,sp])
        qsp = sp

    df = pd.DataFrame(records,columns=columns)
    df = df.set_index('rq')
    start = str2datetime(start)
    end = str2datetime(end)

    if start == None or end==None :
        return df
    else :
        return df[start:end]

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
def tdxday2pd(gpdm,start=None,end=None):
    
    gpdm=sgpdm(gpdm)

    sc = 'sh' if gpdm[0]=='6' else 'sz'
#    dayfn =getdisk()+'\\tdx\\'+sc+'lday\\'+sc+gpdm+'.day'
    dayfn =gettdxdir()+'\\vipdoc\\'+sc+'\\lday\\'+sc+gpdm+'.day'

    if os.path.exists(dayfn) :
        return day2pd_ln(dayfn,start,end)
    else :
        return []

###############################################################################
#将分红数据读入pands
###############################################################################
def fhsql2pd(gpdm):
    
    gpdm=lgpdm(gpdm)
    conn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    sql="select rq as gqdjr,fh,szg from fh where gpdm=='" +gpdm+"';"
    df=pd.read_sql_query(sql, con=conn)
    df.columns=df.columns.map(lambda x:x.lower())
    conn.close()

    df['gqdjr']=df['gqdjr'].map(lambda x:x.replace('-',''))    
    df['date']=df['gqdjr'].map(str2datetime)

    return df.set_index('date')

###############################################################################
#将配股数据读入pands
###############################################################################
def pgsql2pd(gpdm):
    
    gpdm=lgpdm(gpdm)
    conn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    sql="select rq as gqdjr,pgj,pgbl from pg where gpdm=='" +gpdm+"';"
    df=pd.read_sql_query(sql, con=conn)
    df.columns=df.columns.map(lambda x:x.lower())
    conn.close()

    df['gqdjr']=df['gqdjr'].map(lambda x:x.replace('-',''))    
    df['date']=df['gqdjr'].map(str2datetime)

    return df.set_index('date')

###############################################################################
#合并分红配股数据，并按股权登记日降序排列
###############################################################################
def getfhpg(gpdm):
    
    gpdm=lgpdm(gpdm)
    fh=fhsql2pd(gpdm)
    pg=pgsql2pd(gpdm)
    fhpg=pd.merge(fh,pg,how='outer',on='gqdjr')
    fhpg=fhpg.sort_values(by='gqdjr', ascending=False)
    fhpg=fhpg.loc[:,['gqdjr','fh','szg','pgj','pgbl']]
    return fhpg.fillna(0)


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

################################################################################
#提取DataFrame时间索引，返回日期
################################################################################
def df_timeindex_to_datelist(df):
    dfti = df.index
    dftia = np.vectorize(lambda s: s.strftime('%Y%m%d'))(dfti.to_pydatetime())
    return dftia.tolist()

###############################################################################
#分红日期为股权登记日前复权收盘价
###############################################################################
def adj_close_ln(df,fhpg):

    if len(df)==0 or len(fhpg)==0:
        return df
    fhpg=fhpg.fillna(0)
    for i in range(len(fhpg)):
        date, fh, szg, pgj, pgbl = fhpg.iloc[i]
#        date=nextdtstr(date,-1) #如果是除权除息日则将除权基准日推前一天变为股权登记日

        fqyes = False     #如果股权登记日不在数据范围内则不能进行复权处理

        if len(df.loc[date:date])==1 :      #股权登记日存在交易
            fqyes = True
        else :
            date = df_next_date(df,date,-1) #股权登记日不存在交易则前找交易日
            if len(df.loc[date:date])==1 :   #股权登记日前有交易则进行复权
                fqyes = True

        if fqyes :
            oldclose = df.loc[date,'adj_close']
            newclose = (oldclose - fh + pgj*pgbl)/(1+szg+pgbl)
            newclose = round(newclose,2)    #四舍五入，保留2位小数
            df.loc[date,'adj_close'] = newclose
            nextdate = df_next_date(df,date,1)
            if nextdate == None :
                break
            df.loc[nextdate,'pre_close'] = newclose
            df.loc[nextdate,'rate'] =  df.loc[nextdate,'close']/df.loc[nextdate,'pre_close'] - 1
            df.loc[nextdate,'adj_rate'] =  math.log(df.loc[nextdate,'close']/df.loc[nextdate,'pre_close'])

    ti = df_timeindex_to_datelist(df)
    ti.reverse()
    for i in range(len(ti)):
        date = ti[i]
#        print(i,date)
        if i== 0 :
            df.loc[date,'adj_close'] = df.loc[date,'adj_close']
        else :
#            print(df.loc[date,'adj_close'])
            df.loc[date,'adj_close'] = next_close /(1+next_rate)

        next_close = df.loc[date,'adj_close']
        next_rate = df.loc[date,'rate']

    return df

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

##########################################################################
#n天后日期串，不成功返回None
##########################################################################
def nextdtstr(s,n,sep=None):
    dt = str2datetime(s)
    if dt :
        dt += datetime.timedelta(n)
        if sep is None:
            return dt.strftime("%Y%m%d")
        if sep=='-':
            return dt.strftime("%Y-%m-%d")
        if sep=='/':
            return dt.strftime("%Y/%m/%d")
    else :
        return None

###############################################################################
# 复权股价存入Sqlite3
###############################################################################
def qgjfq2sql(gpdm,start,end):
    today=datetime.datetime.now().strftime('%Y-%m-%d')
 
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    
    gj=tdxday2pd(gpdm,start,end)
    
    fhpg = getfhpg(gpdm)
    
    if len(fhpg)>0 :
        fqgj=adj_close_ln(gj,fhpg)
    else:
        fqgj=gj
        
    if len(fqgj)>0:
        fqgj['gpdm']=lgpdm(gpdm)
        
        fqgj=fqgj.loc[:,['gpdm','date','open','high','low','close','amout','volume',
                         'rate','pre_close','adj_rate','adj_close']]
        
        fqgj=fqgj.round({'amout':4,'volume':2,'rate': 4,'adj_rate': 4, 'adj_close': 3})
        
        data=np.array(fqgj).tolist()
        
        try:
            dbcn.executemany('''INSERT OR REPLACE INTO GJ (GPDM,RQ,OPEN,HIGH,LOW,CLOSE,
                         AMOUT,VOLUME,RATE,PRE_CLOSE,ADJ_RATE,ADJ_CLOSE) 
                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''', data)
            dbcn.execute('''UPDATE GJFQ SET GJFQRQ=? WHERE GPDM==? ;''',(today,gpdm))
            dbcn.commit()
        except:
            dbcn.close()
            return False
    else:
        dbcn.execute('''UPDATE GJFQ SET GJFQRQ=? WHERE GPDM==? ;''',(today,gpdm))
        dbcn.commit()
    
    dbcn.close()
    return True


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
        
    #去重
    zxg = list(set(zxglst))
    #排序
    zxg.sort(key=zxglst.index)
    #变长代码
    zxglst=[]
    for x in zxg:
        zxglst.append(lgpdm(x))

    return zxglst

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

def main1(): 
    start=(datetime.date.today()-datetime.timedelta(days=365)).strftime('%Y%m%d')
    end=datetime.date.today().strftime('%Y%m%d')
    
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    gpdmb=get_gpdm()

    createDataBase()
#    dm='002673'
#    j=list(gpdmb['gpdm']).index(dm)
#    k=len(gpdmb)       #最大值为自选股总数len(gpdmb)        
    i=300
    k=301
    while i<k:
        gpdm = gpdmb.loc[i-1,'gpdm']
        gpmc = gpdmb.loc[i-1,'gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (k,i,gpdm,gpmc)) 
        if qgjfq2sql(gpdm,start,end):
            i += 1
            

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)


#########################################################################
#自选股复权
#########################################################################
def main2():
    start=(datetime.date.today()-datetime.timedelta(days=365)).strftime('%Y%m%d')
    end=datetime.date.today().strftime('%Y%m%d')
    
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    createDataBase()
    zxgfile="zxg.blk"
    tdxblkdir = gettdxblkdir()
    zxgfile = os.path.join(tdxblkdir,zxgfile)
    lst1 = zxglist(zxgfile,"tdxblk")

    gpdmb=get_gpdm()
    gpdmb=gpdmb.set_index('gpdm')

    lst2=yfqgplst(0)
    zxglb=list_diff(lst1,lst2)
    
    i=0
    k=len(zxglb)
    while i<k:
        gpdm=lgpdm(zxglb[i])
        gpmc=gpdmb.loc[gpdm,'gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (k,i+1,gpdm,gpmc)) 
        if qgjfq2sql(gpdm,start,end):
            i += 1

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)


########################################################################
#库内已复权股票信息
########################################################################
def gjfqkzxx():
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
#    curs=dbcn.cursor()
#    
#    curs.execute('select distinct gpdm,rq from gj;')
#    data = curs.fetchall()
 
    sql="select distinct gpdm,rq from gj order by rq desc;"
    df=pd.read_sql_query(sql, con=dbcn)
    df.columns=df.columns.map(lambda x:x.lower())
    df1=df.drop_duplicates(['gpdm'],keep='first')

    sql='''select distinct gpdm,rq from (
        select gpdm,rq from fh union
        select gpdm,rq from pg) order by gpdm,rq desc;'''
    
    df=pd.read_sql_query(sql, con=dbcn)
    df.columns=df.columns.map(lambda x:x.lower())
#    df=df.sort_values(by=['gpdm','rq'], ascending=False)
    df2=df.drop_duplicates(['gpdm'],keep='first')
    
    df3=pd.merge(df1,df2,how='outer',on='gpdm')
    df3.columns=['gpdm','fqgjrq','fhpgrq']
    
    return df3


########################################################################
#列表差集
########################################################################
def list_diff(a,b):
    ret_list = []
    for item in a:
        if item not in b:
            ret_list.append(item)
    return ret_list
     
########################################################################
#更新库内GJFQ表股票代码
########################################################################
def update_gpdm():  

    gpdm1=np.array(get_gpdm()['gpdm']).tolist()     #最新代码表

    createDataBase()
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    cursor = dbcn.cursor()
    sql="select gpdm from gjfq;"

    #数据库查询
    cursor.execute(sql)
    gpdm2=cursor.fetchall()         #在库内的代码

    #变为一维列表
    gpdm3=[]
    for e in gpdm2:
        gpdm3.append(e[0])
 
    gpdm4=list_diff(gpdm1,gpdm3)   #去掉已入库的
    
    #变为二维列表
    gpdm5=[]
    for e in gpdm4:
        gpdm5.append([e])
        
    if len(gpdm5)>0:
        dbcn.executemany('''INSERT OR IGNORE INTO GJFQ (GPDM) VALUES (?)''', gpdm5)

    dbcn.commit()
    dbcn.close()
   

########################################################################
#提取GJFQ日期到pandas，股票代码位索引
########################################################################
def get_fqrq():
    update_gpdm()
     
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')

    sql="select gpdm,gjfqrq,fhpgrq from gjfq;"
    df=pd.read_sql_query(sql, con=dbcn)
    
    df.columns=df.columns.map(lambda x:x.lower())
    
    return df.set_index('gpdm')

########################################################################
#5天内进行复权处理过的股票列表
########################################################################
def yfqgplst(n=-10):
    fqrqb=get_fqrq() 
    today=datetime.datetime.now().strftime('%Y-%m-%d')
    lastday=nextdtstr(today,n,'-')
    #已复权股票代码列表
    return fqrqb.loc[fqrqb['gjfqrq']>=lastday,:].index.tolist()
    
def main3():
    start=(datetime.date.today()-datetime.timedelta(days=365)).strftime('%Y%m%d')
    end=datetime.date.today().strftime('%Y%m%d')
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    gpdmb=get_gpdm()
    gpdmb=gpdmb.set_index('gpdm',drop=False)

    #全部股票
    lst1=np.array(gpdmb['gpdm']).tolist()     #最新代码表

    #已复权股票列表
    lst2=yfqgplst(-5)

    zxglb=list_diff(lst1,lst2)

    i=0
    k=len(zxglb)
    while i<k:
        gpdm=lgpdm(zxglb[i])
        gpmc=gpdmb.loc[gpdm,'gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (k,i+1,gpdm,gpmc)) 
        if qgjfq2sql(gpdm,start,end):
            i += 1
    
    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)


########################################################################
#日期计算
########################################################################
def calcdate(datestr,years=None,months=None,days=None):

    dt=str2datetime(datestr)

    if dt==None :
        return None
    if years==None:
        years=0
    if months==None:
        months=0
    if days==None:
        days=0

    y=dt.year+years+((dt.month+months-1)//12)
    
    m=((dt.month+months-1)%12)+1
    
    dt=datetime.datetime(y,m,dt.day)+datetime.timedelta(days=days)
    if '/' in datestr:
        return dt.strftime('%Y/%m/%d')
    if '-' in datestr:
        return dt.strftime('%Y-%m-%d')        
    return dt.strftime('%Y%m%d')
    
########################################################################
#个股涨幅
########################################################################
def ggzf(gpdm,start,end):
     
    gj=tdxday2pd(gpdm,start,end)
    
    fhpg = getfhpg(gpdm)
    
    if len(fhpg)>0 :
        fqgj=adj_close(gj,fhpg)
    else:
        fqgj=gj
        
    if len(fqgj)>0 :
        return fqgj.iloc[len(fqgj)-1]['adj_close']/fqgj.iloc[0]['adj_close']-1
    else:
        return None

########################################################################
#股票基本信息表
########################################################################
def gpjbxxb():
    
    gpdmb=get_gpdm()

    gpssrq=getssdate()
    ssrq = pd.merge(gpdmb,gpssrq,on="gpdm")

    return ssrq

def zftj():
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    gpdmb=gpjbxxb()
    gpdmb=gpdmb.set_index('gpdm',drop=False)
    dmb=gpdmb.loc[gpdmb['ssdate']<'2016-12-31']['gpdm']
    zfb=[]
#    n=-5
#    end=datetime.date.today().strftime('%Y%m%d')
#    start=calcdate(end,years=n)

    end="20171231"
    start="20170101"
    i=1
    for gpdm in dmb:
        gpmc=gpdmb.loc[gpdm,'gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(dmb),i,gpdm,gpmc)) 
        zf=ggzf(gpdm,start,end)
        zfb.append([gpdm,zf])

        if int(i/100)==i/100 :
            now2 = datetime.datetime.now().strftime('%H:%M:%S')
            print('开始运行时间：%s' % now1)
            print('结束运行时间：%s' % now2)

        i+=1
            
        
    df = pd.DataFrame(zfb,columns=['gpdm','zf'])
    df = pd.merge(gpdmb,df,on="gpdm")
    df.columns=['股票代码','股票名称','股票拼音','上市日期','涨幅']
    df.to_excel(r'd:\hyb\zftj2017.xlsx',index=False)
    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

if __name__ == "__main__":  
    zftj()

#    gpdm='600145'
#    start='20170101'
#    end='20171231'
#    gj=tdxday2pd(gpdm,start,end)
#    
#    fhpg = getfhpg(gpdm)
#    
#    if len(fhpg)>0 :
#        fqgj=adj_close(gj,fhpg)
#    else:
#        fqgj=gj

#    fqrqb=get_fqrq() 
#    
#    today=datetime.datetime.now().strftime('%Y-%m-%d')
#    lastday=nextdtstr(today,-5,'-')
#    
#    yfq=fqrqb.loc[fqrqb['gjfqrq']>=lastday,:].index.tolist()
#
#    
#    now1 = datetime.datetime.now().strftime('%H:%M:%S')
#    print('开始运行时间：%s' % now1)
#
#    zxgfile="zxg.blk"
#    tdxblkdir = gettdxblkdir()
#    zxgfile = os.path.join(tdxblkdir,zxgfile)
#    zxglb = zxglist(zxgfile,"tdxblk")
#    zxg=[]
#    for x in zxglb:
#        zxg.append(lgpdm(x))
#
#    gpdmb=get_gpdm()
#    gpdmb=gpdmb.set_index('gpdm')
    
#    i=0
#    k=len(zxglb)
#    while i<k:
#        gpdm=lgpdm(zxglb[i])
#        gpmc=gpdmb.loc[gpdm,'gpmc']
#        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (k,i+1,gpdm,gpmc)) 
#        if qgjfq2sql(gpdm):
#            i += 1
#
#    now2 = datetime.datetime.now().strftime('%H:%M:%S')
#    print('开始运行时间：%s' % now1)
#    print('结束运行时间：%s' % now2)
#     


    
    
    
    
