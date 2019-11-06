# -*- coding: utf-8 -*-
"""
Created on Thu Nov 16 11:43:30 2017

@author: lenovo
"""
import os
import sys
import struct
import datetime
import pandas as pd
import sqlite3

########################################################################
#建立数据库
########################################################################
def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')

    '''
    GPDMB股票代码表
    
    GPDM股票代码
    GPMC股票名称
    PY拼音
    SSRQ上市日期
    
    GPDM为主键
    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS SSDATE
           (GPDM TEXT PRIMARY KEY NOT NULL,
           GPMC TEXT,
           PY TEXT,
           SSRQ TEXT);''')



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
#将字符串转换为时间戳，不成功返回None
##########################################################################
def datefmt(s):
    try:
        dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
        dt = s[:4]+'-'+s[4:6]+'-'+s[6:8]
    except :
        dt = None
    return dt

###############################################################################
#从通达信系统读取股票代码表
###############################################################################
def getcode():
    datacode = []
    for sc in ('h','z'):
        fn = r'C:\new_hxzq_hc\T0002\hq_cache\s'+sc+'m.tnf'
        f = open(fn,'rb')
        f.seek(50)
        ss = f.read(314)
        while len(ss)>0:
            gpdm=ss[0:6].decode('GBK')
            gpmc=ss[23:31].strip(b'\x00').decode('GBK').replace(' ','')
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
#从通达信系统读取股票上市时间
###############################################################################
def getssdate():
    fn=r"C:\new_hxzq_hc\T0002\hq_cache\base.dbf"
    sssj = dbf2pandas(fn,['gpdm', 'ssdate']) 
    sssj['ssdate'] = sssj['ssdate'].map(datefmt)

    sssj=sssj[sssj['gpdm'].map(lambda x:x[0]).isin(['0','3','6'])]
    sssj['gpdm']=sssj['gpdm'].map(lambda x: x+('.SH' if x[0]=='6' else '.SZ'))
    
    
    return sssj

if __name__ == '__main__':
    gpdmb=getcode()
    
    gpsssj=getssdate()
    gpdmsssj = pd.merge(gpdmb,gpsssj,on="gpdm")
    gpdmsssj.columns=['股票代码','股票简称','股票拼音','上市日期']

    gpdmsssj.to_excel(r'd:\hyb\股票代码表.xlsx',sheet_name='股票代码表',index=False)
    
    gpdmsssj.columns=['GPDM','GPMC','PY','SSRQ']
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    gpdmsssj.to_sql('GPDMB', dbcn,index=False,if_exists='replace')
    dbcn.commit()
    dbcn.close()



#    gpdmsssj.to_excel(r'd:\hyb\股票代码上市时间.xlsx',index=False)
