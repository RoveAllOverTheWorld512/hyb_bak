# -*- coding: utf-8 -*-
"""
Created on Tue Nov 28 08:13:07 2017

@author: lenovo
"""

import sqlite3
import pandas as pd
import datetime
import sys
import struct
import winreg
#import numpy as np

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
#从通达信系统读取股票上市日期
###############################################################################
def getssdate():
    fn=gettdxdir()+"\\T0002\\hq_cache\\base.dbf"
    ssrq = dbf2pandas(fn,['gpdm', 'ssdate']) 
    ssrq['ssdate'] = ssrq['ssdate'].map(str2datetime)

    ssrq=ssrq[ssrq['gpdm'].map(lambda x:x[0]).isin(['0','3','6'])]
    ssrq['gpdm']=ssrq['gpdm'].map(lambda x: x+('.SH' if x[0]=='6' else '.SZ'))
    
    return ssrq

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

def getdrive():
    return sys.argv[0][:2]

if __name__ == '__main__':
    
    gpdmb=get_gpdm()
    ssrq=getssdate()

    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)
    curs = dbcn.cursor()
    
    curs.execute('''select gpdm,rq,gdhs from gdhs 
              where rq>='2017-09-30' order by rq desc;''')
    
    data = curs.fetchall()
    
    df=pd.DataFrame(data,columns=['gpdm','rq','gdhs'])
    #保留最新户数
    df1=df.drop_duplicates(['gpdm'],keep='first')
    
    curs.execute('''select gpdm,rq,gdhs from gdhs 
              where rq=='2017-09-30';''')
    data = curs.fetchall()
    df2=pd.DataFrame(data,columns=['gpdm','rq','gdhs'])
    
    dbcn.commit()
    dbcn.close()
    
    df3=pd.merge(df1, df2, how='left', on='gpdm')
    
    df4=pd.merge(df3, gpdmb, how='left', on='gpdm')
    
    df4=pd.merge(df4, ssrq, how='left', on='gpdm')

    df4=df4.loc[:,['gpdm','gpmc','gppy','ssdate','gdhs_x','gdhs_y']]
    
    #提取股东户数变化的股票
#    df4=df4.loc[df4['gdhs_x']!=df4['gdhs_y']]
    df4['hbzzl']=(df4['gdhs_x']/df4['gdhs_y']-1)*100
    
    df4.columns=['股票代码','股票简称','股票拼音','上市日期','股东户数(户)2017.12.31','股东户数(户)2017.09.30','股东户数环比增长率']
    
    today = datetime.datetime.now().strftime("%Y%m%d")    
    
    fn=getdrive()+'\\hyb\\股东户数_'+today+'.xlsx'

    writer = pd.ExcelWriter(fn, engine='xlsxwriter')
    
    df4.to_excel(writer, sheet_name='最新股东户数',index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['最新股东户数']
    
    #    format1 = workbook.add_format({'num_format': '0.0000'})
    #    format2 = workbook.add_format({'num_format': '0.00'})
    #format3 = workbook.add_format({'num_format': 'yyyy-mm-dd'})
    #
    #worksheet.set_column('A:P', 10)
    #worksheet.set_column('I:J', 12, format3)
    #worksheet.set_column('O:P', 12, format3)
    #worksheet.freeze_panes(1, 0)
    
    writer.save()
    print('请打开%s文件查看最新股东户数股票名单。' % fn)
    
