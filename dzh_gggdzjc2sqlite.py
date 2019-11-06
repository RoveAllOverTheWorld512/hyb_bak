# -*- coding: utf-8 -*-
"""
从大智慧F10提取高管股东增减持数据导入Sqlite数据库
"""
from pyquery import PyQuery as pq
import datetime
import sqlite3
import sys
import os
import re
import pandas as pd
import numpy as np
import winreg
import struct

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
#从大智慧网F10获取股东户数
########################################################################
def get_gdhs(gpdm):

    sc=scdm(gpdm)
    gpdm=sgpdm(gpdm)
    
    data=[]
    url = 'http://webf10.gw.com.cn/'+sc+'/B10/'+sc+gpdm+'_B10.html'

    try :
        html = pq(url,encoding="utf-8")
        #第3个区块
        #sect = pq(html('section').eq(2).html())
        #提取预测明细
        sect=html('section').filter('#股东人数').html()
        tr=pq(sect)
    except : 
        print("出错退出")
        return data

    for i in range(1,len(tr('ul'))):
        
        il=tr('ul').eq(i).text().split(' ')
        rq=il[0]
        gdhs=il[1]

        data.append([lgpdm(gpdm),rq,gdhs])

    return data
    
########################################################################
#从大智慧网F10获取限售解禁
########################################################################
def get_xsjj(gpdm):

    sc=scdm(gpdm)
    gpdm=sgpdm(gpdm)
    
    data=[]
    url = 'http://webf10.gw.com.cn/'+sc+'/B11/'+sc+gpdm+'_B11.html'

    try :
        html = pq(url,encoding="utf-8")
        #第3个区块
        #sect = pq(html('section').eq(2).html())
        #提取预测明细
        sect=html('section').filter('#解禁流通').html()
                 
        tbl=pq(sect)
        tr=pq(tbl('table').eq(1).html())
        
    except : 
        print("出错退出")
        return data

    for i in range(1,len(tr('tr'))):
        
        td=tr('tr').eq(i).text().split(' ')
        jjrq=td[0].replace('/','-')
        bcjj=round(float(td[1])/10000,4)
        wlt=float(td[6])/10000
        try:
            qltbl=float(td[3].replace('%',''))
            qlt=round(bcjj/qltbl*100,4)
            hlt=round(qlt+bcjj,4)
            hltbl=round(bcjj/hlt*100,4)
        except:            
            qltbl=None
            qlt=None
            hlt=None
            hltbl=None
            
        data.append([lgpdm(gpdm),jjrq,bcjj,qlt,qltbl,hlt,hltbl,None,None,wlt])

    return data
 
########################################################################
#从大智慧网F10获取高管股东持股变动
########################################################################
def get_gggdcgbd(gpdm):
    sc=scdm(gpdm)
    gpdm=sgpdm(gpdm)
    
    data=[]
    url = 'http://webf10.gw.com.cn/'+sc+'/B10/'+sc+gpdm+'_B10.html'

    try :
        html = pq(url,encoding="utf-8")

        sect=html('section').filter('#高管持股变动').html()
                 
        tbl=pq(sect)

        if len(tbl('tr'))>1:
            for i in range(1,len(tbl('tr'))):
                
                row=pq(tbl('tr').eq(i))
                bdrq=row('td').eq(0).text().replace('/','-')
                bdr=row('td').eq(1).text().split(' ')[0]
                bdfx=row('td').eq(3).text()[:2]
                try:
                    bdsl=float(row('td').eq(3).text()[2:])
                    data.append([lgpdm(gpdm),bdrq,bdr,bdsl,bdfx])
                except:
                    pass
    except : 
        print("高管持股变动,出错退出")


    try :
        html = pq(url,encoding="utf-8")

        sect=html('section').filter('#股东持股变动').html()
                 
        tbl=pq(sect)
        
        if len(tbl('tr'))>1:
            for i in range(1,len(tbl('tr'))):
                
                row=pq(tbl('tr').eq(i))
                
                bdrq=row('td').eq(0).text().replace('/','-')
                bdr=row('td').eq(1).text()
                bdfx=row('td').eq(3).text()
                try:
                    bdsl=float(row('td').eq(2).text())*10000
                    bdsl=-bdsl if bdfx=='减持' else bdsl
                    data.append([lgpdm(gpdm),bdrq,bdr,bdsl,bdfx])
                except:
                    pass
    except : 
        print("股东持股变动,出错退出")

    
    return data
    
'''
CREATE TABLE [XSJJ](
  [GPDM] TEXT NOT NULL, 
  [JJRQ] TEXT NOT NULL, 
  [JJSL] REAL NOT NULL, 
  [QLTGB] REAL, 
  [QLTBL] REAL, 
  [HLTGB] REAL, 
  [HLTBL] REAL, 
  [QZD] REAL, 
  [HZD] REAL, 
  [WLT] REAL);

CREATE UNIQUE INDEX [GPDM_JJRQ_XSJJ]
ON [XSJJ](
  [GPDM], 
  [JJRQ]);

        
CREATE TABLE [GGGDCGBD](
  [GPDM] TEXT NOT NULL, 
  [BDRQ] TEXT NOT NULL, 
  [BDR] TEXT NOT NULL, 
  [BDSL] REAL NOT NULL, 
  [BDFX] TEXT NOT NULL);

CREATE UNIQUE INDEX [GPDM_BDRQ_BDR_GGGDCGBD]
ON [GGGDCGBD](
  [GPDM], 
  [BDRQ], 
  [BDR]);

'''    

###############################################################################
#高管股东增减持统计
###############################################################################
def gggdzjctj():
    td=datetime.datetime.now()

    m0=td.strftime("%Y-%m-%d")
    m1=(td+datetime.timedelta(-30*6)).strftime("%Y-%m-%d")

    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)
    curs = dbcn.cursor()
    sql='''select distinct gpdm from gggdcgbd 
               where bdrq>="%s" and bdrq<"%s" 
            ;'''  % (m1,m0)
    curs.execute(sql)        
    data = curs.fetchall()
    dbcn.close()
    
    cols=['gpdm']
    df = pd.DataFrame(data,columns=cols)
    df = df.set_index('gpdm',drop=False)
    zjchz=df['gpdm'].map(gggdzjc_sum)

    zjcdf=zjchz.to_frame()

    #高管股东增减持金额汇总，最近六个月、三个月、一个月
    zjcdf['m6zjc']=[e[0] for e in zjcdf['gpdm']]
    zjcdf['m3zjc']=[e[1] for e in zjcdf['gpdm']]
    zjcdf['m1zjc']=[e[2] for e in zjcdf['gpdm']]

    zjcdf=zjcdf.drop(columns=['gpdm'])
    zjcdf=zjcdf.round(4)

    #转换NaN
    zjcdf=zjcdf.replace(0,np.nan) 

    zjcdf.to_csv('zjcdf.csv')
    
    return        
    
###############################################################################
#高管股东增减持统计汇总
###############################################################################
def gggdzjc_sum(gpdm):
    print(gpdm)
    td=datetime.datetime.now()

    m0=td.strftime("%Y-%m-%d")
    m1=(td+datetime.timedelta(-30*6)).strftime("%Y-%m-%d")
    m2=(td+datetime.timedelta(-30*3)).strftime("%Y-%m-%d")
    m3=(td+datetime.timedelta(-30*1)).strftime("%Y-%m-%d")

    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)
    curs = dbcn.cursor()
    sql='''select bdrq,bdsl,bdfx from gggdcgbd 
               where gpdm="%s" and bdrq>="%s" and bdrq<"%s" 
            ;'''  % (lgpdm(gpdm),m1,m0)
    curs.execute(sql)        
    data = curs.fetchall()
    cols=['date','bdsl','bdfx']
    df = pd.DataFrame(data,columns=cols)
    df=df.set_index('date')    
    
    gjdf=get_fqgj(gpdm,'close')
    df=df.join(gjdf)    
    #计算变动金额，单位“亿元”
    df=df.assign(bdje=df['bdsl']*df['close']/100000000)
    m6sum=df.loc[m1:m0,'bdje'].sum()
    m3sum=df.loc[m2:m0,'bdje'].sum()
    m1sum=df.loc[m3:m0,'bdje'].sum()

    return [m6sum,m3sum,m1sum]

###############################################################################
#前复权复权，复权因子是从tushare下载的csv文件
###############################################################################
def get_fqgj(gpdm,cols=None):

    data=tdxday2pd(gpdm)
    
    #未上市新股
    if isinstance(data,list):
        return []
    
    cols1=['date','adj_close']
    
    if len(gpdm)==6 and (gpdm[0] =='6' or gpdm[:2]=='00' or gpdm[:3] in ('002','300')):    
        #计算前复权因子
        adjf=factor_adj(gpdm)
        
        #前复权收盘价
        data=data.join(adjf)
        #补缺，可以解决tushare还没有更新，而当天收盘结束后缺失复权因子导致无法进行后续涨幅计算
        data['adj_factor'].fillna(method='ffill',inplace=True)
        data=data.eval('adj_close = adj_close * adj_factor')
    
    if cols!=None:    
        if isinstance(cols,list) or isinstance(cols,tuple):
            for col in cols:
                if col in data.columns and col not in cols1:
                    cols1.append(col)
        elif isinstance(cols,str):
            if cols in data.columns and cols not in cols1:
                cols1.append(cols)
            
    fqgj=data[cols1]
        
    '''
    pandas 将“字符类型的日期列”转化成“时间戳索引（DatetimeIndex）”
    https://www.jianshu.com/p/4ece5843d383
    '''
    fqgj=fqgj.set_index('date')    

    fqgj.index = pd.DatetimeIndex(fqgj.index)
    
    return fqgj.round(3)


###############################################################################
#复权因子是从tushare下载的csv文件
###############################################################################
def factor_adj(gpdm):
    df = pd.read_csv('http://file.tushare.org/tsdata/f/factor/%s.csv' % gpdm)
    df = df.set_index('datetime')
    df.index.name = 'rq'
    df=df.sort_index(ascending = False)
    a = df.iloc[0][0]
    df=df/a

    return df

###############################################################################
#前复权复权，复权因子是从tushare下载的csv文件
###############################################################################
def get_fqgj1(gpdm,cols=None):

    data=tdxday2pd(gpdm)
    
    #未上市新股
    if isinstance(data,list):
        return []
    
    cols1=['date','adj_close']
    
    if len(gpdm)==6 and (gpdm[0] =='6' or gpdm[:2]=='00' or gpdm[:3] in ('002','300')):    
        #计算前复权因子
        adjf=factor_adj(gpdm)
        
        #前复权收盘价
        data=data.join(adjf)
        data=data.eval('adj_close = adj_close * adj_factor')
    
    if cols!=None:    
        if isinstance(cols,list) or isinstance(cols,tuple):
            for col in cols:
                if col in data.columns and col not in cols1:
                    cols1.append(col)
        elif isinstance(cols,str):
            if cols in data.columns and cols not in cols1:
                cols1.append(cols)
            
    fqgj=data[cols1]
        
    '''
    pandas 将“字符类型的日期列”转化成“时间戳索引（DatetimeIndex）”
    https://www.jianshu.com/p/4ece5843d383
    '''
    fqgj=fqgj.set_index('date')    

    fqgj.index = pd.DatetimeIndex(fqgj.index)
    
    return fqgj.round(3)

###############################################################################
#将通达信.day读入pands。
#gpdm形如：600027，sh600027,600027.sh
#对于6位数字的代码赚表示A股代码
#对于sh000001上证指数、sz399300沪深300指数则直接        
###############################################################################
def tdxday2pd(gpdm,start=None,end=None):
    if len(gpdm)==8 and gpdm[:2] in ('sh','sz'):
        sc=gpdm[:2]
        gpdm=gpdm[2:]
        dayfn =getdrive()+'\\tdx\\'+sc+'lday\\'+sc+gpdm+'.day'
    else:
        gpdm=sgpdm(gpdm)
        sc = 'sh' if gpdm[0]=='6' else 'sz'
        dayfn =getdrive()+'\\tdx\\'+sc+'lday\\'+sc+gpdm+'.day'

    if os.path.exists(dayfn) :
        return day2pd(dayfn,start,end)
    else :
        return []

###############################################################################
#将通达信.day读入pands
###############################################################################
def day2pd(dayfn,start=None,end=None):
    
    if end == None:
        end=datetime.datetime.now().strftime('%Y%m%d')
    if start == None:
        start='20080101'

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


##########################################################################
#读取高管股东增减持信息
##########################################################################
def read_gggdzjc():

    gpdmb=get_gpdm()
    
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)
    j=1
    for i in range(j,len(gpdmb)):
        gpdm=gpdmb.index[i]
        gpmc = gpdmb.iloc[i]['gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(gpdmb),i+1,gpdm,gpmc)) 
        data = get_gggdcgbd(gpdm)
        
        if len(data)>0 :
            dbcn.executemany('''INSERT OR REPLACE INTO GGGDCGBD (GPDM,BDRQ,BDR,BDSL,BDFX)
            VALUES (?,?,?,?,?)''', data)

        if ((i+1) % 10 ==0) or i>=len(gpdmb)-1 :
            dbcn.commit()

    dbcn.close()

    return

if __name__ == "__main__":  

#def temp():
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

#    read_gggdzjc()

    gggdzjctj()
    
    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

'''
python使用pyquery库总结 
https://blog.csdn.net/baidu_21833433/article/details/70313839

'''
