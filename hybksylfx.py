# -*- coding: utf-8 -*-
"""
Created on Wed Mar  1 22:07:07 2017

@author: Lenovo
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
import xlrd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
import pandas as pd
from configobj import ConfigObj
import statsmodels.api as sm
import xlwt
import winreg
from dateutil import parser

##########################################################################
#获取运行程序所在驱动器
##########################################################################
def getdrive():
    return os.path.splitdrive(sys.argv[0])[0]

##########################################################################
#获取运行程序所在路径
##########################################################################
def getpath():
    return os.path.dirname(sys.argv[0])

##########################################################################
#创建目录
##########################################################################
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

########################################################################
#行业代码表
########################################################################
def hydmdict():
    fn = 'hydmb.txt'
    with open(fn) as f:
        hydmb = f.read()
        f.close()

    dmb = re.findall('(\d+)\t(.+)\n',hydmb)
    dm = {}
    for (hydm,hymc) in dmb :
        dm[hydm] = hymc

    return dm


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
                value = value.strip(b'\x00').decode('GBK').replace(' ','')
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

def dbf2pandas(dbffn,cols=[]):
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

def bkfx():
    hybksyl = dbf2pandas('hybksyl.dbf')
    hybksyl['rq'] = hybksyl['rq'].map(str2datetime)
    hybksyl['hydm'] = hybksyl['hydm']
    hybksyl = hybksyl.set_index('hydm')

    hsagsyl = hybksyl.ix['HSAG',['syl','rq']]
    hsagsyl=hsagsyl.set_index('rq')
    hsagsyl.rename(columns={'syl':'hsagpe'}, inplace = True)

    cybsyl = hybksyl.ix['CYB',['syl','rq']]
    cybsyl = cybsyl.set_index('rq')
    cybsyl.rename(columns={'syl':'cybpe'}, inplace = True)
    syl1 = pd.merge(hsagsyl,cybsyl,left_index = True, right_index = True)

    zxbsyl = hybksyl.ix['ZXB',['syl','rq']]
    zxbsyl = zxbsyl.set_index('rq')
    zxbsyl.rename(columns={'syl':'zxbpe'}, inplace = True)
    syl2 = pd.merge(syl1,zxbsyl,left_index = True, right_index = True)
    syl2.eval('zxbbz=zxbpe/hsagpe')
    syl2.eval('cybbz=cybpe/hsagpe')

    shsyl = hybksyl.ix['SHAG',['syl','rq']]
    shsyl = shsyl.set_index('rq')
    shsyl.rename(columns={'syl':'shpe'}, inplace = True)
    syl3 = pd.merge(syl2,shsyl,left_index = True, right_index = True)

    szsyl = hybksyl.ix['SZAG',['syl','rq']]
    szsyl = szsyl.set_index('rq')
    szsyl.rename(columns={'syl':'szpe'}, inplace = True)
    syl4 = pd.merge(syl3,szsyl,left_index = True, right_index = True)

    szzbsyl = hybksyl.ix['SSZB',['syl','rq']]
    szzbsyl = szzbsyl.set_index('rq')
    szzbsyl.rename(columns={'syl':'szzbpe'}, inplace = True)
    syl5 = pd.merge(syl4,szzbsyl,left_index = True, right_index = True)
#    syl5 = syl5['20130101':'20141231']
    syl5.describe()

    font = FontProperties(fname=r"c:\windows\fonts\simhei.ttf", size=14)

    fig, ax1 = plt.subplots(figsize=(18,6))
    ax2 = ax1.twinx()

    ax1.plot(syl5.index,syl5['hsagpe'],color="b",linewidth=2,label='沪深A股PE')
    ax1.plot(syl5.index,syl5['zxbpe'],color="r",linewidth=2,label='中小板PE')
    ax1.plot(syl5.index,syl5['cybpe'],color="g",linewidth=2,label='创业板PE')
#    ax1.plot(syl5.index,syl5['shpe'],color="c",linewidth=2,label='沪市A股PE')
#    ax1.plot(syl5.index,syl5['szpe'],color="m",linewidth=2,label='深市A股PE')
#    ax1.plot(syl5.index,syl5['szzbpe'],color="m",linewidth=2,label='深市主板PE')

    ax2.plot(syl5.index,syl5['zxbbz'],color="r",linestyle=':',linewidth=1,label='中小板PE/沪深A股PE')
    ax2.plot(syl5.index,syl5['cybbz'],color="g",linestyle=':',linewidth=1,label='创业板PE/沪深A股PE')

    title = "2013年1月1日至2017年3月3 日沪深A股、中小板、创业板市盈率及中小板创业板与沪深A股市盈率比值走势图"
    fig.suptitle(title, fontproperties=font,fontsize = 14, fontweight='bold')

    ax1.set_xlabel('日期', fontproperties=font,fontsize = 12)
    ax1.set_ylabel('滚动市盈率', color="blue", fontproperties=font, fontsize = 12)
    ax1.legend(loc='upper left', prop={'family':'SimHei','size':10})
    ax1.tick_params('y', colors='b')
    ax1.grid(True,color='b')

    ax2.set_ylabel('中小板、创业板市盈率/沪深A股市盈率比值',color="red", fontproperties=font,fontsize = 16)
    ax2.legend(loc='upper right', prop={'family':'SimHei','size':10})
    ax2.tick_params('y', colors='red')
    ax2.grid(True,color='r')

#    imgfn = imgdir+'\\'+gpdmdict()[stock].replace(' ','').replace('*','')+'.png'
#    plt.savefig(imgfn)

    plt.show()

##########################################################################
#将字符串转换为时间戳，不成功返回None
##########################################################################
def str2datetime(s):
    try:
        dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
    except :
        dt = None
    return dt

def hyfx():
    hydmlb = hydmdict()

    hybksyl = dbf2pandas('hybksyl.dbf')
    hybksyl['rq'] = hybksyl['rq'].map(str2datetime)
    hybksyl['hydm'] = hybksyl['hydm']
    hybksyl = hybksyl.set_index('hydm')

    hsagsyl = hybksyl.ix['HSAG',['syl','rq']]
    hsagsyl=hsagsyl.set_index('rq')
    hsagsyl.rename(columns={'syl':'hsagpe'}, inplace = True)
    
    for i in range(1,5):
        print('正在处理%i级行业，请稍后。' % i)
        imgdir = getpath()+"/bkfx/"+str(i)
        makedir(imgdir)

        hydmb = list(set([e for e in hybksyl.index if len(e)==2*i]))
        for hydm in hydmb :
            hype = hybksyl.ix[hydm,['syl','rq']]
            if len(hype)<200 or (hydm not in hydmlb.keys()) :
                continue
            hype=hype.set_index('rq')
            hype.rename(columns={'syl':'hype'}, inplace = True)
            syl = pd.merge(hsagsyl,hype,left_index = True, right_index = True)
            syl.eval('hybz=hype/hsagpe')
    
            font = FontProperties(fname=r"c:\windows\fonts\simhei.ttf", size=14)
        
            fig, ax1 = plt.subplots(figsize=(18,6))
            ax2 = ax1.twinx()
            ax1.plot(syl.index,syl['hsagpe'],color="b",linewidth=2,label='沪深A股PE')
            ax1.plot(syl.index,syl['hype'],color="r",linewidth=2,label=hydmlb[hydm]+'行业PE')
            ax2.plot(syl.index,syl['hybz'],color="r",linestyle=':',linewidth=1,label='行业板块PE/沪深A股PE')
            title = hydm+hydmlb[hydm]+"板块市盈率及行业板块市盈率与沪深A股市盈率比值走势图"
            fig.suptitle(title, fontproperties=font,fontsize = 14, fontweight='bold')
        
            ax1.set_xlabel('日期', fontproperties=font,fontsize = 12)
            ax1.set_ylabel('滚动市盈率', color="blue", fontproperties=font, fontsize = 12)
            ax1.legend(loc='upper left', prop={'family':'SimHei','size':10})
            ax1.tick_params('y', colors='b')
            ax1.grid(True,color='b')
        
            ax2.set_ylabel('行业板块板市盈率/沪深A股市盈率比值',color="red", fontproperties=font,fontsize = 16)
            ax2.legend(loc='upper right', prop={'family':'SimHei','size':10})
            ax2.tick_params('y', colors='red')
            ax2.grid(True,color='r')
            imgfn = imgdir+'/'+hydm+hydmlb[hydm]+'.png'
            plt.savefig(imgfn)
            plt.show()
           
            
if __name__ == '__main__':
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
            
    hybksyl = dbf2pandas('hybksyl.dbf')
    hybksyl['rq'] = hybksyl['rq'].map(str2datetime)
    hybksyl['hydm'] = hybksyl['hydm']
    hybksyl = hybksyl.set_index('hydm')

    hsagsyl = hybksyl.ix['HSAG',['syl','rq']]
    hsagsyl=hsagsyl.set_index('rq')
    hsagsyl.rename(columns={'syl':'hsagpe'}, inplace = True)

    cybsyl = hybksyl.ix['CYB',['syl','rq']]
    cybsyl = cybsyl.set_index('rq')
    cybsyl.rename(columns={'syl':'cybpe'}, inplace = True)
    syl1 = pd.merge(hsagsyl,cybsyl,left_index = True, right_index = True)

    zxbsyl = hybksyl.ix['ZXB',['syl','rq']]
    zxbsyl = zxbsyl.set_index('rq')
    zxbsyl.rename(columns={'syl':'zxbpe'}, inplace = True)
    syl2 = pd.merge(syl1,zxbsyl,left_index = True, right_index = True)
    syl2.eval('zxbbz=zxbpe/hsagpe')
    syl2.eval('cybbz=cybpe/hsagpe')

    shsyl = hybksyl.ix['SHAG',['syl','rq']]
    shsyl = shsyl.set_index('rq')
    shsyl.rename(columns={'syl':'shpe'}, inplace = True)
    syl3 = pd.merge(syl2,shsyl,left_index = True, right_index = True)

    szsyl = hybksyl.ix['SZAG',['syl','rq']]
    szsyl = szsyl.set_index('rq')
    szsyl.rename(columns={'syl':'szpe'}, inplace = True)
    syl4 = pd.merge(syl3,szsyl,left_index = True, right_index = True)

    szzbsyl = hybksyl.ix['SSZB',['syl','rq']]
    szzbsyl = szzbsyl.set_index('rq')
    szzbsyl.rename(columns={'syl':'szzbpe'}, inplace = True)
    syl5 = pd.merge(syl4,szzbsyl,left_index = True, right_index = True)
#    syl5 = syl5['20130101':'20141231']
    syl5.describe()

    font = FontProperties(fname=r"c:\windows\fonts\simhei.ttf", size=14)

    fig, ax1 = plt.subplots(figsize=(18,6))
    ax2 = ax1.twinx()

    ax1.plot(syl5.index,syl5['hsagpe'],color="b",linewidth=2,label='沪深A股PE')
    ax1.plot(syl5.index,syl5['zxbpe'],color="r",linewidth=2,label='中小板PE')
    ax1.plot(syl5.index,syl5['cybpe'],color="g",linewidth=2,label='创业板PE')
#    ax1.plot(syl5.index,syl5['shpe'],color="c",linewidth=2,label='沪市A股PE')
#    ax1.plot(syl5.index,syl5['szpe'],color="m",linewidth=2,label='深市A股PE')
#    ax1.plot(syl5.index,syl5['szzbpe'],color="m",linewidth=2,label='深市主板PE')

    ax2.plot(syl5.index,syl5['zxbbz'],color="r",linestyle=':',linewidth=1,label='中小板PE/沪深A股PE')
    ax2.plot(syl5.index,syl5['cybbz'],color="g",linestyle=':',linewidth=1,label='创业板PE/沪深A股PE')

    title = "2013年1月1日至2017年3月3 日沪深A股、中小板、创业板市盈率及中小板创业板与沪深A股市盈率比值走势图"
    fig.suptitle(title, fontproperties=font,fontsize = 14, fontweight='bold')

    ax1.set_xlabel('日期', fontproperties=font,fontsize = 12)
    ax1.set_ylabel('滚动市盈率', color="blue", fontproperties=font, fontsize = 12)
    ax1.legend(loc='upper left', prop={'family':'SimHei','size':10})
    ax1.tick_params('y',direction='in', colors='b')
    ax1.grid(True,color='b')

    ax2.set_ylabel('中小板、创业板市盈率/沪深A股市盈率比值',color="red", fontproperties=font,fontsize = 16)
    ax2.legend(loc='upper right', prop={'family':'SimHei','size':10})
    ax2.tick_params('y', colors='red')
    ax2.grid(True,color='r')

#    imgfn = imgdir+'\\'+gpdmdict()[stock].replace(' ','').replace('*','')+'.png'
#    plt.savefig(imgfn)

    plt.show()

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

