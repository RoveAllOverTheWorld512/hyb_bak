# -*- coding: utf-8 -*-
"""
Created on Sat Feb  4 00:33:48 2017
查询市盈率
@author: huangyunbin@sina.com
"""

import os
import sys
import re
import datetime
from configobj import ConfigObj
import xlrd
import sqlite3

########################################################################
#建立数据库
########################################################################
def createDataBase():
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    cn = sqlite3.connect(dbfn)

    cn.execute('''CREATE TABLE IF NOT EXISTS PE_PB
           (GPDM TEXT,
           RQ TEXT,
           PE_LYR REAL,
           PE_TTM REAL,
           PB REAL);''')

    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ_PEPB ON PE_PB(GPDM,RQ);''')

    cn.execute('''CREATE TABLE IF NOT EXISTS GPHY
           (GPDM TEXT,
           RQ TEXT,
           HYDM TEXT);''')

    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ_GPHY ON GPHY(GPDM,RQ);''')

    cn.execute('''CREATE TABLE IF NOT EXISTS GPDMB
           (GPDM TEXT PRIMARY KEY,
           GPMC TEXT,
           SSRQ TEXT);''')


########################################################################
#初始化本程序配置文件
########################################################################
def iniconfig():
    inifile = os.path.splitext(sys.argv[0])[0]+'.ini'  #设置缺省配置文件
    return ConfigObj(inifile,encoding='GBK')


########################################################################
#读取个股静态市盈率、滚动市盈率、市净率
########################################################################
def ggsyl(file,date):
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name("个股数据")
    nrows = table.nrows #行数
    data =[]
    for rownum in range(1,nrows):
        row = table.row_values(rownum)
        dm=row[0]
        dm=dm+('.SH' if dm[0]=='6' else '.SZ')
        data.append((dm,date,row[10],row[11],row[12]))
    return data

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

##########################################################################
#用法
##########################################################################
def Usage():
    print ('用法:')
    print ('-h, --help: 显示帮助信息。')
    print ('-v, --version: 显示版本信息。')
    print ('-i, --input: 股票列表文本文件。')
    print ('-o, --output: 市盈率保存文件。')

##########################################################################
#版本
##########################################################################
def Version():
    print ('版本 2.0.0')

def makedir(dirname):
    if not os.path.exists(dirname):
        try :
            os.mkdir(dirname)
        except(OSError):
            print("创建目录%s出错，请检查！" % dirname)
            return False
    else :
        return True

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


if __name__ == '__main__':
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    config = iniconfig()
    lastdate = readkey(config,'lastdate')
    createDataBase()

    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    syldir = getdrive()+'\\syl'

    jyrlst = jyrlist(syldir)
    if lastdate != '' :
        jyrlst = [e for e in jyrlst if e>lastdate]

    if len(jyrlst) >0 :

        lastdate = jyrlst[0]

        data=[]
        i=1
        for jyrq in jyrlst:

            sylfn = os.path.join(syldir,"csi"+jyrq+".xls")

            print('共有%d个文件，正在处理第%d个文件：%s，请等待。' % (len(jyrlst),i,jyrq))

            ggsj = ggsyl(sylfn,jyrq[0:4]+"-"+jyrq[4:6]+"-"+jyrq[6:8])

            data.extend(ggsj)
            i += 1
    
        dbcn.executemany('INSERT OR IGNORE INTO PE_PB (GPDM,RQ,PE_LYR,PE_TTM,PB) VALUES (?,?,?,?,?)', data)
        dbcn.commit()

        config['lastdate'] = lastdate
        config.write()

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

    dbcn.close()
    

