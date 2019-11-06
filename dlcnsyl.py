# -*- coding: utf-8 -*-
"""
Created on Tue Jan 17 14:30:05 2017

@author: lenovo
"""

import os
import sys
import getopt
import requests
import zipfile
import datetime
from configobj import ConfigObj

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
#从配置文件中读取休市日期
#########################################################################
def readclosedate(config):
    keys = config.keys()
    if keys.count('stockclosedate') :
        return eval(config['stockclosedate'])
    else :
        return []

#########################################################################
#读取键值,如果键值不存在，就设置为defvl
#########################################################################
def readkey(config,key,defvl=None):
    keys = config.keys()

    if not keys.count(key) :
        if defvl != None:
            config[key] = defvl
            config.write()
        return defvl

    return config[key]


#######################################################################
#下载板块市盈率文件，参数为日期“20170108”表示2017年1月8日
#######################################################################
def download_syl(sdate,cwd,zl=""):
    urlpath = 'http://115.29.204.48/syl/'
    ifn = zl+sdate+'.zip'
    ofn = zl+sdate+'.xls'
    zip_file = cwd + "\\" + ifn
    xls_file = cwd + "\\" + ofn

    url = urlpath+ifn
    if not os.path.exists(zip_file) and not os.path.exists(xls_file):
        r = requests.get(url)
#如果下载文件不存在 ，r返回 <Response [404]>， r.ok为False
#如果下载文件存在 ，r返回 <Response [200]>，r.ok为True
        if not r.ok :
            print ("你所下载的文件%s不存在！" % ifn)

        if r.ok and not os.path.exists(zip_file):
            print ("正在下载的文件%s，请等待！" % ifn)
            with open(zip_file, "wb") as code:
                code.write(r.content)


    if os.path.exists(zip_file) and not os.path.exists(xls_file):
            f_zip = zipfile.ZipFile(zip_file, 'r')
            f_zip.extractall(cwd)
            f_zip.close()
            os.remove(zip_file)

##########################################################################
#将字符串转换为时间戳，不成功返回errdate
##########################################################################
def str2datetime(s):
    try:
        dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
    except(ValueError):
        dt = "errdate"
    return dt

##########################################################################
#n天后日期串，不成功返回errdate
##########################################################################
def nextdtstr(s,n):
    dt = str2datetime(s)
    if dt != "errdate" :
        dt += datetime.timedelta(n)
        return dt.strftime("%Y%m%d")
    else :
        return "errdate"

def Usage():
    print ('用法:')
    print ('-h, --help: 显示帮助信息。')
    print ('-v, --version: 显示版本信息。')
    print ('-c, --cfg: 配置文件。')
    print ('-b, --bgdate: 起始日期，“20170118”表示查询2017年1月18日市盈率。')
    print ('-e, --eddate: 截止日期，“20170118”表示查询2017年1月18日市盈率。')

def Version():
    print ('版本 2.0.0')

def makedir(dirname):
    if dirname == None :
        return False

    if not os.path.exists(dirname):
        try :
            os.mkdir(dirname)
        except(OSError):
            print("创建目录%s出错，请检查！" % dirname)
            return False
    else :
        return True

def filename(pathname):
    wjm = os.path.splitext(os.path.basename(pathname))
    return wjm[0]

def main(argv):

    try:
        opts, args = getopt.getopt(argv[1:], 'hvc:b:e:', ['help','version','cfg=','bgdate=','eddate='])
    except (getopt.GetoptError):
        Usage()
        sys.exit(1)

#当程序不带参数运行时，查询日期为当日，查询股票列表为通达信自选股，查询结果保存在syl.xls中
    td = datetime.datetime.now().strftime("%Y%m%d") #今天
    eddate = td
    bgdate = (datetime.datetime.now() - datetime.timedelta(10)).strftime("%Y%m%d")
    inifile = ""
    iniyes = False

    for o, a in opts:
        if o in ('-h', '--help'):
            Usage()
            sys.exit(0)
        elif o in ('-v', '--version'):
            Version()
            sys.exit(0)
        elif o in ('-c', '--cfg'):
            inifile = a
            iniyes = True
        elif o in ('-b', '--bgdate'):
            bgdate = a
        elif o in ('-e', '--eddate'):
            eddate = a
        else:
            print ('无效参数！')
            sys.exit(3)

    wkdir = os.getcwd()
    diskroot = wkdir[:2]+"\\"
    if iniyes and not os.path.exists(inifile) :
        print("配置文件%s不存在，请检查。" % inifile)
        iniyes = False
    if not iniyes:
        config = iniconfig()
        if not os.path.exists(config.filename):
            config['pe'] = os.path.join(diskroot,"pe")
            config['syl'] = os.path.join(diskroot,"syl")
            config['bksyl'] = os.path.join(diskroot,"bksyl")
            config.write()
    else:
        config = readini(inifile)

    pedir = readkey(config,'pe')
    if len(pedir)>0:
        if not makedir(pedir):
            pedir = wkdir
    else:
        pedir = wkdir

    syldir = readkey(config,'syl')
    if len(syldir)>0:
        if not makedir(syldir):
            syldir = wkdir
    else:
        syldir = wkdir


    bksyldir = readkey(config,'bksyl')
    if len(bksyldir)>0:
        if not makedir(bksyldir):
            bksyldir = wkdir
    else:
        syldir = wkdir

    closedate = readclosedate(config)
    if str2datetime(bgdate) == "errdate" or str2datetime(eddate) == "errdate" :
        print("指定日期%s、%s无效，请检查！" % (bgdate,eddate))
        sys.exit(3)

    if str2datetime(eddate)>str2datetime(td):
        eddate = td

    if str2datetime(bgdate) > str2datetime(eddate):
        bgdate, eddate = eddate, bgdate

    jyrq = bgdate
    while str2datetime(jyrq) <= str2datetime(eddate):
        if (not jyrq in closedate) and str2datetime(jyrq).weekday()<5 :
            if jyrq==td :
                if datetime.datetime.now().hour>=18 :
                    download_syl(jyrq,pedir,'')
                    download_syl(jyrq,syldir,'csi')
                    download_syl(jyrq,bksyldir,'bk')
            else:
                download_syl(jyrq,pedir,'')
                download_syl(jyrq,syldir,'csi')
                download_syl(jyrq,bksyldir,'bk')

        jyrq = nextdtstr(jyrq,1)

    print("下载文件保存目录：%s、%s、%s" % (pedir,syldir,bksyldir))

if __name__ == '__main__':
    main(sys.argv)

