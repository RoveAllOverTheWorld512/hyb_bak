# -*- coding: utf-8 -*-
"""
Created on Mon Jan 30 11:18:09 2017

@author: huangyunbin@sina.com
"""

import os
import sys
import getopt
import re
import datetime
import time
import struct
import winreg
from configobj import ConfigObj
from selenium import webdriver

########################################################################
#初始化本程序配置文件
########################################################################
def iniconfig():
    inifile = os.path.splitext(sys.argv[0])[0]+'.ini'  #设置缺省配置文件
    return ConfigObj(inifile,encoding='GBK')

#########################################################################
#读取键值
#########################################################################
def readkey(config,key):
    keys = config.keys()
    if keys.count(key) :
        return config[key]
    else :
        return ""

###############################################################################
#下载文件名，参数1表示如果文件存在则将原有文件名用其创建时间命名
###############################################################################
def dlfn(n=0):
    config = iniconfig()
    cus_profile_dir = readkey(config,'firefox_profiledir')  # 你自定义profile的路径
    today=datetime.datetime.now().strftime("%Y-%m-%d")
    pzfn = cus_profile_dir + "\\prefs.js"
    with open(pzfn) as f:
        pzstr = f.read()
        f.close()

    if len(pzstr)==0 :
        sys.exit()

    pz = re.findall('download\.dir.{4}(.*)\"',pzstr)
    dldir = pz[0]
    dlfn = today+'.xls'
    fn = os.path.join(dldir,dlfn)
    
    if n==1 :
        if os.path.exists(fn):
            ctime=os.path.getctime(fn)  #文件建立时间
            ltime=time.localtime(ctime)
            newfn = time.strftime("%Y%m%d%H%M%S",ltime)+'.xls'
            os.rename(fn,os.path.join(os.path.dirname(fn),newfn))
        create_html()
        
    return fn


def create_html():
    config = iniconfig()
    cus_profile_dir = readkey(config,'firefox_profiledir')  # 你自定义profile的路径
    username = readkey(config,'iwencaiusername')
    pwd = readkey(config,'iwencaipwd')
    kw = readkey(config,'iwencaikw')

    cus_profile = webdriver.FirefoxProfile(cus_profile_dir)
    browser = webdriver.Firefox(cus_profile)

    #browser.implicitly_wait(30)
    #browser = webdriver.Firefox()

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

    #查询“i问财”网
    #kw="连续2年主营业务收入增长率>10%,连续2年净利润增长率>10%，2016年9月30日主营业务收入增长率>10% 2016年12月31日业绩预增 医药股"
    #kw="3季度营业收入同比增长率>10% 净利润同比增长率>10% 净利润同比增长率>营业收入同比增长率 经营活动现金流>购建固定资产、无形资产和其他长期资产支付的现金 2016年12月31日业绩预增 2014年1月1日前上市"
    browser.get("http://www.iwencai.com/")
    time.sleep(5)
    browser.find_element_by_id("auto").clear()
    browser.find_element_by_id("auto").send_keys(kw)
    browser.find_element_by_id("qs-enter").click()
    time.sleep(20)

    #打开查询项目选单
    trigger = browser.find_element_by_class_name("showListTrigger")
    trigger.click()
    time.sleep(1)
#    获取查询项目选单
    checkboxes = browser.find_elements_by_class_name("showListCheckbox")
#    去掉选项前的“√”
    for checkbox in checkboxes:
        if checkbox.is_selected():
            checkbox.click()
            time.sleep(1)
#    向上滚屏
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
#    browser.quit()


def gettdxblkdir():
    try :
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人生")
        value, type = winreg.QueryValueEx(key, "InstallLocation")
        return value + '\\T0002\\blocknew'
    except :
        print("本机未安装【华西证券华彩人生】软件系统。")
        sys.exit()


def cutstr(s):
    i=0
    while s[i]!=0:
        i += 1
    return s[:i]


def readblockcfg():
    blkcfgfn=gettdxblkdir() + '\\blocknew.cfg'
    with open(blkcfgfn,'rb') as f:
        blks=f.read()
    f.close()
    blknum = int(len(blks)/120)
    blklst = []
    for i in range(blknum):
        blkstr=blks[i*120:(i+1)*120]
        blkname,shtname = struct.unpack('50s70s',blkstr)
        blkname = cutstr(blkname).decode('GBK')
        shtname = cutstr(shtname).decode('GBK')
        blklst.append([blkname,shtname])
    return blklst


def writeblock(stklst,blkname):
    blkdir=gettdxblkdir()
    blkfn = blkdir+'\\'+blkname+'.blk'
    with open(blkfn,'w') as f:
        for s in stklst:
            f.write(s+"\r\n")
    f.close()


def writeblockcfg(blklst):
    blkcfgfn=gettdxblkdir() + '\\blocknew.cfg'
    with open(blkcfgfn,'wb') as f:
        for blkname,shtname in blklst:
            blkname = blkname.ljust(50, '\x00').encode('GBK')
            shtname = shtname.ljust(70, '\x00').encode('GBK')
            f.write(struct.pack('50s70s', blkname, shtname))
    f.close()


def Usage():
    print ('用法:')
    print ('-h, --help: 显示帮助信息。')
    print ('-v, --version: 显示版本信息。')
    print ('-d, --downloadn: 下载数据。')
    print ('-b, --block: 板块名称。')
    print ('-s, --short: 板块简称。')

def Version():
    print ('版本 1.0.0')

def main(argv):
    try:
        opts, args = getopt.getopt(argv[1:], 'hvdb:s:', ['help','version','download','block=','short='])
    except (getopt.GetoptError):
        Usage()
        sys.exit(1)

#当程序不带参数运行时，查询日期为当日，查询股票列表为通达信自选股，查询结果保存在syl.xls中
    n=0
    config = iniconfig()
    blkname = readkey(config,'iwencaip2blkname')
    shtname = readkey(config,'iwencaip2shtname')
    stklst = []
    for o, a in opts:
        if o in ('-h', '--help'):
            Usage()
            sys.exit(0)
        elif o in ('-v', '--version'):
            Version()
            sys.exit(0)
        elif o in ('-d', '--download'):
            n = 1
        elif o in ('-b', '--block'):
            blkname = a
            if len(blkname)==0 :
                print("板块名称不能空！")
                sys.exit(3)

        elif o in ('-s', '--short'):
            shtname = a.upper()
            if not shtname.isalnum() :
                print("板块简称必须是英文字母或数字！")
                sys.exit(3)
            if len(shtname)==0 :
                print("板块简称不能空！")
                sys.exit(3)
        else:
            print ('无效参数！')
            sys.exit(3)

    fn = dlfn(n)
    if os.path.exists(fn) :
        #用二进制方式打开再转成字符串，可以避免直接打开转换出错
        with open(fn,'rb') as dtf:
            zxg = dtf.read().decode('utf8','ignore')
        zxglst =re.findall("(\d{6})",zxg)
        for s in zxglst:
            if s[0]=='6':
                s = '1'+s
            else:
                s = '0'+s
            stklst.append(s)
    else:
        print("文件%s不存在！" % fn)
        sys.exit()
    if len(stklst)==0:
        print("股票列表为空,请检查%s文件。" % a)
        sys.exit(3)

    if len(stklst)==0 or len(blkname)==0 or len(shtname)==0 :
        print("股票列表为空或没有指定板块名称、简称，请检查。")
        sys.exit(3)
    else :
        blklst = readblockcfg()
        if shtname in [e[1] for e in blklst] :
            prompt = "板块简称"+shtname+"已存在，可以重写吗？输入Y将重写："
            ask = input(prompt)
            if ask.upper() == "Y" :
                writeblock(stklst,shtname)
            else :
                sys.exit(3)
        else :
            blklst.append([blkname,shtname])
            writeblockcfg(blklst)
            writeblock(stklst,shtname)


if __name__ == '__main__':
    main(sys.argv)

