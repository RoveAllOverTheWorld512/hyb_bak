# -*- coding: utf-8 -*-
"""
Created on Fri Nov  3 12:06:21 2017

@author: lenovo
"""
import getpass
import os
import sys
import re
from configobj import ConfigObj
import time
import datetime
from selenium import webdriver

########################################################################
#初始化本程序配置文件
########################################################################
def iniconfig():
    myname=filename(sys.argv[0])
    wkdir = os.getcwd()
    inifile = os.path.join(wkdir,myname+'.ini')  #设置缺省配置文件
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

#########################################################################
#读取文件名
#########################################################################
def filename(pathname):
    wjm = os.path.splitext(os.path.basename(pathname))
    return wjm[0]

########################################################################
#获取firefoxprofile路径
########################################################################
def getfirefoxprofiledir():
    user = getpass.getuser()
    firefoxprofilepath = 'C:\\Users\\'+user+'\\AppData\\Roaming\\Mozilla\\Firefox'
    inifile = firefoxprofilepath + '\\profiles.ini'
    config = ConfigObj(inifile,encoding='GBK')
    pfnum = config['General']['StartWithLastProfile']
    pfs = 'Profile'+str(pfnum)
    IsRelative = config[pfs]['IsRelative']
    if IsRelative == '0' :               #绝对路径
        return config[pfs]['Path']
    else :
        return firefoxprofilepath +'\\'+ config[pfs]['Path']

###############################################################################
#下载文件名，参数1表示如果文件存在则将原有文件名用其创建时间命名
###############################################################################
def dlfn():
    cus_profile_dir = getfirefoxprofiledir()  # 你自定义profile的路径
    today=datetime.datetime.now().strftime("%Y-%m-%d")
    pzfn = cus_profile_dir + "\\prefs.js"
    with open(pzfn,'rb') as f:
        pzstr = f.read().decode('utf8','ignore')
        f.close()

    if len(pzstr)==0 :
        sys.exit()

    pz = re.findall('download\.dir.{4}(.*)\"',pzstr)
    dldir = pz[0]
    dlfn = today+'.xls'
    fn = os.path.join(dldir,dlfn)

    if os.path.exists(fn):
        ctime=os.path.getctime(fn)  #文件建立时间
        ltime=time.localtime(ctime)
        newfn = time.strftime("%Y%m%d%H%M%S",ltime)+'.xls'
        os.rename(fn,os.path.join(os.path.dirname(fn),newfn))

    return fn


def tmp():
    config = iniconfig()
    cus_profile_dir = getfirefoxprofiledir()  # 你自定义profile的路径
    username = readkey(config,'iwencaiusername')
    pwd = readkey(config,'iwencaipwd')
    sn = int(readkey(config,'sn'))
    qn = int(readkey(config,'qn'))
    
    browser = webdriver.PhantomJS()
    
        
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
        
    for j in range(sn,qn+1):

        dafn = dlfn()
        newfn = os.path.splitext(dafn)[0]+"-"+str(j)+".xls"            
        if os.path.exists(newfn):
            os.remove(newfn)
        
        kw = readkey(config,'iwencaikw'+str(j))
        sele = readkey(config,'iwencaiselestr'+str(j))
        browser.get("http://www.iwencai.com/")
        time.sleep(5)
        browser.find_element_by_id("auto").clear()
        browser.find_element_by_id("auto").send_keys(kw)
        browser.find_element_by_id("qs-enter").click()
        time.sleep(10)
        
        #打开查询项目选单
        trigger = browser.find_element_by_class_name("showListTrigger")
        trigger.click()
        time.sleep(1)
        
        #获取查询项目选单
        checkboxes = browser.find_elements_by_class_name("showListCheckbox")
        indexstrs = browser.find_elements_by_class_name("index_str")
        
        #去掉选项前的“√”
        #涨幅、股价保留
        for i in range(0,len(checkboxes)):
            checkbox=checkboxes[i]
            indexstr=indexstrs[i].text.replace(",","_")
            if checkbox.is_selected() and not indexstr in sele :
                checkbox.click()
            if not checkbox.is_selected() and indexstr in sele :
                checkbox.click()

        #向上滚屏
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

        if os.path.exists(dafn):
            newfn = os.path.splitext(dafn)[0]+"-"+str(j)+".xls"
            os.rename(dafn,newfn)
       
        
    #关闭浏览器
    browser.quit()


if __name__ == '__main__':
    driver = webdriver.PhantomJS()  
    driver.get("http://www.baidu.com")  
    data = driver.title 
    print(data)