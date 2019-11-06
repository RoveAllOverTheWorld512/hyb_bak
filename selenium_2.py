# -*- coding: utf-8 -*-
"""
Created on Sun Feb 19 13:36:39 2017

@author: Lenovo
"""
import os
import sys
import re
import datetime
import time
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

def dlfn(prof):
    today=datetime.datetime.now().strftime("%Y-%m-%d")

    cus_profile_dir = prof  # 你自定义profile的路径
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
    return fn

print(1,datetime.datetime.now().strftime("%H:%M:%S"))
config = iniconfig()
print(2,datetime.datetime.now().strftime("%H:%M:%S"))
cus_profile_dir = readkey(config,'firefox_profiledir')  # 你自定义profile的路径
username = readkey(config,'iwencaiusername')
print(3,datetime.datetime.now().strftime("%H:%M:%S"))
pwd = readkey(config,'iwencaipwd')
kw = readkey(config,'iwencaikw')
print(4,datetime.datetime.now().strftime("%H:%M:%S"))

fn = dlfn(cus_profile_dir)

if os.path.exists(fn):
    ctime=os.path.getctime(fn)  #文件建立时间
    ltime=time.localtime(ctime)
    newfn = time.strftime("%Y%m%d%H%M%S",ltime)+'.xls'
    os.rename(fn,os.path.join(os.path.dirname(fn),newfn))

print(5,datetime.datetime.now().strftime("%H:%M:%S"))

cus_profile = webdriver.FirefoxProfile(cus_profile_dir)
print(6,datetime.datetime.now().strftime("%H:%M:%S"))
browser = webdriver.Firefox(cus_profile)
#browser = webdriver.Firefox()
#print(7,datetime.datetime.now().strftime("%H:%M:%S"))

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
#获取查询项目选单
checkboxes = browser.find_elements_by_class_name("showListCheckbox")
#去掉选项前的“√”
for checkbox in checkboxes:
    if checkbox.is_selected():
        checkbox.click()
    time.sleep(1)
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

#关闭浏览器
#    browser.quit()
