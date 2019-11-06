# -*- coding: utf-8 -*-
"""
Created on Tue May  1 12:14:20 2018

@author: lenovo
python中的selenium中的鼠标悬停事件！
https://blog.csdn.net/qq_39248703/article/details/75533117

玩转python selenium鼠标键盘操作（ActionChains）
http://www.jb51.net/article/92682.htm

Selenium2+python自动化 18种定位方法（find_elements）
https://www.cnblogs.com/yoyoketang/p/6557421.html

"""

from selenium import webdriver
import time 

'''
注意:ddir不能是中文目录，中文不起作用
'''

gpdm='601558'
ddir = r'D:\report'
url="http://basic.10jqka.com.cn/%s/finance.html#stockpage" % gpdm

profile = webdriver.FirefoxProfile()  #注意:ddir不能是中文目录，中文不起作用
profile.set_preference('browser.download.dir', ddir)
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.manager.showWhenStarting', False)

#http://www.w3school.com.cn/media/media_mimeref.asp
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'application/vnd.ms-excel')

browser = webdriver.Firefox(firefox_profile=profile)

#浏览器窗口最大化
browser.maximize_window()
            
for cls in ('icons_main','icons_page','icons_pie','icons_coin'):
    browser.get(url)
    time.sleep(3)

    print(cls)
    
    elem = browser.find_element_by_class_name(cls)
    elem.click()
    time.sleep(3)
    
    tbls = browser.find_elements_by_class_name("tableTab")
    tbls[1].click()     #0报告期，1年度，2单季度
    time.sleep(3)
    
    trigger = browser.find_element_by_class_name("export_data")
    trigger.click()
    time.sleep(3)   #很重要，没有就不会执行下载操作

browser.get(url)
time.sleep(3)
elem = browser.find_element_by_class_name("icons_main")
for i in range(4):
    #下面的js非常重要
    menubox = browser.find_element_by_class_name("menubox")
    js = '	$(".menubox").css("display","block");' #编写jQuery语句
    browser.execute_script(js) #执行JS
    
    ls = menubox.find_elements_by_tag_name("a")
    
    print(ls[i].get_attribute('innerHTML'))  

    ls[i].click()
    time.sleep(3)

    tbls = browser.find_elements_by_class_name("tableTab")
    tbls[1].click()     #0报告期，1年度，2单季度
    time.sleep(3)

    
    trigger = browser.find_element_by_class_name("export_data")
    trigger.click()
    time.sleep(3)   #很重要，没有就不会执行下载操作

browser.quit()
