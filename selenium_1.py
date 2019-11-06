# -*- coding: utf-8 -*-
"""
Created on Tue Feb  7 22:25:24 2017

@author: lenovo
"""
import time
from pyquery import PyQuery as pq
from selenium import webdriver

#browser = webdriver.Ie()
browser = webdriver.Firefox()
#browser = webdriver.Chrome()
#browser.implicitly_wait(30)
#browser = webdriver.Firefox()

#浏览器窗口最大化
browser.maximize_window()
#登录同花顺
browser.get("http://www.cninfo.com.cn/information/companyinfo_n.html")
#time.sleep(1)
elem = browser.find_element_by_id("fhpg")
elem.click()
time.sleep(2)

browser.switch_to_frame("i_nr")
elem = browser.find_element_by_id("stockID_")
elem.clear()
elem.send_keys('002294')

elem = browser.find_element_by_class_name("input2")
elem.click()
time.sleep(2)

html = browser.find_element_by_xpath("//*").get_attribute("outerHTML")
# 不要用 browser.page_source，那样得到的页面源码不标准
html = pq(html)
html.find("script").remove()    # 清理 <script>...</script>
html.find("style").remove()     # 清理 <style>...</style>

tb = html('tr')

for row in range(3,len(tb)+1):
    tr=pq(tb.eq(row).html())
    fhnd=tr.find('td').eq(0).text()
    fhfa=tr.find('td').eq(1).text()
    gqdjr=tr.find('td').eq(2).text()
    cqjzr=tr.find('td').eq(3).text()
    hgssr=tr.find('td').eq(4).text()
    print(fhnd,fhfa,gqdjr,cqjzr,hgssr)

