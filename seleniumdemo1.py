# -*- coding: utf-8 -*-
"""
Created on Tue Apr 10 08:57:33 2018

https://blog.csdn.net/pushiqiang/article/details/51290509
https://www.cnblogs.com/zhaof/p/6953241.html

"""

from selenium import webdriver
driver = webdriver.PhantomJS()

driver.get("http://news.sohu.com/scroll/")

#或得js变量的值
r = driver.execute_script("return newsJason")
print(r)

#selenium在webdriver的DOM中使用选择器来查找元素，名字直接了当，by对象可使用的选择策略有：id,class_name,css_selector,link_text,name,tag_name,tag_name,xpath等等
#print(driver.find_element_by_tag_name("div").text)
#print(driver.find_element_by_css_selector("#contentA").text)
#print(driver.find_element_by_id("contentA").text)