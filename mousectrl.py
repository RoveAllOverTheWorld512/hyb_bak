# -*- coding: utf-8 -*-
"""
Created on Tue May  1 16:22:46 2018

@author: lenovo
"""
 
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from time import sleep
 
 
driver = webdriver.Firefox()
driver.implicitly_wait(10)
driver.maximize_window()
driver.get('http://sahitest.com/demo/clicks.htm')
 
click_btn = driver.find_element_by_xpath('//input[@value="click me"]') # 单击按钮
doubleclick_btn = driver.find_element_by_xpath('//input[@value="dbl click me"]') # 双击按钮
rightclick_btn = driver.find_element_by_xpath('//input[@value="right click me"]') # 右键单击按钮
 
 
ActionChains(driver).click(click_btn).double_click(doubleclick_btn).context_click(rightclick_btn).perform() # 链式用法
 
print(driver.find_element_by_name('t2').get_attribute('value'))
 
sleep(2)
#driver.quit()