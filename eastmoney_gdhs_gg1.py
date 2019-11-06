# -*- coding: utf-8 -*-
"""
Created on Tue Feb  7 22:25:24 2017

@author: lenovo
"""
import time
from pyquery import PyQuery as pq
from selenium import webdriver
import pandas as pd

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def str2float(num):
    try:
        return float(num)
    except ValueError:
        return num

def tmp():
    
    #browser = webdriver.Ie()
    browser = webdriver.Firefox()
    #browser = webdriver.Chrome()
    #browser.implicitly_wait(30)
    #browser = webdriver.Firefox()
    
    #浏览器窗口最大化
    browser.maximize_window()
    #登录同花顺
    browser.get("http://data.eastmoney.com/gdhs/")
    
    time.sleep(3)
    
    pages = browser.find_element_by_id("PageCont").find_elements_by_tag_name('a')[6].text
    
    print("总页数：",pages)
    tbody = browser.find_elements_by_tag_name("tbody")
    tblrows = tbody[0].find_elements_by_tag_name('tr')
    print("总行数:",len(tblrows))
    data = []
    data.append(['股票代码','股票名称','本次户数','上次户数','增减户数','增加比例','本次截止日期','上次截止日期','公告日期'])
    
    #for i in range(pages-1) :
    browser.get("http://data.eastmoney.com/gdhs/")
    elem = browser.find_element_by_id("PageContgopage")
    elem.clear()
    elem.send_keys(i+1)
    elem.click()
    time.sleep(2)

    tbody = browser.find_elements_by_tag_name("tbody")
    tblrows = tbody[0].find_elements_by_tag_name('tr')
    #table = browser.find_element_by_id('dt_1')
    #table的总行数，包含标题
    #table_rows = table.find_elements_by_tag_name('tr')
    #print( "总行数:",len(table_rows))

    #tabler的总列数
    '''
    #在table中找到第一个tr,之后在其下找到所有的th,即是tabler的总列数
    '''
    table_cols = table_rows[0].find_elements_by_tag_name('th')
    print( "总列数:",len(table_cols))
    table_cols1 = table_rows[10].find_elements_by_tag_name('td')
    abc=table_cols1[16].find_elements_by_tag_name('span')
    print(abc[0].get_property("title"))
    print(table_cols1[16].find_elements_by_tag_name('span')[0].get_property("title"))

    

if __name__ == "__main__":  
    
    gpdm = "002536"
    data = []
    data.append(['股票代码','股票简称','截止日期','区间涨幅(%)','本次户数','上次户数','增减户数','增加比例','户均市值(万)','户均持股(万)',
                 '总市值(亿)','总股本(亿)','股本变动','股本变动原因','公告日期'])
   
    browser = webdriver.Firefox()
    browser.maximize_window()
    url = "http://data.eastmoney.com/gdhs/detail/"+gpdm+".html"
    
    browser.get(url)
    
    gpmc=browser.find_element_by_class_name("tit")
    gpmc=gpmc.text[0:gpmc.text.find(gpdm)-1]
    
    tbl = browser.find_elements_by_id("dt_1")
    tbody = tbl[0].find_elements_by_tag_name("tbody")
    tblrows = tbody[0].find_elements_by_tag_name('tr')
       
    for j in range(len(tblrows)):
        rowdat = [gpdm,gpmc]
        tblcols = tblrows[j].find_elements_by_tag_name('td')
        for i in range(len(tblcols)):
            coldat = str2float(tblcols[i].text)
            rowdat.append(coldat)
    
        data.append(rowdat)
 
    gdhs = pd.DataFrame(data[1:],columns=data[0])
    gdhs.to_excel('d:\\hyb\\'+gpdm+gpmc+'股东户数.xlsx', sheet_name='股东户数',index=False)
        
