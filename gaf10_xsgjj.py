# -*- coding: utf-8 -*-
"""
Created on Tue Feb  7 22:25:24 2017
限售股解禁
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

    


def lngbbd(gpdm): 

    data = []

    data.append(['股票代码','变更日期','总股本(万股)','流通A股(万股)','实际流通A股(万股)','变更原因'])

    browser = webdriver.Firefox()
    browser.maximize_window()
    url  = "http://web-f10.gaotime.com/stock/"+gpdm+"/gbjg/lngbbd.html"

    browser.get(url)
    
    div = browser.find_element_by_class_name("content_02")
    tbl = div.find_element_by_id("TableHover")
    
    tblrows = tbl.find_elements_by_tag_name('tr')
    rows=len(tblrows)
#    print(rows)
    cols = len(tblrows[0].find_elements_by_tag_name('th'))
#    print(cols)
    for i in range(1,rows-1):
        tblcols = tblrows[i].find_elements_by_tag_name('td')
        rowdat = [gpdm]
        for j in range(cols):
            rowdat.append(str2float(tblcols[j].text))
        
        data.append(rowdat)
 
    return pd.DataFrame(data[1:],columns=data[0])       


def xsjj(gpdm):
    browser = webdriver.Firefox()
    browser.maximize_window()
    gpdm = "002322"
    data = []
    data.append(['股票代码','股东名称','流通日期','新增可售股数(万股)','限售类型','限售条件说明'])
    url  = "http://web-f10.gaotime.com/stock/"+gpdm+"/gbjg/xsjj.html"
    browser.get(url)
    
    div = browser.find_element_by_class_name("content_02")
    tbl = div.find_element_by_id("TableHover")
    
    tblrows = tbl.find_elements_by_tag_name('tr')
    rows=len(tblrows)
#    print(rows)
#    tblcols = tblrows[0].find_elements_by_tag_name('th')
#    cols=len(tblcols)
#    print(cols)
#    rowdat = []
#
#    for i in range(cols):
#        rowdat.append(tblcols[i].text)
#    
#    data.append(rowdat)        
    gdmc=""
    
    for i in range(1,rows-1):
        rowdat = [gpdm]
        tblcols = tblrows[i].find_elements_by_tag_name('td')
        if len(tblcols)==5 :
            gdmc=tblcols[0].text
        else:
            rowdat.append(gdmc)
        
        for j in range(len(tblcols)):
            rowdat.append(str2float(tblcols[j].text))
        
        data.append(rowdat)

    return pd.DataFrame(data[1:],columns=data[0])       

    
if __name__ == "__main__":  
    gpdm = "002322"
    df1 = lngbbd(gpdm)
    df2 = xsjj(gpdm)
    with pd.ExcelWriter('d:\\hyb\\限售股解禁时间表'+gpdm+'.xlsx',engine='xlsxwriter') as writer:
        df1.to_excel(writer, sheet_name='历年股本变动',index=False)
        df2.to_excel(writer, sheet_name='限售股解禁',index=False)    
    writer.save()
