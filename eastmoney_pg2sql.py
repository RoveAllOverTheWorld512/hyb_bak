# -*- coding: utf-8 -*-
"""
本程序将从东方财富网抓取股东户数信息存入sqlite数据库
"""

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import sqlite3
import time
import datetime

def createDataBase():
    cn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    '''
    GPDM股票代码
    RQ股权登记日，除权除息日一般为该股下一个交易日
    FH每股分红
    SZG每股送转股
    PGBL每股配股比例
    PGJ每股配股价
    GPDM与RQ构成为唯一索引
    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS PG
           (GPDM TEXT NOT NULL,
           RQ TEXT NOT NULL,
           FH REAL NOT NULL DEFAULT 0.00,
           SZG REAL NOT NULL DEFAULT 0.00,
           PGBL REAL NOT NULL DEFAULT 0.00,
           PGJ REAL NOT NULL DEFAULT 0.00);''')
    
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ_PG ON PG(GPDM,RQ);''')

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def str2float(num):
    try:
        return float(num)
    except ValueError:
        return num


###############################################################################
#从东方财富网抓取配股数据
###############################################################################
def getdata(pgn):
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    url = "http://data.eastmoney.com/zrz/pg.html"
    browser = webdriver.PhantomJS()
    browser.get(url)
    time.sleep(5)

    print('正在处理第%d页，请等待。' % pgn)
    try:
        elem = browser.find_element_by_id("PageContgopage")
        elem.clear()
        elem.send_keys(pgn)
        elem = browser.find_element_by_class_name("btn_link")        
        elem.click()
        time.sleep(5)
        tbl = browser.find_elements_by_id("dt_1")
        tbody = tbl[0].find_elements_by_tag_name("tbody")
        tblrows = tbody[0].find_elements_by_tag_name('tr')
    except NoSuchElementException as e:
        print(e.msg)
        return False
        
    data = []
       
    for j in range(len(tblrows)):
        pgj=0
        pgbl=0
        try:
            tblcols = tblrows[j].find_elements_by_tag_name('td')
            dm=tblcols[0].text
            dm=dm+('.SH' if dm[0]=='6' else '.SZ')
            pgj=str2float(tblcols[6].text)
            pgqgb=str2float(tblcols[8].text)
            pghgb=str2float(tblcols[9].text)
            pgbl=pghgb/pgqgb-1
            rq = tblcols[10].find_elements_by_tag_name('span')[0].get_property("title")        
        except NoSuchElementException as e:
            print(e.msg)
            print(j)
            return False

        if rq!='-' and pgj>0 and pgbl>0 :
            rowdat = [dm,rq,pgj,pgbl]
            data.append(rowdat)
        
    dbcn.executemany('INSERT OR REPLACE INTO PG (GPDM,RQ,PGJ,PGBL) VALUES (?,?,?,?)', data)
    dbcn.commit()
    
    browser.quit()
    dbcn.close()
    return True

if __name__ == "__main__":  
    
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    pgs=2
    j=1
    while j<pgs+1:
        if getdata(j):
            j+=1

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

    

    
        
