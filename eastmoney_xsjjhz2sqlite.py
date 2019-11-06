# -*- coding: utf-8 -*-
"""
功能：本程序从东方财富网提取股东户数的最新变化情况，保存sqlite
用法：每天运行
"""
import time
from selenium import webdriver
import sqlite3
import sys

###############################################################################
#长股票代码
###############################################################################
def lgpdm(dm):
    return dm[:6]+('.SH' if dm[0]=='6' else '.SZ')

###############################################################################
#短股票代码
###############################################################################
def sgpdm(dm):
    return dm[:6]

########################################################################
#建立数据库
########################################################################
def createDataBase():
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    cn = sqlite3.connect(dbfn)

    cn.execute('''CREATE TABLE IF NOT EXISTS GDHS
           (GPDM TEXT NOT NULL,
           RQ TEXT NOT NULL,
           GDHS INTEGER NOT NULL);''')
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GPDM_RQ_GDHS ON GDHS(GPDM,RQ);''')



def get_xsjjhz(pgn):

    data = []
    
    print("正在处理第%d页，请等待。" % pgn)
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    browser = webdriver.PhantomJS()
    browser.get("http://data.eastmoney.com/dxf/detail.aspx?market=0")
    time.sleep(5)

    try :    
        elem = browser.find_element_by_id("PageContgopage")
        elem.clear()
        #输入页面
        elem.send_keys(pgn)
        elem = browser.find_element_by_class_name("btn_link")     
        #点击Go
        elem.click()
        time.sleep(5)
        #定位到表体
        tbody = browser.find_elements_by_tag_name("tbody")
        #表体行数
        tblrows = tbody[0].find_elements_by_tag_name('tr')
    except :
        dbcn.close()
        return False

    #遍历行
    for j in range(len(tblrows)):
        dm = None
        jjrq = None
        bcjj = None
        qltzb = None
        try :    
            tblcols = tblrows[j].find_elements_by_tag_name('td')
            dm = tblcols[1].text
            jjrq = tblcols[4].text
            bcjj =  tblcols[6].text
            qltzb = tblcols[8].text
        except :
            dbcn.close()
            return False

        dm = lgpdm(dm)       
        rowdat = [dm,jjrq,bcjj,qltzb]
        data.append(rowdat)

    browser.quit()
    
    if len(data)>0:
        dbcn.executemany('INSERT OR REPLACE INTO XSJJ_DFCF (GPDM,JJRQ,JJSL,QLTBL) VALUES (?,?,?,?)', data)
        
    dbcn.commit()
    dbcn.close()

    return True

def getdrive():
    return sys.argv[0][:2]

def main():
#    createDataBase()
    j=1
    while j<=128:
        if get_xsjjhz(j):
            j+=1


if __name__ == "__main__": 
    print('%s Running' % sys.argv[0])
    main()