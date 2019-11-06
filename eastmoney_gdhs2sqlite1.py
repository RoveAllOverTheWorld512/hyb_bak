# -*- coding: utf-8 -*-
"""
功能：本程序从东方财富网提取股东户数的最新变化情况，保存sqlite
用法：每天运行
"""
from selenium import webdriver
import sqlite3
import sys
import time
import datetime

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



def getgdhs(browser,pgn,dbcn):

    data = []
    
    print("正在处理第%d页，请等待。" % pgn)
#    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
#    dbcn = sqlite3.connect(dbfn)
#
#    browser = webdriver.PhantomJS()
    browser.get("http://data.eastmoney.com/gdhs/")
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

        return False

    #遍历行
    for j in range(len(tblrows)):
        dm = None
        rq = None
        hs = None
        try :    
            tblcols = tblrows[j].find_elements_by_tag_name('td')
            dm = tblcols[0].text
            hs =  tblcols[5].text
            rq = tblcols[10].text
        except :
            return False

        if dm != None and rq != None and hs != None :
            dm = lgpdm(dm)       
            rq = rq.replace('/','-')
            rowdat = [dm,rq,hs]
            data.append(rowdat)

    
    if len(data)>0:
        dbcn.executemany('INSERT OR REPLACE INTO GDHS (GPDM,RQ,GDHS) VALUES (?,?,?)', data)
        
    dbcn.commit()

    return True


def getdrive():
    return sys.argv[0][:2]

def main():
    createDataBase()
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

#    browser = webdriver.PhantomJS()
    fireFoxOptions = webdriver.FirefoxOptions()
    fireFoxOptions.set_headless()
    browser = webdriver.Firefox(firefox_options=fireFoxOptions)

    j=1
    while j<=15:
        if getgdhs(browser,j,dbcn):
            j+=1

    dbcn.close()
    browser.quit()
        
if __name__ == "__main__": 
    print('%s Running' % sys.argv[0])
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    main()
    
    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)
    