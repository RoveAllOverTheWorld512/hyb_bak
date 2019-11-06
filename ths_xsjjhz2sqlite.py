# -*- coding: utf-8 -*-
"""
功能：本程序从同花顺网提取股东户数的最新变化情况，保存sqlite
用法：每天运行
"""
import time
from selenium import webdriver
import sqlite3
import sys
from pyquery import PyQuery as pq

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

    cn.execute('''CREATE TABLE IF NOT EXISTS [XSJJ_THS](
                  [GPDM] TEXT NOT NULL, 
                  [JJRQ] TEXT NOT NULL, 
                  [JJSL] TEXT NOT NULL, 
                  [JJSZ] TEXT,
                  [ZGBBL] TEXT);''')
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS [GPDM_JJRQ_XSJJ_THS]
                ON [XSJJ_THS]([GPDM], [JJRQ]);''')



def getdrive():
    return sys.argv[0][:2]


if __name__ == "__main__": 
    print('%s Running' % sys.argv[0])
    data = []
    
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    browser = webdriver.Firefox()
#    browser = webdriver.PhantomJS() #用本句替代上一句不成功，总停留在第一页，不能执行点击下一页操作
    browser.get("http://data.10jqka.com.cn/market/xsjj/")
    time.sleep(5)

    elem = browser.find_element_by_class_name("page_info")
    pgs=int(1/eval(elem.text))
    

    while True:
        
        '''
        Selenium在定位的class含有空格的复合类的解决办法
        https://blog.csdn.net/cyjs1988/article/details/75006167
        '''
        pages=browser.find_element_by_class_name('m-page')
        cur=eval(pages.find_element_by_class_name('cur').text)
        print(cur)
        dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
        html = pq(dbtbl)

        rows=html('tr')
 
    
        for i in range(1,len(rows)):

            row=rows.eq(i).text().split(' ')
            dm=row[1]
            jjrq=row[3]
            bcjj=row[4]
            jjsz=row[6]
            zgbzb=row[7]


            dm = lgpdm(dm)       
            rowdat = [dm,jjrq,bcjj,jjsz,zgbzb]
            data.append(rowdat)

        if cur<pgs:
            elem = browser.find_element_by_xpath("//a[text()='下一页']")
            elem.click()
            time.sleep(5)
        else:
            break
    
    if len(data)>0:
        dbcn.executemany('INSERT OR REPLACE INTO XSJJ_THS (GPDM,JJRQ,JJSL,JJSZ,ZGBBL) VALUES (?,?,?,?,?)', data)
        
    browser.quit()
    dbcn.commit()
    dbcn.close()
