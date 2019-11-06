# -*- coding: utf-8 -*-
"""
功能：本程序从同花顺网提取大宗交易数据，保存sqlite
用法：每天运行
"""
import time
import datetime
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

    cn.execute('''CREATE TABLE IF NOT EXISTS [DZJY_THS](
                  [GPDM] TEXT NOT NULL, 
                  [RQ] TEXT NOT NULL, 
                  [CJJ] REAL NOT NULL, 
                  [CJL] REAL NOT NULL, 
                  [CJE] REAL NOT NULL, 
                  [ZYL] REAL NOT NULL, 
                  [MRF] TEXT NOT NULL, 
                  [MCF] TEXT NOT NULL);
                ''')

'''
CREATE TABLE [THS](
  [GPDM] TEXT NOT NULL, 
  [RQ] TEXT NOT NULL, 
  [TS1] TEXT, 
  [TS2] TEXT, 
  [TSLX] TEXT NOT NULL);

CREATE UNIQUE INDEX [GPDM_RQ_TS1_TS2_THS]
ON [THS](
  [GPDM], 
  [RQ], 
  [TS1], 
  [TS2]);
'''
    
def getdrive():
    return sys.argv[0][:2]


if __name__ == "__main__": 
    print('%s Running' % sys.argv[0])

    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/market/dzjy/'
    browser.get(url)
    time.sleep(5)

    elem = browser.find_element_by_class_name("page_info")
    pgs=int(1/eval(elem.text))
    
    while True:
        pages=browser.find_element_by_class_name('m-page')
        cur=eval(pages.find_element_by_class_name('cur').text)
        dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
        html = pq(dbtbl)

        print("正在处理第%d/%d页，请等待。" % (cur,pgs))

        data = []
        rows=html('tr')
 
        for i in range(1,len(rows)):

            rowdat=[]    
            row=pq(rows('tr').eq(i))
            
            rq=row('td').eq(1).text()
            dm=row('td').eq(2).text()            
            cjj=float(row('td').eq(5).text())            
            cjl=float(row('td').eq(6).text())            
            cje=round(cjj*cjl,2)
            zyl=float(row('td').eq(7).text().replace('%',''))
            mrf=row('td').eq(8).text()            
            mcf=row('td').eq(9).text()            

            dm = lgpdm(dm)       

            rowdat = [dm,rq,cjj,cjl,cje,zyl,mrf,mcf]
            data.append(rowdat)
    
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO DZJY_THS (GPDM,RQ,CJJ,CJL,CJE,ZYL,MRF,MCF) VALUES (?,?,?,?,?,?,?,?)', data)
            dbcn.commit()

        if cur<pgs:
            elem = browser.find_element_by_xpath("//a[text()='下一页']")
            elem.click()
            time.sleep(5)
        else:
            break
        
    browser.quit()
    dbcn.commit()
    dbcn.close()

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)
