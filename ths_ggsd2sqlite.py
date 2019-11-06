# -*- coding: utf-8 -*-
"""
功能：本程序从同花顺网公告速递，保存sqlite
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


    
def getdrive():
    return sys.argv[0][:2]


if __name__ == "__main__": 
    print('%s Running' % sys.argv[0])

    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/market/ggsd/'
    browser.get(url)
    time.sleep(5)

    bbs = browser.find_elements_by_class_name('J-board-item')
    for n in range(len(bbs)):
        bbs[n].click()
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
                tds=row('td')
                if len(tds)==2:
                    ggbt=tds.eq(0).text()
                    gglx=tds.eq(1).text()
                else:
                    rq=tds.eq(1).text()
                    dm=tds.eq(2).text() 
                    dm = lgpdm(dm)    
                    ggbt=None
                    gglx=None
    
    #            if gglx in ('持股变动公告','股权激励','股票质押公告','资产购买公告','增发事项公告',):
                if gglx!=None and gglx!='--':
                    rowdat = [dm,rq,gglx,ggbt,'0']
                    data.append(rowdat)
                    
            #由于这个网页存在嵌套表，pyQuery分析时行数会被递归多次计算，出现数据重复，下面是对数据去重        
            data1=[]
            for dt in data:
                if dt not in data1:
                    data1.append(dt)                
                    
            if len(data1)>0:
                dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data1)
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
