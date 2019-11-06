# -*- coding: utf-8 -*-
"""
功能：本程序从同花顺网提取大宗交易数据生成提示信息，保存sqlite
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

def dstx_dzjy():
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
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
    ts1='大宗交易'
    while True:
        pages=browser.find_element_by_class_name('m-page')
        cur=eval(pages.find_element_by_class_name('cur').text)
        dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
        html = pq(dbtbl)

        print("正在处理第%d/%d页，请等待。" % (cur,pgs))

        data = []
        rows=html('tr')
        
        mrf0=None
        zyl0=None
        zje=0
        for i in range(1,len(rows)):
            
            rowdat=[]    
            row=pq(rows('tr').eq(i))
            
            rq=row('td').eq(1).text()
            dm=row('td').eq(2).text()            
            cjj=float(row('td').eq(5).text())            
            cjl=float(row('td').eq(6).text())            
            cje=round(cjj*cjl,2)
            zyl=row('td').eq(7).text()
            mrf=row('td').eq(8).text()       

            if zyl==zyl0 and mrf==mrf0 :
                zje=zje+cje
                continue
            else:
                zje=cje
                
            if float(zyl.replace('%',''))<0:
                ts1='大宗交易折价,折溢率%s%%' % zyl
            else:
                ts1='大宗交易溢价,折溢率%s%%' % zyl

                
            ts2='买方：%s,折溢率%s,成交额%d万元' % (mrf,zyl,zje)
            dm = lgpdm(dm)       

            rowdat = [dm,rq,ts1,ts2,'0']
            data.append(rowdat)
            mrf0=mrf
            zyl0=zyl
            zje=0
    
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data)
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

    return    

def dstx_yjyg():
    
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/financial/yjyg/'
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
            
            rq=row('td').eq(7).text()
            dm=row('td').eq(1).text()            
            ts1=row('td').eq(3).text()            
            ts2=row('td').eq(4).text()

            dm = lgpdm(dm)       

            rowdat = [dm,rq,ts1,ts2,'0']
            data.append(rowdat)
    
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data)
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

    return


def dstx_ggsd():    
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

    return

def dstx_yjkb():
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/financial/yjkb/'
    browser.get(url)
    time.sleep(5)
    try:
        elem = browser.find_element_by_class_name("page_info")
        pgs=int(1/eval(elem.text))
    except:
        pgs=1

    while True:
        try:
            pages=browser.find_element_by_class_name('m-page')
            cur=eval(pages.find_element_by_class_name('cur').text)
        except:
            cur=1
            
        dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
        html = pq(dbtbl)

        print("正在处理第%d/%d页，请等待。" % (cur,pgs))

        data = []
        rows=html('tr')
 
        for i in range(2,len(rows)):

            rowdat=[]    
            row=pq(rows('tr').eq(i))
            
            rq=row('td').eq(3).text()
            dm=row('td').eq(1).text()    
            yysr=row('td').eq(4).text()
            yysr_g = row('td').eq(6).text()
            jlr=row('td').eq(8).text()
            jlr_g = row('td').eq(10).text()
            eps = row('td').eq(12).text()
            roe = row('td').eq(12).text()
            
            if float(jlr_g)>0 :
                ts1='业绩快报:净利润增长,净利润同比%s%%' % jlr_g
            else :
                ts1='业绩快报:净利润减少,净利润同比%s%%' % jlr_g

            ts2='业绩快报,营业收入%s,同比%s%%,净利润%s,同比%s%%,EPS%s元,ROE%s%%' % (yysr,yysr_g,jlr,jlr_g,eps,roe)
                
            dm = lgpdm(dm)       

            rowdat = [dm,rq,ts1,ts2,'0']
            data.append(rowdat)
    
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data)
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
    return

def dstx_yjgg():
    
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/financial/yjgg/'

    browser.get(url)
    time.sleep(5)
    try:
        elem = browser.find_element_by_class_name("page_info")
        pgs=int(1/eval(elem.text))
    except:
        pgs=1

    while True:
        try:
            pages=browser.find_element_by_class_name('m-page')
            cur=eval(pages.find_element_by_class_name('cur').text)
        except:
            cur=1
            
        dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
        html = pq(dbtbl)

        print("正在处理第%d/%d页，请等待。" % (cur,pgs))

        data = []
        rows=html('tr')
 
        for i in range(2,len(rows)):

            rowdat=[]    
            row=pq(rows('tr').eq(i))
            
            rq=row('td').eq(3).text()
            dm=row('td').eq(1).text()  
            
            yysr=row('td').eq(4).text()
            yysr_g = row('td').eq(5).text()
            jlr=row('td').eq(7).text()
            jlr_g = row('td').eq(8).text()
            eps = row('td').eq(10).text()
            roe = row('td').eq(12).text()
            
            if float(jlr_g)>0 :
                ts1='业绩公告:净利润增长,净利润同比%s%%' % jlr_g
            else :
                ts1='业绩公告:净利润减少,净利润同比%s%%' % jlr_g

            ts2='业绩公告,营业收入%s,同比%s%%,净利润%s,同比%s%%,EPS%s元,ROE%s%%' % (yysr,yysr_g,jlr,jlr_g,eps,roe)
                
            dm = lgpdm(dm)       

            rowdat = [dm,rq,ts1,ts2,'0']
            data.append(rowdat)
    
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data)
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
    
    return

if __name__ == "__main__": 
    print('%s Running' % sys.argv[0])

    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

#    dstx_dzjy()    #大宗交易  
#    dstx_yjyg()    #业绩预告
#    dstx_yjkb()    #业绩快报
#    dstx_yjgg()    #业绩公告
    
#    dstx_ggsd()    #公告速递
     
    dbfn=getdrive()+'\\hyb\\STOCKDSTX.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.10jqka.com.cn/financial/yjgg/'

    browser.get(url)
    time.sleep(5)
    try:
        elem = browser.find_element_by_class_name("page_info")
        pgs=int(1/eval(elem.text))
    except:
        pgs=1

    while True:
        try:
            pages=browser.find_element_by_class_name('m-page')
            cur=eval(pages.find_element_by_class_name('cur').text)
        except:
            cur=1
            
        dbtbl = browser.find_element_by_class_name("page-table").get_attribute("innerHTML")
        html = pq(dbtbl)

        print("正在处理第%d/%d页，请等待。" % (cur,pgs))

        data = []
        rows=html('tr')
 
        for i in range(2,len(rows)):

            rowdat=[]    
            row=pq(rows('tr').eq(i))
            
            rq=row('td').eq(3).text()
            dm=row('td').eq(1).text()  
            
            yysr=row('td').eq(4).text()
            yysr_g = row('td').eq(5).text()
            jlr=row('td').eq(7).text()
            jlr_g = row('td').eq(8).text()
            eps = row('td').eq(10).text()
            roe = row('td').eq(12).text()
            
            if float(jlr_g)>0 :
                ts1='业绩公告:净利润增长,净利润同比%s%%' % jlr_g
            else :
                ts1='业绩公告:净利润减少,净利润同比%s%%' % jlr_g

            ts2='业绩公告,营业收入%s,同比%s%%,净利润%s,同比%s%%,EPS%s元,ROE%s%%' % (yysr,yysr_g,jlr,jlr_g,eps,roe)
                
            dm = lgpdm(dm)       

            rowdat = [dm,rq,ts1,ts2,'0']
            data.append(rowdat)
    
        if len(data)>0:
            dbcn.executemany('INSERT OR REPLACE INTO THS (GPDM,RQ,TS1,TS2,TSLX) VALUES (?,?,?,?,?)', data)
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
