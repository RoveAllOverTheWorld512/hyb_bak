# -*- coding: utf-8 -*-
"""
本程序将从东方财富网抓取股东户数信息存入sqlite数据库
"""

from selenium import webdriver
import pandas as pd
import sqlite3
import time
import datetime

def getgdhs(gpdm,browser):
    
    dm=gpdm[:6]
    data = []
    url = "http://data.eastmoney.com/gdhs/detail/"+dm+".html"
    
    browser.get(url)
    time.sleep(5)
    
    pgnv = browser.find_elements_by_id("PageCont")
    pgs=pgnv[0].find_elements_by_tag_name("a")
    if len(pgs)==0 :
        pg=1
    else:
        pg=int(pgs[len(pgs)-3].text)
        
    tbl = browser.find_elements_by_id("dt_1")
    tbody = tbl[0].find_elements_by_tag_name("tbody")
    tblrows = tbody[0].find_elements_by_tag_name('tr')
    #没有相关数据时也有1行
    if len(tblrows)>1:  
        for j in range(len(tblrows)):
            tblcols = tblrows[j].find_elements_by_tag_name('td')
            rq=tblcols[0].text
            rq=rq.replace('/','-')
            hs=tblcols[2].text
            rowdat = [gpdm,rq,hs]
            data.append(rowdat)
            
        if pg>1 :
            for k in range(2,pg+1):
                elem = browser.find_element_by_id("PageContgopage")
                elem.clear()
                elem.send_keys(k)
                elem = browser.find_element_by_class_name("btn_link")        
                elem.click()
                time.sleep(2)
                tbl = browser.find_elements_by_id("dt_1")
                tbody = tbl[0].find_elements_by_tag_name("tbody")
                tblrows = tbody[0].find_elements_by_tag_name('tr')
                   
                for j in range(len(tblrows)):
                    tblcols = tblrows[j].find_elements_by_tag_name('td')
                    rq=tblcols[0].text
                    rq=rq.replace('/','-')
                    hs=tblcols[2].text
                    rowdat = [gpdm,rq,hs]
                    data.append(rowdat)
                    
    return data


###############################################################################
#从通达信系统读取股票代码表
###############################################################################
def get_gpdm():
    datacode = []
    for sc in ('h','z'):
        fn = r'C:\new_hxzq_hc\T0002\hq_cache\s'+sc+'m.tnf'
        f = open(fn,'rb')
        f.seek(50)
        ss = f.read(314)
        while len(ss)>0:
            gpdm=ss[0:6].decode('GBK')
            gpmc=ss[23:31].strip(b'\x00').decode('GBK').replace(' ','').replace('*','')
            gppy=ss[285:289].strip(b'\x00').decode('GBK')
            #剔除非A股代码
            if (sc=="h" and gpdm[0]=='6') :
                gpdm=gpdm+'.SH'
                datacode.append([gpdm,gpmc,gppy])
            if (sc=='z' and (gpdm[0]=='0' or gpdm[0:2]=='30')) :
                gpdm=gpdm+'.SZ'
                datacode.append([gpdm,gpmc,gppy])
            ss = f.read(314)
        f.close()
    gpdmb=pd.DataFrame(datacode,columns=['gpdm','gpmc','gppy'])
    return gpdmb

if __name__ == "__main__":  

    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    gpdmb=get_gpdm()
    dbcn = sqlite3.connect('d:\\hyb\\STOCKDATA.db')
    curs = dbcn.cursor()
#   获取最后一行数据
    rec=curs.execute('''select gpdm from gdhs where rowid in (select max(rowid) from gdhs);''')
    row=rec.fetchone()
    dm=row[0]   

#    browser = webdriver.Firefox()
    browser = webdriver.PhantomJS()
#    browser.maximize_window()
#    dm='000717.SZ'

    j=list(gpdmb['gpdm']).index(dm)
    k=len(gpdmb)       #最大值为自选股总数len(gpdmb)        

    for i in range(j+1,k):
        gpdm = gpdmb.loc[i,'gpdm']
        gpmc = gpdmb.loc[i,'gpmc']
        print("共有%d只股票，正在处理第%d只：%s%s，请等待…………" % (len(gpdmb),i+1,gpdm,gpmc)) 

        data=getgdhs(gpdm,browser)
        dbcn.executemany('INSERT OR IGNORE INTO GDHS (GPDM,RQ,GDHS) VALUES (?,?,?)', data)
        dbcn.commit()
    
    browser.quit()
    dbcn.close()

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)


    
        
