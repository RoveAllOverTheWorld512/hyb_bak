# -*- coding: utf-8 -*-
"""
功能：本程序从东方财富网提取大股东进出的最新变化情况，保存sqlite
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

    '''
    gddm股东代码,gdmc股东名称,gdlx股东类型,gdpm股东排名,gpdm股票代码,gpmc股票名称,
    bgq报告期,cgsl持股数量,ltzb持股占流通股比例,zjsl持股增减数量,cgbd持股变动,ggrq公告日期
    '''
    cn.execute('''CREATE TABLE IF NOT EXISTS GDFX
           (GDDM TEXT NOT NULL,
           GDMC TEXT,
           GDLX TEXT,
           GDPM TEXT,
           GPDM TEXT NOT NULL,
           GPMC TEXT,
           BGQ TEXT NOT NULL,
           CGSL REAL,
           LTZB REAL,
           ZJSL REAL,
           CGBD TEXT,
           GGRQ TEXT NOT NULL);''')
    
    cn.execute('''CREATE UNIQUE INDEX IF NOT EXISTS GDFX_GPDM_GPDM_BGQ ON GDFX(GDDM,GPDM,BGQ);''')



def getgdfx(gddm,gdmc,pgn):

    data = []
    
    print("正在处理第%d页，请等待。" % pgn)
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    browser = webdriver.PhantomJS()
    browser.get("http://data.eastmoney.com/gdfx/ShareHolderDetail.aspx?hdCode=%s&hdName=%s" % (gddm,gdmc))
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
        try :
            tblcols = tblrows[j].find_elements_by_tag_name('td')
            gdlx = tblcols[2].text     #股东类型
            gdpm = tblcols[3].text     #股东排名
            gpdm = tblcols[4].text     #股票代码
            gpmc = tblcols[5].text     #股票名称
            bgq = tblcols[7].text      #报告期
            
            cgsl = tblcols[8].text     #持股数量
            cgsl = cgsl.replace(',','')
            
            ltzb = tblcols[9].text     #持股占流通股比例
            
            zjsl = tblcols[10].text    #增减数量
            zjsl = zjsl.replace(',','')
            zjsl = zjsl if zjsl !='-' else None
            
            cgbd = tblcols[12].text    #持股变动
            
            ggrq = tblcols[14].find_element_by_tag_name('span').get_property('title')    #公告日期
        except :
            dbcn.close()
            return False

        if gddm != None and gpdm != None and bgq != None :
            gpdm = lgpdm(gpdm)       
            rowdat = [gddm,gdmc,gdlx,gdpm,gpdm,gpmc,bgq,cgsl,ltzb,zjsl,cgbd,ggrq]
            data.append(rowdat)

    browser.quit()
    
    if len(data)>0:
        dbcn.executemany('''INSERT OR REPLACE INTO GDFX 
                         (GDDM,GDMC,GDLX,GDPM,GPDM,GPMC,BGQ,CGSL,LTZB,ZJSL,CGBD,GGRQ)
                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''', data)
        
    dbcn.commit()
    dbcn.close()

    return True

def getdrive():
    return sys.argv[0][:2]

def main():
    createDataBase()
#    gddm='80475097'
#    gdmc='中央汇金资产管理有限责任公司'
    gddm='80188285'
    gdmc='中国证券金融股份有限公司'
    j=1
    while j<=104:
        if getgdfx(gddm,gdmc,j):
            j+=1


if __name__ == "__main__": 
    print('%s Running' % sys.argv[0])
    main()
