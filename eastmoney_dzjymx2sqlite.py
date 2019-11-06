# -*- coding: utf-8 -*-
"""
功能：本程序从东方财富网提取最近30个交易日大宗交易数据，保存sqlite
用法：每天运行
"""
import time
import datetime
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
    
    cn.execute('DROP TABLE IF EXISTS [DZJY_DFCF];')
   
    cn.execute('''CREATE TABLE [DZJY_DFCF](
                  [GPDM] TEXT NOT NULL, 
                  [RQ] TEXT NOT NULL,
                  [ZDF] REAL NOT NULL,
                  [SPJ] REAL NOT NULL, 
                  [CJJ] REAL NOT NULL,
                  [CJJZDF] REAL NOT NULL, 
                  [ZYL] REAL NOT NULL, 
                  [CJL] REAL NOT NULL, 
                  [CJE] REAL NOT NULL, 
                  [LTZB] REAL NOT NULL, 
                  [MRF] TEXT NOT NULL, 
                  [MCF] TEXT NOT NULL, 
                  [D1ZD] REAL, 
                  [D5ZD] REAL,
                  [D10ZD] REAL, 
                  [D20ZD] REAL);
                ''')

    cn.commit()
    
    
#'''
#有全部信息重复的情况，不能建唯一索引
#http://data.eastmoney.com/dzjy/dzjy_mrmxa.aspx?TimeSpanType=30
#
#
#'''


def getdrive():
    return sys.argv[0][:2]




def get_dzjymx():
    
    print('%s Running' % sys.argv[0])
    createDataBase()

    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    data = []
    
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 


#    fireFoxOptions = webdriver.FirefoxOptions()
#    fireFoxOptions.set_headless()
#    browser = webdriver.Firefox(firefox_options=fireFoxOptions)

    
    url = 'http://data.eastmoney.com/dzjy/dzjy_mrmxa.aspx?TimeSpanType=30'
    browser.get(url)
    time.sleep(5)

    try:
        elem = browser.find_element_by_id("PageCont")
        pgs=int(elem.find_element_by_xpath("//a[@title='转到最后一页']").text)
    except:
        browser.quit()
        exit()

    pgn=1
    while pgn<=pgs:

        print("正在处理第%d/%d页，请等待。" % (pgn,pgs))
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


        #遍历行
        data = []
        sc=True     #本页处理成功
        for j in range(len(tblrows)):
            try:
            
                tblcols = tblrows[j].find_elements_by_tag_name('td')
                
                rq = tblcols[1].get_property("title")         
                dm = tblcols[2].text
                zdf = float(tblcols[5].text)
                spj = float(tblcols[6].text)
                cjj = float(tblcols[7].text)
                zyl = float(tblcols[8].text)
                cjl = float(tblcols[9].text)
                cje = float(tblcols[10].text)
                ltzb = float(tblcols[11].text.replace('%',''))
                mryyb = tblcols[12].text
                mcyyb = tblcols[13].text
                d1zd = tblcols[14].text
                d1zd = float(d1zd) if d1zd!='-' else None                
                d5zd = tblcols[15].text
                d5zd = float(d5zd) if d5zd!='-' else None                
                
                

                if dm[0] in ('0','3','6'):
        
                    dm = lgpdm(dm)       
                    data.append([lgpdm(dm),rq,zdf,spj,cjj,zyl,cjl,cje,ltzb,mryyb,mcyyb,d1zd,d5zd])
            except:
                sc=False    #本页处理不成功
                break
        
        if len(data)>0 and sc:
            dbcn.executemany('''INSERT OR REPLACE INTO DZJY_DFCF (GPDM,RQ,ZDF,SPJ,CJJ,ZYL,CJL,CJE,LTZB,MRF,MCF,D1ZD,D5ZD) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''', data)
            dbcn.commit()
            pgn+=1
        else:
            browser.get(url)
            time.sleep(5)
            

    dbcn.close()
    browser.quit()

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)

if __name__ == "__main__": 
#    get_dzjymx()

    createDataBase()
    gpdm='002449'
#    get_dzjymx_gg(gpdm)


    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    data = []
    
    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

    chrome_options = webdriver.ChromeOptions() 
    chrome_options.add_argument("--headless") 
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_options) 


#    fireFoxOptions = webdriver.FirefoxOptions()
#    fireFoxOptions.set_headless()
#    browser = webdriver.Firefox(firefox_options=fireFoxOptions)

    
    url = 'http://data.eastmoney.com/dzjy/detail/%s.html' % gpdm
    browser.get(url)
    time.sleep(5)

    try:
        #Python selenium —— 父子、兄弟、相邻节点定位方式详解
        #https://blog.csdn.net/huilan_same/article/details/52541680
        elem = browser.find_element_by_xpath("//a[text()='下一页']//preceding-sibling::a[1]")
        pgs=int(elem.text)
    except:
        pgs=1


    pgn=1
    while pgn<=pgs:

        print("正在处理第%d/%d页，请等待。" % (pgn,pgs))

        if pgn>1:
            elem = browser.find_element_by_id("PageContgopage")
            elem.clear()
            #输入页面
            elem.send_keys(pgn)
            elem = browser.find_element_by_class_name("btn_link")     
            #点击Go
            elem.click()
            time.sleep(5)

        tab1 = browser.find_element_by_id("dt_1")
        #定位到表体
        tbody = tab1.find_element_by_tag_name("tbody")
        #表体行数
        tblrows = tbody.find_elements_by_tag_name('tr')


        #遍历行
        data = []
        sc=True     #本页处理成功
        for j in range(len(tblrows)):
            try:
            
                tblcols = tblrows[j].find_elements_by_tag_name('td')
                
                rq = tblcols[1].text         

                zdf = float(tblcols[2].text)
                spj = float(tblcols[3].text)
                cjj = float(tblcols[4].text)
                cjjzdf = round((cjj/(spj/(1+zdf/100))-1)*100,2)
                zyl = float(tblcols[5].text)
                cjl = float(tblcols[6].text)
                cje = float(tblcols[7].text)
                ltzb = float(tblcols[8].text.replace('%',''))
                mryyb = tblcols[9].text
                mcyyb = tblcols[10].text
                d1zd = tblcols[11].text
                d1zd = float(d1zd) if d1zd!='-' else None                
                d5zd = tblcols[12].text
                d5zd = float(d5zd) if d5zd!='-' else None                
                d10zd = tblcols[13].text
                d10zd = float(d10zd) if d10zd!='-' else None                
                d20zd = tblcols[14].text
                d20zd = float(d20zd) if d20zd!='-' else None                
                

                if gpdm[0] in ('0','3','6'):
        
                    dm = lgpdm(gpdm)       
                    data.append([lgpdm(dm),rq,zdf,spj,cjj,cjjzdf,zyl,cjl,cje,ltzb,mryyb,mcyyb,d1zd,d5zd,d10zd,d20zd])
            except:
                sc=False    #本页处理不成功
                break
        
        if len(data)>0 and sc:
            dbcn.executemany('''INSERT OR REPLACE INTO DZJY_DFCF (GPDM,RQ,ZDF,SPJ,CJJ,CJJZDF,ZYL,CJL,CJE,LTZB,MRF,MCF,D1ZD,D5ZD,D10ZD,D20ZD) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', data)
            dbcn.commit()
            pgn+=1
        else:
            browser.get(url)
            time.sleep(5)
            

    dbcn.close()
    browser.quit()

    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)
        