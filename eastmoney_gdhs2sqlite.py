# -*- coding: utf-8 -*-
"""
功能：本程序从东方财富网提取股东户数的最新变化情况，保存sqlite
用法：每天运行
"""
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sqlite3
import sys
import os
import datetime
from pyquery import PyQuery as pq
from configobj import ConfigObj


########################################################################
#初始化本程序配置文件
########################################################################
def iniconfig():
    inifile = os.path.splitext(sys.argv[0])[0]+'.ini'  #设置缺省配置文件
    return ConfigObj(inifile,encoding='GBK')


#########################################################################
#读取键值,如果键值不存在，就设置为defvl
#########################################################################
def readkey(config,key,defvl=None):
    keys = config.keys()
    if defvl==None :
        if keys.count(key) :
            return config[key]
        else :
            return ""
    else :
        if not keys.count(key) :
            config[key] = defvl
            config.write()
            return defvl
        else:
            return config[key]


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



def getdrive():
    return sys.argv[0][:2]

        
if __name__ == "__main__": 
    print('%s Running' % sys.argv[0])
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    today=datetime.datetime.now().strftime('%Y-%m-%d')
    config = iniconfig()
    lastdate = readkey(config,'lastdate')

    dbfn=getdrive()+'\\hyb\\STOCKDATA.db'
    dbcn = sqlite3.connect(dbfn)

#    browser = webdriver.PhantomJS()

    fireFoxOptions = webdriver.FirefoxOptions()
    fireFoxOptions.headless=True
    browser = webdriver.Firefox(options=fireFoxOptions)

#    chrome_options = webdriver.ChromeOptions() 
#    chrome_options.add_argument("--headless") 
#    chrome_options.add_argument('--disable-gpu')
#    browser = webdriver.Chrome(chrome_options=chrome_options) 

    url='http://data.eastmoney.com/gdhs/'
    browser.get(url)
    try:
        '''
        EC.presence_of_element_located()传递的参数是tuple元组
        '''
        elem=WebDriverWait(browser, 3).until(
            EC.presence_of_element_located((By.XPATH, "//a[text()='下一页']//preceding-sibling::a[1]")))
        pgs=int(elem.text)
    except:
        pgs=1

    pgn=1   
#    pgs=1
    while pgn<=pgs:

        print("正在处理第%d/%d页，请等待。" % (pgn,pgs))
        if pgn>1:
            try :    
                elem = browser.find_element_by_id("PageContgopage")
                elem.clear()
                #输入页面
                elem.send_keys(pgn)
                elem = browser.find_element_by_class_name("btn_link")     
                #点击Go
                elem.click()
                                
                #定位到表体
                tbl = WebDriverWait(browser, 3).until(
                        EC.presence_of_element_located((By.ID, "dt_1")))
            except :
                dbcn.close()
                browser.quit()
                print("0出错退出")
                sys.exit()
        else:
            try:
                tbl = WebDriverWait(browser, 3).until(
                        EC.presence_of_element_located((By.ID, "dt_1")))
            except :
                dbcn.close()
                browser.quit()
                print("1出错退出")
                sys.exit()

        tbody = tbl.find_element_by_tag_name('tbody')
        html=pq(tbody.get_attribute('innerHTML'))

        html.find("script").remove()    # 清理 <script>...</script>
        html.find("style").remove()     # 清理 <style>...</style>

        rows=html('tr')
        '''
        Python中的jquery PyQuery库使用小结
        https://www.jb51.net/article/50069.htm        

        '''
        #遍历行
        data = []
        sc=True     #本页处理成功
        for i in range(len(rows)):

            try:

                row=pq(rows('tr').eq(i))
                dm = row('td').eq(0).text()
                hs = row('td').eq(5).text()
                rq = row('td').eq(10).text()
    
                ggrq = pq(row('td.col'))('span').attr('title')
                ggrq=ggrq.replace('/','-')
     
                if dm != None and rq != None and hs != None :
                    dm = lgpdm(dm)       
                    rq = rq.replace('/','-')
                    rowdat = [dm,rq,hs,ggrq]
                    data.append(rowdat)
    
            except:
                print('处理第%d页第%d行出错！' % (pgn,i))
                sc=False    #本页处理不成功
                break
        
        if len(data)>0 and sc:
            dbcn.executemany('INSERT OR REPLACE INTO GDHS (GPDM,RQ,GDHS,GGRQ) VALUES (?,?,?,?)', data)
            dbcn.commit()
            pgn+=1
        else:
            try:
                browser.get(url)
                '''
                EC.presence_of_element_located()传递的参数是tuple元组
                '''
                elem=WebDriverWait(browser, 10).until(
                        EC.presence_of_element_located((By.ID,"PageContgopage")))            
            except:
                dbcn.close()
                browser.quit()
                print("2出错退出")
                sys.exit()
                

        if ggrq<lastdate:
            break
    
    browser.quit()    

    dbcn.commit()
    dbcn.close()
    
    config['lastdate'] = today
    config.write()
    
    now2 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)
    print('结束运行时间：%s' % now2)
    