# -*- coding: utf-8 -*-
"""
Created on Fri Nov 10 22:19:26 2017

@author: lenovo
"""

# coding=utf-8
import os
import re
from selenium import webdriver
import selenium.webdriver.support.ui as ui
import time
from datetime import datetime
import IniFile
# from threading import Thread
from pyquery import PyQuery as pq
import LogFile
import mongoDB
import urllib
class toutiaoSpider(object):
    def __init__(self):

        logfile = os.path.join(os.path.dirname(os.getcwd()), time.strftime('%Y-%m-%d') + '.txt')
        self.log = LogFile.LogFile(logfile)
        configfile = os.path.join(os.path.dirname(os.getcwd()), 'setting.conf')
        cf = IniFile.ConfigFile(configfile)
        webSearchUrl = cf.GetValue("toutiao", "webSearchUrl")
        self.keyword_list = cf.GetValue("section", "information_keywords").split(';')
        self.db = mongoDB.mongoDbBase()
        self.start_urls = []

        for word in self.keyword_list:
            self.start_urls.append(webSearchUrl + urllib.quote(word))

        self.driver = webdriver.PhantomJS()
        self.wait = ui.WebDriverWait(self.driver, 2)
        self.driver.maximize_window()

    def scroll_foot(self):
        '''
                滚动条拉到底部
                :return:
                '''
        js = ""
        # 如何利用chrome驱动或phantomjs抓取
        if self.driver.name == "chrome" or self.driver.name == 'phantomjs':
            js = "var q=document.body.scrollTop=10000"
        # 如何利用IE驱动抓取
        elif self.driver.name == 'internet explorer':
            js = "var q=document.documentElement.scrollTop=10000"
        return self.driver.execute_script(js)



    def date_isValid(self, strDateText):
        '''
        判断日期时间字符串是否合法：如果给定时间大于当前时间是合法，或者说当前时间给定的范围内
        :param strDateText: 四种格式 '2小时前'; '2天前' ; '昨天' ;'2017.2.12 '
        :return: True:合法；False:不合法
        '''
        currentDate = time.strftime('%Y-%m-%d')
        if strDateText.find('分钟前') > 0 or strDateText.find('刚刚') > -1:
            return True, currentDate
        elif strDateText.find('小时前') > 0:
            datePattern = re.compile(r'\d{1,2}')
            ch = int(time.strftime('%H'))  # 当前小时数
            strDate = re.findall(datePattern, strDateText)
            if len(strDate) == 1:
                if int(strDate[0]) <= ch:  # 只有小于当前小时数，才认为是今天
                    return True, currentDate
        return False, ''


    def log_print(self, msg):
        '''
        #         日志函数
        #         :param msg: 日志信息
        #         :return:
        #         '''
        print('%s: %s' % (time.strftime('%Y-%m-%d %H-%M-%S'), msg))

    def scrapy_date(self):
        strsplit = '------------------------------------------------------------------------------------'
        index = 0
        for link in self.start_urls:
            self.driver.get(link)

            keyword = self.keyword_list[index]
            index = index + 1
            time.sleep(1) #数据比较多，延迟下，否则会出现查不到数据的情况

            selenium_html = self.driver.execute_script("return document.documentElement.outerHTML")
            doc = pq(selenium_html)
            infoList = []
            self.log.WriteLog(strsplit)
            self.log_print(strsplit)

            Elements = doc('div[class="articleCard"]')

            for element in Elements.items():
                strdate = element('span[class="lbtn"]').text().encode('utf8').strip()
                flag, date = self.date_isValid(strdate)
                if flag:
                    title = element('a[class="link title"]').find('span').text().encode('utf8').replace(' ', '')
                    if title.find(keyword) > -1:
                        url = 'http://www.toutiao.com' + element.find('a[class="link title"]').attr('href')
                        source = element('a[class="lbtn source J_source"]').text().encode('utf8').replace(' ', '')

                        dictM = {'title': title, 'date': date,
                                 'url': url, 'keyword': keyword, 'introduction': title, 'source': source}
                        infoList.append(dictM)
                        # self.log.WriteLog('title:%s' % title)
                        # self.log.WriteLog('url:%s' % url)
                        # self.log.WriteLog('source:%s' % source)
                        # self.log.WriteLog('kword:%s' % keyword)
                        # self.log.WriteLog(strsplit)

                        self.log_print('title:%s' % dictM['title'])
                        self.log_print('url:%s' % dictM['url'])
                        self.log_print('date:%s' % dictM['date'])
                        self.log_print('source:%s' % dictM['source'])
                        self.log_print('kword:%s' % dictM['keyword'])
                        self.log_print(strsplit)


            if len(infoList)>0:
                self.db.SaveInformations(infoList)

        self.driver.close()
        self.driver.quit()

obj = toutiaoSpider()
obj.scrapy_date()