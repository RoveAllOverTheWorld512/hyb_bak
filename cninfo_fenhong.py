# -*- coding: utf-8 -*-
"""
Created on Tue Feb  7 22:17:19 2017

@author: huangyunbin@sina.com
"""

import os
import sys
import re
import getopt
import datetime  
import time
from selenium import webdriver
from html.parser import HTMLParser
from configobj import ConfigObj  
import xlwt
ezxf= xlwt.easyxf

########################################################################
#初始化本程序配置文件
########################################################################
def iniconfig():
    myname=filename(sys.argv[0])
    wkdir = os.getcwd()
    inifile = os.path.join(wkdir,myname+'.ini')  #设置缺省配置文件
    return ConfigObj(inifile,encoding='GBK') 
        
#########################################################################
#读INI文件
#########################################################################
def readini(inifile):
    config = ConfigObj(inifile,encoding='GBK')
    return config
    
#########################################################################
#读取键值
#########################################################################
def readkey(config,key):
    keys = config.keys()
    if keys.count(key) :
        return config[key]
    else :
        return ""

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def str2float(num):
    try:
        return float(num)
    except ValueError:
        return num

 

def writesheet(book_name,sheet_name, headings, data, heading_xf, data_xfs, width_xfs=None):
    sheet = book_name.add_sheet(sheet_name)
    rowx = 0
    for colx, value in enumerate(headings):
        sheet.write(rowx, colx, value,heading_xf)

    sheet.set_panes_frozen(True)        # 冻结窗口
    sheet.set_horz_split_pos(1)         # 冻结行数
    sheet.set_vert_split_pos(3)         #冻结列数
    sheet.set_remove_splits(True)       # 使用冻结窗口不能分屏

    for row in data:
        rowx += 1
        for colx, value in enumerate(row):
            sheet.write(rowx, colx, value,data_xfs[colx])

    for colx, width in enumerate(width_xfs):
        sheet.col(colx).width = 256*width

def write_xls(xlsfile,data):
    book = xlwt.Workbook()

    hdngs = data[0]
    data = data[1:]

    kinds = 'ctxt ctxt rtxt flt ctxt flt flt ctxt text'.split()
    widths= 'wd1  wd2  wd2  wd1 wd2  wd5 wd3 wd2  wd4'.split()
    heading_xf = ezxf('font: bold on; align:wrap on, vert centre, horiz center')
    kind_to_xf_map = {
        'int': ezxf(num_format_str='#0'),
        'flt': ezxf(num_format_str='#0.00'),
        'text': ezxf('align:wrap on'),
        'ltxt': ezxf("align:horiz left"),
        'ctxt': ezxf("align:horiz center"),
        'rtxt': ezxf("align:horiz right"),
        }
    data_xfs = [kind_to_xf_map[k] for k in kinds]
    width_to_xf_map = {
        'wd1':6,
        'wd2':10,
        'wd3':16,
        'wd4':60,
        'wd5':20,
        }
    width_xfs = [width_to_xf_map[k] for k in widths]
    writesheet(book,'业绩预告', hdngs, data, heading_xf, data_xfs, width_xfs)
   
    book.save(xlsfile)

class HTMLTableParser(HTMLParser):
    """ This class serves as a html table parser. It is able to parse multiple
    tables which you feed in. You can access the result per .tables field.
    这类作为一个HTML表分析器。它能够解析你传入的多个表。你能访问结果的每个.tables字段
    """
    def __init__(self, decode_html_entities=False, data_separator=' ', ):

        HTMLParser.__init__(self)

        self._parse_html_entities = decode_html_entities
        self._data_separator = data_separator

        self._in_td = False
        self._in_th = False
        self._current_table = []
        self._current_row = []
        self._current_cell = []
        self.tables = []

    def handle_starttag(self, tag, attrs):
        """ We need to remember the opening point for the content of interest.
        The other tags (<table>, <tr>) are only handled at the closing point.
        我们需要记住感兴趣内容的开始点。其它标签(<table>, <tr>)是仅在关闭点处理。
        """
        if tag == 'td':
            self._in_td = True
        if tag == 'th':
            self._in_th = True

    def handle_data(self, data):
        """ This is where we save content to a cell 
        在这里保存内容
        """
        if self._in_td or self._in_th:
            self._current_cell.append(data.strip())

    def handle_charref(self, name):
        """ Handle HTML encoded characters 
        处理HTML编码字符
        """

        if self._parse_html_entities:
            self.handle_data(self.unescape('&#{};'.format(name)))

    def handle_endtag(self, tag):
        """ Here we exit the tags. If the closing tag is </tr>, we know that we
        can save our currently parsed cells to the current table as a row and
        prepare for a new row. If the closing tag is </table>, we save the
        current table and prepare for a new one.
        这里是退出标签。如果
        """
        if tag == 'td':
            self._in_td = False
        elif tag == 'th':
            self._in_th = False

        if tag in ['td', 'th']:
            final_cell = self._data_separator.join(self._current_cell).strip()
            self._current_row.append(final_cell)
            self._current_cell = []
        elif tag == 'tr':
            self._current_table.append(self._current_row)
            self._current_row = []
        elif tag == 'table':
            self.tables.append(self._current_table)
            self._current_table = []


def create_html(prof,username,pwd):
    
    today=datetime.datetime.now().strftime("%Y-%m-%d")
    
    cus_profile_dir = prof  # 你自定义profile的路径
    pzfn = cus_profile_dir + "\\prefs.js"
    with open(pzfn) as f:
        pzstr = f.read()
        f.close()
    
    if len(pzstr)==0 :
        sys.exit()
    
    pz = re.findall('download\.dir.{4}(.*)\"',pzstr)
    dldir = pz[0]
    dlfn = today+'.xls'
    fn = os.path.join(dldir,dlfn)
    
    if os.path.exists(fn):
        ctime=os.path.getctime(fn)  #文件建立时间
        ltime=time.localtime(ctime)
        newfn = time.strftime("%Y%m%d%H%M%S",ltime)+'.xls' 
        os.rename(os.path.join(dldir,dlfn),os.path.join(dldir,newfn))  
    
        
    cus_profile = webdriver.FirefoxProfile(cus_profile_dir)
    browser = webdriver.Firefox(cus_profile)
    
    #browser.implicitly_wait(30)
    #browser = webdriver.Firefox()
        
    #浏览器窗口最大化
    browser.maximize_window()
    #登录同花顺
    browser.get("http://www.cninfo.com.cn/information/companyinfo_n.html")
    #time.sleep(1)
    elem = browser.find_element_by_id("stockID_")
    elem.clear()
    elem.send_keys('002294')
    
    elem = browser.find_element_by_class_name("input2")
    elem.click()
    
    elem.send_keys(pwd)
    
    browser.find_element_by_id("loginBtn").click()
    time.sleep(2)
    
    #查询“i问财”网
    kw="连续2年主营业务收入增长率>10%,连续2年净利润增长率>10%，2016年9月30日主营业务收入增长率>0 2016年9月30日roe>10% 2016年12月31日业绩预增"
    #kw="2016年12月31日业绩预告"
    #kw="连续2年主营业务收入增长率>10%,连续2年净利润增长率>10%，2016年9月30日主营业务收入增长率>10% 2016年12月31日业绩预增 医药股"
    #kw="3季度营业收入同比增长率>10% 净利润同比增长率>10% 净利润同比增长率>营业收入同比增长率 经营活动现金流>购建固定资产、无形资产和其他长期资产支付的现金 2016年12月31日业绩预增 2014年1月1日前上市"
    browser.get("http://www.iwencai.com/")
    time.sleep(5)
    browser.find_element_by_id("auto").clear()
    browser.find_element_by_id("auto").send_keys(kw)
    browser.find_element_by_id("qs-enter").click()
    time.sleep(8)
    
    #打开查询项目选单
    trigger = browser.find_element_by_class_name("showListTrigger")
    trigger.click()
    time.sleep(1)
    #获取查询项目选单
    checkboxes = browser.find_elements_by_class_name("showListCheckbox")
    #去掉选项前的“√”
    for checkbox in checkboxes:
        if checkbox.is_selected():
            checkbox.click()
    #        time.sleep(1)
    #向上滚屏
    js="var q=document.documentElement.scrollTop=0"  
    browser.execute_script(js)  
    time.sleep(3) 
    #关闭查询项目选单
    trigger = browser.find_element_by_class_name("showListTrigger")
    trigger.click()
    time.sleep(3)
    
    #导出数据
    elem = browser.find_element_by_class_name("export.actionBtn.do") 
    #在html中类名包含空格
    elem.click() 
    time.sleep(10)
    
    #关闭浏览器
#    browser.quit()
    return fn
            

#############################################################################
#股票列表,通达信板块文件调用时wjtype="tdxbk"
#############################################################################    
def zxglist(zxgfn,wjtype=""):
    zxglst = []
    p = "(\d{6})"
    if wjtype == "tdxblk" :
        p ="\d(\d{6})"
    if os.path.exists(zxgfn) :
        #用二进制方式打开再转成字符串，可以避免直接打开转换出错
        with open(zxgfn,'rb') as dtf:
            zxg = dtf.read().decode('utf8','ignore')
        zxglst =re.findall(p,zxg)
    else:
        print("文件%s不存在！" % zxgfn)
    if len(zxglst)==0:
        print("股票列表为空,请检查%s文件。" % zxgfn)

    return zxglst

def filename(pathname):
    wjm = os.path.splitext(os.path.basename(pathname))
    return wjm[0]

def Usage():
    print ('用法:')
    print ('-h, --help: 显示帮助信息。')
    print ('-v, --version: 显示版本信息。')
    print ('-c, --cfg: 配置文件。')
    print ('-i, --input: 股票列表文本文件。')
    print ('-o, --output: 市盈率保存文件。')

def Version():
    print ('版本 2.0.0')

def main(argv):
    myname=filename(argv[0])
    wkdir = os.getcwd()
    inifile = os.path.join(wkdir,myname+'.ini')  #设置缺省配置文件
    
    try:
        opts, args = getopt.getopt(argv[1:], 'hvc:i:o:', ['help','version','cfg=','inpute=','output='])
    except (getopt.GetoptError):
        Usage()
        sys.exit(1)

#    td = datetime.datetime.now().strftime("%Y%m%d") #今天
    
    sylxls = ""

    zxgfile = ""    
    for o, a in opts:
        if o in ('-h', '--help'):
            Usage()
            sys.exit(0)
        elif o in ('-v', '--version'):
            Version()
            sys.exit(0)
        elif o in ('-c', '--cfg'):
            inifile = a
        elif o in ('-i', '--input'):
            zxgfile = a
        elif o in ('-o', '--output'):
            sylxls = a
        else:
            print ('无效参数！')
            sys.exit(3)


    if not os.path.exists(inifile) :
        print("配置文件%s不存在，无法运行，请检查。" % inifile)
        sys.exit(3)
        
    if len(inifile)==0:
        config = iniconfig()
    else:
        config = readini(inifile)

    tdxblkdir = readkey(config,'tdxblockdir')
    firefoxprof = readkey(config,'firefox_profiledir')

    if len(zxgfile)==0 :
        zxgfile = "zxg.blk"          #没有指定股票列表就用通达信自选股板块
            
    if zxgfile.upper().endswith(".BLK") :             #
        zxgfile = os.path.join(tdxblkdir,zxgfile)
        zxglb = zxglist(zxgfile,"tdxblk")
    else:
        zxglb =  zxglist(zxgfile) 
        
    if len(sylxls)== 0:
        zxgwjm = os.path.splitext(os.path.basename(zxgfile))
        sylxls = os.path.join(wkdir,zxgwjm[0]+"_yjyg.xls")

    usrn = readkey(config,'iwencaiusername')
    pwd = readkey(config,'iwencaipwd')
    htmlfn = create_html(firefoxprof,usrn,pwd)
    print(htmlfn)
    if os.path.exists(htmlfn) :
        html =open(htmlfn,'rb').read().decode("UTF")
    else:
        html=""
        
    if len(html)==0 :
        print("没有返回数据，查询终止")
        sys.exit()
        
    p = HTMLTableParser()
    p.feed(html)
    
    data = []
    data0 = p.tables[0][1:]
    for da in data0 :
        if da[0][:6] in zxglb:
            da[0] = da[0][:6]
            da[2] = str2float(da[2])
            da[3] = str2float(da[3])
            da[5] = str2float(da[5])
            da[6] = str2float(da[6])
            da.insert(0,zxglb.index(da[0]))
            del(da[4])
            data.append(da)

    data.sort(key = lambda x:x[0])        
    hdngs = p.tables[0][0]
    hdngs.insert(0,'序号')
    del(hdngs[4])
    data.insert(0,hdngs)
    write_xls(sylxls,data)   
    print("查询结果保存在%s文件中，请查看。" % sylxls)     

    
if __name__ == '__main__':
    main(sys.argv)
    
