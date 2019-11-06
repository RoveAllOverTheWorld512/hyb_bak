# -*- coding: utf-8 -*-
"""
Created on Sun Jan 29 18:34:52 2017

@author: huangyunbin@sina.com
"""
import getpass
import os
import sys
import re
import struct
import getopt
import datetime
import time
import math
import pandas as pd
import winreg
from selenium import webdriver
from html.parser import HTMLParser
from configobj import ConfigObj
import xlrd
import xlwt
ezxf= xlwt.easyxf

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False


########################################################################
#读取个股滚动市盈率
########################################################################
def ggttmpepd(file,sheet):
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name(sheet)
    nrows = table.nrows #行数
    data =[]

    for rownum in range(1,nrows):
        row = table.row_values(rownum)
        data.append([row[0],row[8],row[9],float(row[11]) if isfloat(row[11]) else '--'])

    pe = pd.DataFrame(data,columns=['gpdm','hydm','hymc','pettm'])
    return pe.set_index('gpdm')

########################################################################
#读取行业滚动市盈率
########################################################################
def hyttmpepd(file,sheet):
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name(sheet)
    nrows = table.nrows #行数
    data =[]

    for rownum in range(1,nrows):
        row = table.row_values(rownum)
        data.append([row[0],row[1],float(row[2]) if isfloat(row[2]) else '--'])

    pe = pd.DataFrame(data,columns=['hydm','hymc','hypettm'])
    return pe.set_index('hydm')




########################################################################
#获取firefoxprofile路径
########################################################################
def getfirefoxprofiledir():
    user = getpass.getuser()
    firefoxprofilepath = 'C:\\Users\\'+user+'\\AppData\\Roaming\\Mozilla\\Firefox'
    inifile = firefoxprofilepath + '\\profiles.ini'
    config = ConfigObj(inifile,encoding='GBK')
    pfnum = config['General']['StartWithLastProfile']
    pfs = 'Profile'+str(pfnum)
    IsRelative = config[pfs]['IsRelative']
    if IsRelative == '0' :               #绝对路径
        return config[pfs]['Path']
    else :
        return firefoxprofilepath +'\\'+ config[pfs]['Path']

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

    kinds = ('ctxt ctxt flt flt flt flt ' +
             'flt flt flt flt flt flt flt flt ' +
             'flt flt '+'flt flt flt flt flt flt flt '+'ctxt ltxt flt ' +
             'ctxt ctxt flt wtxt ctxt').split()

    widths= ('wd1 wd2 wd2 wd2 wd2 wd2 '+
             'wd2 wd2 wd2 wd2 wd2 wd2 wd2 wd2 '+
             'wd2 wd2 '+'wd2 wd2 wd2 wd2 wd2 wd2 wd2 '+'wd3 wd3 wd3 '+
             'wd3 wd5 wd3 wd4 wd3').split()

    heading_xf = ezxf('font: bold on; align:wrap on, vert centre, horiz center')
    kind_to_xf_map = {
        'int': ezxf(num_format_str='#0'),
        'flt': ezxf(num_format_str='#0.00'),
        'text': ezxf('align:wrap on'),
        'ltxt': ezxf("align:horiz left"),
        'ctxt': ezxf("align:horiz center"),
        'wtxt': ezxf("align:wrap on"),
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
    if len(hdngs) >= len(data_xfs) and len(data_xfs) <= len(width_xfs) :
        writesheet(book,'业绩预告', hdngs, data, heading_xf, data_xfs, width_xfs)
    else :
        print('表头长度或格式与数据列数不符')

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

def gettdxblkdir():
    try :
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人生")
        value, type = winreg.QueryValueEx(key, "InstallLocation")
        return value + '\\T0002\\blocknew'
    except :
        print("本机未安装【华西证券华彩人生】软件系统。")
        sys.exit()


###############################################################################
#下载文件名，参数1表示如果文件存在则将原有文件名用其创建时间命名
###############################################################################
def dlfn(n=0):
    cus_profile_dir = getfirefoxprofiledir()  # 你自定义profile的路径
    today=datetime.datetime.now().strftime("%Y-%m-%d")
    pzfn = cus_profile_dir + "\\prefs.js"
    with open(pzfn,'rb') as f:
        pzstr = f.read().decode('utf8','ignore')
        f.close()

    if len(pzstr)==0 :
        sys.exit()

    pz = re.findall('download\.dir.{4}(.*)\"',pzstr)
    dldir = pz[0]
    dlfn = today+'.xls'
    fn = os.path.join(dldir,dlfn)

    if n==1 :
        if os.path.exists(fn):
            ctime=os.path.getctime(fn)  #文件建立时间
            ltime=time.localtime(ctime)
            newfn = time.strftime("%Y%m%d%H%M%S",ltime)+'.xls'
            os.rename(fn,os.path.join(os.path.dirname(fn),newfn))

        create_html()
    return fn

#############################################################################
#获取市盈率文件交易日列表
#############################################################################
def jyrlist(syldir):
    files = os.listdir(syldir)
    fs = [re.findall('csi(\d{8})\.xls',e) for e in files]
    jyrqlist =[]
    for e in fs:
        if len(e)>0:
            jyrqlist.append(e[0])

    return sorted(jyrqlist,reverse=1)


#############################################################################
#获取市盈率文件名
#############################################################################
def sylfilename():
    syldir = getdrive()+'\\syl'

    jyrlst = jyrlist(syldir)
    jyrq = jyrlst[0]

    return os.path.join(syldir,"csi"+jyrq+".xls")

#############################################################################
#
#############################################################################
def create_html():

    config = iniconfig()
    cus_profile_dir = getfirefoxprofiledir()  # 你自定义profile的路径
    username = readkey(config,'iwencaiusername')
    pwd = readkey(config,'iwencaipwd')
    kw = readkey(config,'iwencaikw')

    cus_profile = webdriver.FirefoxProfile(cus_profile_dir)
    browser = webdriver.Firefox(cus_profile)
#    browser = webdriver.Firefox()

    #browser.implicitly_wait(30)
    #browser = webdriver.Firefox()

    #浏览器窗口最大化
    browser.maximize_window()
    #登录同花顺
    browser.get("http://upass.10jqka.com.cn/login")
    #time.sleep(1)
    elem = browser.find_element_by_id("username")
    elem.clear()
    elem.send_keys(username)

    elem = browser.find_element_by_class_name("pwd")
    elem.clear()
    elem.send_keys(pwd)

    browser.find_element_by_id("loginBtn").click()
    time.sleep(2)

    #kw="连续2年主营业务收入增长率>10%,连续2年净利润增长率>10%，2016年9月30日主营业务收入增长率>10% 2016年12月31日业绩预增 医药股"
    #kw="3季度营业收入同比增长率>10% 净利润同比增长率>10% 净利润同比增长率>营业收入同比增长率 经营活动现金流>购建固定资产、无形资产和其他长期资产支付的现金 2016年12月31日业绩预增 2014年1月1日前上市"
    browser.get("http://www.iwencai.com/")
    time.sleep(5)
    browser.find_element_by_id("auto").clear()
    browser.find_element_by_id("auto").send_keys(kw)
    browser.find_element_by_id("qs-enter").click()
    time.sleep(20)

    #打开查询项目选单
#    trigger = browser.find_element_by_class_name("showListTrigger")
#    trigger.click()
#    time.sleep(1)
    #获取查询项目选单
#    checkboxes = browser.find_elements_by_class_name("showListCheckbox")
    #去掉选项前的“√”
#    for checkbox in checkboxes:
#        if checkbox.is_selected():
#            checkbox.click()
    #        time.sleep(1)
    #向上滚屏
#    js="var q=document.documentElement.scrollTop=0"
#    browser.execute_script(js)
#    time.sleep(3)
#    #关闭查询项目选单
#    trigger = browser.find_element_by_class_name("showListTrigger")
#    trigger.click()
#    time.sleep(3)

    #导出数据
    elem = browser.find_element_by_class_name("export.actionBtn.do")
    #在html中类名包含空格
    elem.click()
    time.sleep(10)

    #关闭浏览器
#    browser.quit()


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

##########################################################################
#获取运行程序所在驱动器
##########################################################################
def getdrive():
    return os.path.splitdrive(sys.argv[0])[0]

##########################################################################
#获取运行程序所在驱动器
##########################################################################
def getpath():
    return os.path.dirname(sys.argv[0])


def Usage():
    print ('用法:')
    print ('-h, --help: 显示帮助信息。')
    print ('-v, --version: 显示版本信息。')
    print ('-d, --download: 下载数据。')
    print ('-i, --input: 股票列表文本文件。')
    print ('-o, --output: 市盈率保存文件。')

def Version():
    print ('版本 2.0.0')

########################################################################
#获取滚动市盈率
########################################################################
def getpe(stock,pe):
    try :
        pexx = pe.ix[stock][:]
    except :
        pexx = False
    return pexx

########################################################################
#获取上市日期
########################################################################
def getssrq(stock,sssj):
    try :
        ssrq = sssj.ix[stock]['ssdate']
    except :
        ssrq = ''
    return ssrq

########################################################################
#获取本机通达信安装目录，生成自定义板块保存目录
########################################################################
def gettdxdir():

    try :
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人生")
        value, type = winreg.QueryValueEx(key, "InstallLocation")
    except :
        print("本机未安装【华西证券华彩人生】软件系统。")
        sys.exit()
    return value

#############################################################################
#读取dbf文件
#############################################################################
def dbfreader(f):

    numrec, lenheader = struct.unpack('<xxxxLH22x', f.read(32))
    numfields = (lenheader - 33) // 32

    fields = []
    for fieldno in range(numfields):
        name, typ, size, deci = struct.unpack('<11sc4xBB14x', f.read(32))
        name = name.decode().replace('\x00', '')
        typ  = typ.decode()
        fields.append((name, typ, size, deci))
    yield [field[0] for field in fields]
    yield [tuple(field[1:]) for field in fields]

    terminator = f.read(1)

    fields.insert(0, ('DeletionFlag', 'C', 1, 0))
    fmt = ''.join(['%ds' % fieldinfo[2] for fieldinfo in fields])
    fmtsiz = struct.calcsize(fmt)

    for i in range(numrec):
        record = struct.unpack(fmt, f.read(fmtsiz))
        if record[0].decode() != ' ':
            continue                        # deleted record
        result = []
        for (name, typ, size, deci), value in list(zip(fields, record)):
            if name == 'DeletionFlag':
                continue
            if typ == "C":
                value = value.strip(b'\x00').decode('GBK')
            if typ == "N":
                value = value.strip(b'\x00').strip(b'\x20').decode('GBK')
                if value == '':
                    value = 0
                elif deci:
                    value = float(value)
                else:
                    value = int(value)
            elif typ == 'D':
                value = value.decode('GBK')
            elif typ == 'L':
                value = value.decode('GBK')
                value = (value in 'YyTt' and 'T') or (value in 'NnFf' and 'F')

            result.append(value)

        yield result


########################################################################
#获取本机通达信安装目录，生成自定义板块保存目录
########################################################################
def dbf2pandas(dbffn,cols):
    with open(dbffn,"rb") as f:
        data = list(dbfreader(f))
        f.close()
    columns = data[0]
    columns=[e.lower() for e in columns]
    data = data[2:]
    df = pd.DataFrame(data,columns=columns)
    if len(cols) == 0 :
        return df
    else :
        return df[cols]

#############################################################################
#创建目录
#############################################################################
def makedir(dirname):
    if dirname == None :
        return False

    if not os.path.exists(dirname):
        try :
            os.mkdir(dirname)
            return True
        except(OSError):
            print("创建目录%s出错，请检查！" % dirname)
            return False
    else :
        return True


#############################################################################
#读取中证行业代码
#############################################################################
def zzhy(hydms):
    hylb = re.findall('(.+)_',hydms)[0].lower()
    hydm = re.findall('_(.+)',hydms)[0].upper()
    dmcd = len(hydm)
    if hylb != 'zz' :
        print('参数不对！')
        return None

    file = getdrive()+'\\syl\\csi'+jyrlist()[0]+'.xls'
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name('个股数据')
    nrows = table.nrows #行数

    zxglb = []
    for rownum in range(1,nrows):
        row = table.row_values(rownum)
        if row[8][:dmcd] == hydm :
            zxglb.append(row[0])

    return zxglb

########################################################################
#根据通达信新行业或申万行业代码提取股票列表
########################################################################
def hy(hydms):
    hylb = re.findall('(.+)_',hydms)[0].lower()
    hydm = re.findall('_(.+)',hydms)[0].upper()
    dmcd = len(hydm)
    if hylb not in ['tdx','sw'] :
        print('参数不对！')
        return None

    p = '(\d{6})\t(.+)\t(.*?)\r\n'
    zxgfn = gettdxdir()+r'T0002\hq_cache\tdxhy.cfg'
    with open(zxgfn,'rb') as dtf:
        zxg = dtf.read()
        if zxg[:3] == b'\xef\xbb\xbf' :
            zxg = zxg.decode('UTF8','ignore')   #UTF-8
        elif zxg[:2] == b'\xfe\xff' :
            zxg = zxg.decode('UTF-16','ignore')  #Unicode big endian
        elif zxg[:2] == b'\xff\xfe' :
            zxg = zxg.decode('UTF-16','ignore')  #Unicode
        else :
            zxg = zxg.decode('GBK','ignore')      #ansi编码
    zxg=zxg.replace('|','\t')
    zxglst =re.findall(p,zxg)

    if hylb == 'tdx' :
        zxglb = [gpdm for gpdm,tdxnhy,swhy in zxglst if tdxnhy[:dmcd] == hydm]
    else :
        zxglb = [gpdm for gpdm,tdxnhy,swhy in zxglst if swhy[:dmcd] == hydm]

    return zxglb

########################################################################
#获取本机通达信安装目录，生成自定义板块保存目录
########################################################################
def gettdxblk(lb):

    try :
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人生")
        value, type = winreg.QueryValueEx(key, "InstallLocation")
    except :
        print("本机未安装【华西证券华彩人生】软件系统。")
        sys.exit()

    blkfn = value + '\\T0002\\hq_cache\\block_'+lb+'.dat'
    blk = {}
    with open(blkfn,'rb') as f :
        blknum, = struct.unpack('384xH', f.read(386))
        for i in range(blknum) :
            stk = []
            blkname = f.read(9).strip(b'\x00').decode('GBK')
            stnum, = struct.unpack('H2x', f.read(4))
            for j in range(stnum) :
                stkid = f.read(7).strip(b'\x00').decode('GBK')
                stk.append(stkid)
            blk[blkname] = [blkname,stnum,stk]

            tmp = f.read((400-stnum)*7)
        f.close()


    return blk

def calcpeg(zxglb,pefn):
    basefn = gettdxdir()+r'\T0002\hq_cache\base.dbf'
    sssj = dbf2pandas(basefn,['gpdm','ssdate'])
    sssj = sssj.set_index('gpdm')

#    pefn = sylfilename()

    hype = hyttmpepd(pefn,'中证行业滚动市盈率')
    ggpe = ggttmpepd(pefn,'个股数据')

    today=datetime.datetime.now().strftime("%Y-%m-%d")

    htmlfn = today+'.xls'

    if os.path.exists(htmlfn) :
        html =open(htmlfn,'rb').read().decode("UTF")
    else:
        html=""

    if len(html)==0 :
        print("没有返回数据，查询终止")
        sys.exit()

    p = HTMLTableParser()
    p.feed(html)

    hd = ['股票代码','股票简称','现价','自由流通股','流通a股(股)',
    'a股流通市值','peg(cagr)','peg(ttm)','个股滚动市盈率','eps复合增长率','基本每股收益(同比增长率)(%)2017.03.31','基本每股收益(同比增长率)(%)2016.12.31',
    '基本每股收益(同比增长率)(%)2015.12.31',
    '基本每股收益(同比增长率)(%)2014.12.31','基本每股收益(同比增长率)(%)2013.12.31',
    '基本每股收益(元)2016.12.31','基本每股收益(元)2015.12.31','基本每股收益(元)2014.12.31',
    '基本每股收益(元)2013.12.31','销售毛利率(%)2016.12.31','销售毛利率(%)2015.12.31','销售毛利率(%)2014.12.31',
    '四级行业代码','四级行业名称','行业滚动市盈率',
    '上市日期','预告日期','预告净利润变动幅度','业绩预告摘要','业绩预告类型']

    ####注意下面增加了“序号”
    hd1 = ['序号','股票代码','股票简称','现价','自由流通股(亿股)','流通a股(亿股)',
    'a股流通市值(亿元)','peg(cagr)','peg(ttm)','个股滚动市盈率','eps复合增长率','2017Q1eps增长率','2016eps增长率','2015eps增长率',
    '2014eps增长率','2013eps增长率',
    'eps2016','eps2015','eps2014','eps2013','毛利率2016','毛利率2015','毛利率2014',
    '四级行业代码','四级行业名称','行业滚动市盈率',
    '上市日期','预告日期','预告净利润变动幅度','业绩预告摘要','业绩预告类型']
    hdngs = p.tables[0][0]
    hdn =[]
    for i in range(len(hd)):
        n = -1
        for j in range(len(hdngs)):
            if hd[i] in hdngs[j]:
                n = j
                continue
        hdn.append(n)

    if hdn[0] == -1 :
        print('必须包含股票代码！')
        sys.exit()

    print(hdn)
    data = []

    data0 = p.tables[0][1:]

    for da in data0 :
        row = []
        gpdm = da[0][:6]
        if gpdm in zxglb :
            row.append(gpdm)
            for i in range(1,len(hdn)) :
                if hdn[i] != -1 :
                    if i in (3,4,5) :
                        if str2float(da[hdn[i]]) == '--':
                            row.append('--')
                        else :
                            row.append(str2float(da[hdn[i]])/100000000)
                    elif i in (2,10,11,12,13,14,15,16,17,18,19,20,21,27) :
                        row.append(str2float(da[hdn[i]]))
                    else :
                        row.append(da[hdn[i]])
                else:
                    row.append('--')
            ggxx = getpe(gpdm,ggpe)  #查询个股市盈率信息
            if not isinstance(ggxx,bool) :
                hydm = ggxx['hydm']
                hymc = ggxx['hymc']
                ggsyl = ggxx['pettm']
                hyxx = getpe(hydm,hype)
                if not isinstance(hyxx,bool)  :
                    hysyl = hyxx['hypettm']
                else :
                    hysyl = '--'

                row[8]=ggsyl
                row[22]=hydm
                row[23]=hymc
                row[24]=hysyl

            if row[15]!='--' and row[16]!='--' and row[17]!='--' and row[18]!='--' and row[15]>0 and row[16]>0 and row[17]>0 and row[18]>0 :
                if row[8] != '--' and row[11] != '--' and row[8] != 0 and row[11] > 0:
                    row[7] = row[8]/row[11]

                if row[11] != '--' and row[12] != '--' and row[13] != '--' and row[14] != '--' :
                    tmp = (row[11]+100)*(row[12]+100)*(row[13]+100)*(row[14]+100)/100000000
                    if tmp > 0:
                        row[9] = (math.pow(tmp,0.25)-1)*100.00

                if row[8] != '--' and row[9] != '--' and row[8] != 0 and row[9] > 0:
                    row[6] = row[8]/row[9]

            row[25] = getssrq(row[0],sssj)

            row.insert(0,zxglb.index(row[0])+1)

            data.append(row)

    data.sort(key = lambda x:x[0])
    data.insert(0,hd1)     #在数据前插入标题栏

    return data


#2016年12月31日业绩预告 流通A股 自由流通股 流通市值 2014年eps增长率  2015年eps增长率  2016年eps增长率
#2016年12月31日业绩预告 流通A股 自由流通股 流通市值 2013年eps增长率 2014年eps增长率  2015年eps增长率  2016年eps增长率


########################################################################
#提取文件名
########################################################################
def flnm(pathname):
    wjm = os.path.splitext(os.path.basename(pathname))
    return wjm[0]


###############################################################################
#主程序
###############################################################################
def main(argv):

    try:
        opts, args = getopt.getopt(argv[1:], 'hvd:k:i:o:',
                   ['help','version','date=','kind=','inpute=','output='])
    except (getopt.GetoptError):
        Usage()
        sys.exit(1)

    td = datetime.datetime.now().strftime("%Y%m%d") #今天

    pegxls = ""
    jyrq = td
    bklb = "zd"
    zxgfile = ""
    for o, a in opts:
        if o in ('-h', '--help'):
            Usage()
            sys.exit(0)
        elif o in ('-v', '--version'):
            Version()
            sys.exit(0)
        elif o in ('-d', '--date'):
            jyrq = a
        elif o in ('-k', '--kind'):
            bkxx = a
            bklb = re.findall('(.+)_',bkxx)[0].lower()
        elif o in ('-i', '--input'):
            zxgfile = a
        elif o in ('-o', '--output'):
            pegxls = a
        else:
            print ('无效参数！')
            sys.exit(3)

    if bklb not in ['fg','gn','zs','tdx','sw','zd','zz'] :
        print('板块类别参数不对，请查查。')
        sys.exit(3)
#    if bklb not in ['zd','fg','gn','zs']:
#        print ('无效参数！')
#        sys.exit(3)

    syldir = getdrive()+'\\syl'

    jyrlst = jyrlist(syldir)
    if not jyrq in jyrlst:
        jyrq = jyrlst[0]

    sylfn = os.path.join(syldir,"csi"+jyrq+".xls")

    if bklb in ['tdx','sw'] :
        zxglb = hy(bkxx)

    if bklb=='zz' :
        zxglb = zzhy(bkxx)

    if bklb in ['fg','gn','zs'] :
        bklb = re.findall('(.+)_',bkxx)[0].lower()
        bkjc = re.findall('_(.+)',bkxx)[0].upper()
        bkinfo = gettdxblk(bklb)
        try :
            zxglb = bkinfo[bkjc][2]
        except :
            zxglb = []

    if bklb == 'zd' :
        if len(zxgfile)==0 :
            zxgfile = "zxg.blk"          #没有指定股票列表就用通达信自选股板块
        if zxgfile.upper().endswith(".BLK") or zxgfile.upper().endswith(".EBK") :
            tdxblkdir = gettdxblkdir()
            zxgfile = os.path.join(tdxblkdir,zxgfile)
            zxglb = zxglist(zxgfile,"tdxblk")
        else:
            zxglb =  zxglist(zxgfile)

        if len(pegxls)== 0:

            zxgwjm = os.path.splitext(os.path.basename(zxgfile))
            pegxls = os.path.join(getpath(),'zd_'+zxgwjm[0]+"_peg.xls")
    else :
        pegxls = os.path.join(getpath(),bkxx+"_peg.xls")



    data = calcpeg(zxglb,sylfn)

    write_xls(pegxls,data)
    print("股票列表文件%s" % zxgfile)
    print("查询结果保存在%s文件中，请查看。" % pegxls)

if __name__ == '__main__':
    main(sys.argv)

###############################################################################
#2016年12月31日业绩预告 流通A股 自由流通股 流通市值 2013年eps增长率 2014年eps增长率
#2015年eps增长率  2016年eps增长率 2016年毛利率 2015年毛利率 2014年毛利率
#
###############################################################################



