# -*- coding: utf-8 -*-
"""
Created on Sat Feb  4 00:33:48 2017
查询市盈率
@author: huangyunbin@sina.com
"""

import os
import sys
import re
import getopt
import datetime
from configobj import ConfigObj
import xlrd
import xlwt
import winreg
import struct
import pandas as pd

ezxf= xlwt.easyxf


def writesheet1(book_name,sheet_name, headings, data, heading_xf, data_xfs, width_xfs=None):
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
        tsh = (isfloat(row[12]) and isfloat(row[15]) and row[12] < row[15])
        for colx, value in enumerate(row):
            if tsh :
                sheet.write(rowx, colx, value,data_xfs[colx][1])
            else :
                sheet.write(rowx, colx, value,data_xfs[colx][0])

    for colx, width in enumerate(width_xfs):
        sheet.col(colx).width = 256*width

def writesheet2(book_name,sheet_name, headings, data, heading_xf, data_xfs, width_xfs=None):
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
            sheet.write(rowx, colx, value,data_xfs[colx][0])

    for colx, width in enumerate(width_xfs):
        sheet.col(colx).width = 256*width

def write_xls(file,xlsfile,gg,hy):
    book = xlwt.Workbook()

    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name("个股数据")
    hdngs = table.row_values(0)
    hdngs.insert(0,'序号')
    hdngs.append('四级行业市盈率')
    hdngs.append('日期')
    kinds = 'cint ctxt text ctxt text ctxt text ctxt text ctxt text flt flt flt flt flt ctxt'.split()
    widths= 'wd1 wd2 wd2 wd1 wd3 wd1 wd3 wd2 wd3 wd2 wd4 wd2 wd2 wd2 wd2 wd2 wd3'.split()
    heading_xf = ezxf('font: bold on; align:wrap on, vert centre, horiz center')
    kind_to_xf_map = {
        'cint': [ezxf('align:horiz center',num_format_str='#0'),
                 ezxf('pattern: pattern solid,fore_colour red;align:horiz center',num_format_str='#0')],
        'int': [ezxf(num_format_str='#0'),
                ezxf('pattern: pattern solid,fore_colour red',num_format_str='#0')],
        'flt': [ezxf(num_format_str='#0.00'),
                ezxf('pattern: pattern solid,fore_colour red',num_format_str='#0.00')],
        'text': [ezxf(),
                 ezxf('pattern: pattern solid,fore_colour red')],
        'ctxt': [ezxf('align:horiz center'),
                 ezxf('pattern: pattern solid,fore_colour red;align:horiz center')],
        }
    data_xfs = [kind_to_xf_map[k] for k in kinds]
    width_to_xf_map = {
        'wd1':6,
        'wd2':10,
        'wd3':16,
        'wd4':30,
        }
    width_xfs = [width_to_xf_map[k] for k in widths]
    writesheet1(book,'个股数据', hdngs, gg, heading_xf, data_xfs, width_xfs)

    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name("中证行业滚动市盈率")
    hdngs = table.row_values(0)

    kinds = 'ctxt text flt ctxt ctxt ctxt ctxt ctxt ctxt'.split()
    widths= 'wd3 wd4 wd3 wd3 wd3 wd3 wd3 wd3 wd3'.split()

    data_xfs = [kind_to_xf_map[k] for k in kinds]

    width_xfs = [width_to_xf_map[k] for k in widths]

    writesheet2(book,'中证行业滚动市盈率', hdngs, hy, heading_xf, data_xfs, width_xfs)

    book.save(xlsfile)

########################################################################
#初始化本程序配置文件
########################################################################
def iniconfig():
    inifile = os.path.splitext(sys.argv[0])[0]+'.ini'  #设置缺省配置文件
    return ConfigObj(inifile,encoding='GBK')


def gettdxblkdir():
    try :
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人生")
        value, type = winreg.QueryValueEx(key, "InstallLocation")
        return value + '\\T0002\\blocknew'
    except :
        print("本机未安装【华西证券华彩人生】软件系统。")
        sys.exit()

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
#读取个股市盈率
########################################################################
def ggsyl(file,sheet,keylst,colkey=1,rowbg=2,s2flst=[11,12,13,14]):
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name(sheet)
    nrows = table.nrows #行数
    ncols = table.ncols #列数

    data =[]

    for rownum in range(rowbg,nrows+1):
        row = table.row_values(rownum-1)
        da = []
        if row[colkey-1] in keylst:
            da.append(keylst.index(row[colkey-1])+1)
            for colnum in range(1,ncols+1):
                if colnum in s2flst and isfloat(row[colnum-1]):
                    da.append(float(row[colnum-1]))
                else :
                    da.append(row[colnum-1])
            data.append(da)

    return data


########################################################################
#读取行业市盈率
########################################################################
def hysyl(file,sheet,keylst,colkey=1,rowbg=2,s2flst=[3]):
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name(sheet)
    nrows = table.nrows #行数
    ncols = table.ncols #列数

    data =[]

    for rownum in range(rowbg,nrows+1):
        row = table.row_values(rownum-1)
        da = []
        if row[colkey-1] in keylst:
            for colnum in range(1,ncols+1):
                if colnum in s2flst and isfloat(row[colnum-1]):
                    da.append(float(row[colnum-1]))
                else :
                    da.append(row[colnum-1])
            data.append(da)

    return data

#########################################################################
#从配置文件中读取休市日期
#########################################################################
def readclosedate(config):
    keys = config.keys()
    if keys.count('stockclosedate') :
        return eval(config['stockclosedate'])
    else :
        return []

#########################################################################
#读取键值
#########################################################################
def readkey(config,key):
    keys = config.keys()
    if keys.count(key) :
        return config[key]
    else :
        return ""


##########################################################################
#将字符串转换为时间戳，不成功返回errdate
##########################################################################
def str2datetime(s):
    try:
        dt = datetime.datetime(int(s[:4]),int(s[4:6]),int(s[6:8]))
    except(ValueError):
        dt = "errdate"
    return dt

##########################################################################
#n天后日期串，不成功返回errdate
##########################################################################
def nextdtstr(s,n):
    dt = str2datetime(s)
    if dt != "errdate" :
        dt += datetime.timedelta(n)
        return dt.strftime("%Y%m%d")
    else :
        return "errdate"

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

##########################################################################
#用法
##########################################################################
def Usage():
    print ('用法:')
    print ('-h, --help: 显示帮助信息。')
    print ('-v, --version: 显示版本信息。')
    print ('-i, --input: 股票列表文本文件。')
    print ('-o, --output: 市盈率保存文件。')

##########################################################################
#版本
##########################################################################
def Version():
    print ('版本 2.0.0')

def makedir(dirname):
    if not os.path.exists(dirname):
        try :
            os.mkdir(dirname)
        except(OSError):
            print("创建目录%s出错，请检查！" % dirname)
            return False
    else :
        return True

class PinYin(object):
    def __init__(self):
        self.word_dict = {}

    def load_word(self, dict_file):
        self.dict_file = dict_file
        if not os.path.exists(self.dict_file):
            raise IOError("NotFoundFile")

        with open(self.dict_file) as f_obj:
            for f_line in f_obj.readlines():
                try:
                    line = f_line.split('    ')
                    self.word_dict[line[0]] = line[1]
                except:
                    line = f_line.split('   ')
                    self.word_dict[line[0]] = line[1]


    def hanzi2pinyin(self, string="", firstcode=False):
        result = []

        for char in string:
            key = '%X' % ord(char)
            value = self.word_dict.get(key, char)
            outpinyin = str(value).split()[0][:-1].lower()
            if not outpinyin:
                outpinyin = char
            if firstcode:
                result.append(outpinyin[0])
            else:
                result.append(outpinyin)

        return result


    def hanzi2pinyin_split(self, string="", split="", firstcode=False):
        """提取中文的拼音
        @param string:要提取的中文
        @param split:分隔符
        @param firstcode: 提取的是全拼还是首字母？如果为true表示提取首字母，默认为False提取全拼
        """
        result = self.hanzi2pinyin(string=string, firstcode=firstcode)
        return split.join(result)

########################################################################
#获取本机通达信安装目录，生成自定义板块保存目录
########################################################################
def gettdxblk(lb):
    py = PinYin()
    py.load_word('word.data')

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
            shtname = py.hanzi2pinyin_split(string=blkname, split="", firstcode=True)
            stnum, = struct.unpack('H2x', f.read(4))
            for j in range(stnum) :
                stkid = f.read(7).strip(b'\x00').decode('GBK')
                stk.append(stkid)
            blk[blkname] = [blkname,shtname,stnum,stk]

            tmp = f.read((400-stnum)*7)
        f.close()

    da = []
    for key in blk.keys():
#        print(key)
        da.append([blk[key][0],blk[key][1],blk[key][2]])

    df = pd.DataFrame(da,columns=['blockname','shortname','num'])
    csvfn = 'block_'+lb+'.csv'
    try :
        df.to_csv(csvfn,encoding='GBK',sep=',',index=False)
    except :
        pass

    return blk

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

    zxg = list(set(zxglst))
    zxg.sort(key=zxglst.index)

    return zxg

def fn(pathname):
    wjm = os.path.splitext(os.path.basename(pathname))
    return wjm[0]

def main(argv):

    try:
        opts, args = getopt.getopt(argv[1:], 'hvd:k:i:o:',
                   ['help','version','date=','kind=','inpute=','output='])
    except (getopt.GetoptError):
        Usage()
        sys.exit(1)

    td = datetime.datetime.now().strftime("%Y%m%d") #今天

    sylxls = ""
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
            sylxls = a
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

        if len(sylxls)== 0:

            zxgwjm = os.path.splitext(os.path.basename(zxgfile))
            sylxls = os.path.join(getpath(),'zd_'+zxgwjm[0]+"_syl.xls")
    else :
        sylxls = os.path.join(getpath(),bkxx+"_syl.xls")



#    print(zxglb)

    ggsj = ggsyl(sylfn,'个股数据',zxglb)
    hylb = [e[9] for e in ggsj]
    hysj = hysyl(sylfn,'中证行业滚动市盈率',hylb)

    hysyllb = [[e[0],e[2]] for e in hysj]
    hydmb = [e[0] for e in hysj]
    zxgg = []
    for gg in ggsj:
        szhysyl = hysyllb[hydmb.index(gg[9])][1]
        if isfloat(szhysyl) :
            szhysyl = float(szhysyl)

        gg.append(szhysyl)
        gg.append(jyrq)

        zxgg.append(gg)

    zxgg.sort(key = lambda x:x[0])

    write_xls(sylfn,sylxls,zxgg,hysj)
    print("股票列表文件%s" % zxgfile)
    print("查询结果保存在%s文件中，请查看。" % sylxls)

if __name__ == '__main__':
    main(sys.argv)


