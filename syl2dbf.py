# -*- coding: utf-8 -*-
"""
Created on Sat Feb  4 00:33:48 2017
查询市盈率
@author: huangyunbin@sina.com
"""

import os
import sys
import re
import struct
import datetime
import unicodedata
from configobj import ConfigObj
import xlrd
import winreg

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
def ggsyl(file,sheet,date):
    wb = xlrd.open_workbook(file,encoding_override="cp1252")
    table = wb.sheet_by_name(sheet)
    nrows = table.nrows #行数
    data =[]
    for rownum in range(1,nrows):
        row = table.row_values(rownum)
        if isfloat(row[11]) :
            data.append([row[0],row[1],date,float(row[11])])
        else :
            data.append([row[0],row[1],date,0])
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
            zxg = dtf.read()
            if zxg[:3] == b'\xef\xbb\xbf' :
                zxg = zxg.decode('UTF8','ignore')   #UTF-8
            elif zxg[:2] == b'\xfe\xff' :
                zxg = zxg.decode('UTF-16','ignore')  #Unicode big endian
            elif zxg[:2] == b'\xff\xfe' :
                zxg = zxg.decode('UTF-16','ignore')  #Unicode
            else :
                zxg = zxg.decode('GBK','ignore')      #ansi编码
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

#############################################################################
#检查字段类型和宽度，如果N，C型宽度不够，则扩展宽度，如果D，L型宽度不符则改为C型
#############################################################################
def checkdata(fieldnames, fieldspecs, records):
    flds = []
    for name, (typ, size, deci) in list(zip(fieldnames, fieldspecs)):
        i = fieldnames.index(name)
        if typ=="N" :
            p="{:>"+str(size)+"."+str(deci)+"f}"
            maxlen = max([len(p.format(e[i])) for e in records])
            if maxlen>size :
                size = maxlen
        if typ=="C" :
            maxlen = max([len(e[i]) for e in records])
            if maxlen>size :
                size = maxlen
        if typ in ("D","L"):
            maxlen = max([len(e[i]) for e in records])
            minlen = min([len(e[i]) for e in records])
            if maxlen!=8 or minlen!=8 :
                typ = 'C'
                size = maxlen
        flds.append([typ,size,deci])
    return flds

#############################################################################
#返回包含中文的byte字符串转的长度(一个汉字的长度为2)
#############################################################################
def str_width(s):
    w=0
    for c in s:
        if (unicodedata.east_asian_width(c) in ('F','W')):
            w +=2
        else:
            w +=1
    return(w)

#############################################################################
#将包含中文的byte字符串转变为指定长度（一个汉字为2个宽度,后面用空格补齐)
#############################################################################
def cnstrjust(cnstr,length):
    cnstrw=str_width(cnstr)
    if cnstrw>length :
        i=0
        while i<len(cnstr):
            i += 1
            cutstr = cnstr[:i]
            if str_width(cutstr)>length :
                break
        cnstr = cutstr[:i-1]
        cnstrw=str_width(cnstr)

    return cnstr+" "*(length-cnstrw)

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
    m = numrec//50
    for i in range(numrec):
#        if i % m == 0 :
#            print("已读取%d" % int(i/numrec*100))
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
                value = value.strip(b'\x00').decode('GBK')
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


#############################################################################
#写dbf文件
#############################################################################
def dbfwriter(f, fieldnames, fieldspecs, records):
    #对数据与字段的类型和宽度进行检查、优化
    fieldspecs = checkdata(fieldnames, fieldspecs, records)
    # 文件头部信息
    ver = 3
    now = datetime.datetime.now()
    yr, mon, day = now.year-2000, now.month, now.day
    numrec = len(records)
    numfields = len(fieldspecs)
    lenheader = numfields * 32 + 33
    lenrecord = sum(field[1] for field in fieldspecs) + 1
    codepageid = 122
    #Code Pages Supported by Visual FoxPro:936Chinese (PRC, Singapore) Windows
    #https://technet.microsoft.com/zh-cn/learning/aa975345
    hdr = struct.pack('<BBBBLHH17xB2x', ver, yr, mon, day, numrec, lenheader, lenrecord, codepageid)
    f.write(hdr)

    # 字段名信息
    addr = 1
    for name, (typ, size, deci) in list(zip(fieldnames, fieldspecs)):
        name = name.ljust(11, '\x00').encode('GBK')
        typ = typ.encode('GBK')
        fld = struct.pack('<11sciBB14x', name, typ, addr, size, deci)
        addr += size
        f.write(fld)

    # 终止符
    f.write('\r'.encode())

    # 记录
    for record in records:
        f.write(' '.encode())                        # deletion flag
        for (typ, size, deci), value in list(zip(fieldspecs, record)):
            if typ == "C":
                value = cnstrjust(value,size)
            if typ == "N":
                p="{:>"+str(size)+"."+str(deci)+"f}"
                value = p.format(value)

            if typ == 'D':
                value = value.ljust(8, ' ')
            if typ == 'L':
                value = value.upper()

            f.write(value.encode("GBK"))

    # 文件尾
    f.write('\x1A'.encode())


if __name__ == '__main__':
    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('开始运行时间：%s' % now1)

    config = iniconfig()
    lastdate = readkey(config,'lastdate')

#    main(sys.argv)
    syldir = getdrive()+'\\syl'

    jyrlst = jyrlist(syldir)
    if lastdate != '' :
        jyrlst = [e for e in jyrlst if e>lastdate]

    if len(jyrlst) >0 :

        lastdate = jyrlst[0]

        with open('pettm.dbf',"rb") as f:
            data = list(dbfreader(f))
            f.close()
        data = data[2:]
        now1 = datetime.datetime.now().strftime('%H:%M:%S')
        print('当前时间：%s' % now1)

        ggpe = []
        for jyrq in jyrlst:

            sylfn = os.path.join(syldir,"csi"+jyrq+".xls")

            print(jyrq)

            ggsj = ggsyl(sylfn,'个股数据',jyrq)

            ggpe.extend(ggsj)

        ggpe.extend(data)

        fieldnames = ['gpdm','gpmc','date','pe_ttm']
        fieldspecs = [('C', 6, 0),('C', 8, 0),('D', 8, 0),('N', 8, 2)]
        with open('pettm.dbf',"wb") as f:
            dbfwriter(f, fieldnames, fieldspecs, ggpe)
            f.close()

        config['lastdate'] = lastdate
        config.write()

    now1 = datetime.datetime.now().strftime('%H:%M:%S')
    print('结束运行时间：%s' % now1)

