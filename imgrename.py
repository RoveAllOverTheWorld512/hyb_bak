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
def imglist(imgdir):
    files = os.listdir(imgdir)
    fs = [re.findall('.*(\d{6})\.png',e) for e in files]
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

def main(argv):

    try:
        opts, args = getopt.getopt(argv[1:], 'hvd:i:o:', ['help','version','date=','inpute=','output='])
    except (getopt.GetoptError):
        Usage()
        sys.exit(1)

    td = datetime.datetime.now().strftime("%Y%m%d") #今天

    sylxls = ""
    jyrq = td
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
        elif o in ('-i', '--input'):
            zxgfile = a
        elif o in ('-o', '--output'):
            sylxls = a
        else:
            print ('无效参数！')
            sys.exit(3)

    syldir = getdrive()+'\\syl'

    if len(zxgfile)==0 :
        zxgfile = "zxg.blk"          #没有指定股票列表就用通达信自选股板块

    if zxgfile.upper().endswith(".BLK") :             #
        tdxblkdir = gettdxblkdir()
        zxgfile = os.path.join(tdxblkdir,zxgfile)
        zxglb = zxglist(zxgfile,"tdxblk")
    else:
        zxglb =  zxglist(zxgfile)

    if len(zxglb)==0 :
        print("股票列表为空,请检查。")
        sys.exit()


    if len(sylxls)== 0:
        zxgwjm = os.path.splitext(os.path.basename(zxgfile))
        sylxls = os.path.join(getpath(),zxgwjm[0]+"_syls.xlsx")

    jyrlst = jyrlist(syldir)


    zxgg0 = []
    hysj0 = []
    for jyrq in jyrlst:

        sylfn = os.path.join(syldir,"csi"+jyrq+".xls")

        print(jyrq)

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

        zxgg0.extend(zxgg)
        hysj0.extend(hysj)


    write_xls(sylfn,sylxls,zxgg0,hysj0)
    print("股票列表文件%s" % zxgfile)
    print("查询结果保存在%s文件中，请查看。" % sylxls)

if __name__ == '__main__':
    main(sys.argv)

