# -*- coding: utf-8 -*-
"""
Created on Fri Feb 24 17:07:45 2017

@author: Lenovo
"""

import os
import sys
import struct
import pandas as pd
from configobj import ConfigObj
import winreg

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

def getdisk():
    return sys.argv[0][:2]

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

def filename(pathname):
    wjm = os.path.splitext(os.path.basename(pathname))
    return wjm[0]

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
        stk = []
        for i in range(blknum) :
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
        da.append([blk[key][0],blk[key][1],blk[key][2]])

    df = pd.DataFrame(da,columns=['blockname','shortname','num'])
    csvfn = 'block_'+lb+'.csv'
    df.to_csv(csvfn,encoding='GBK',sep=',',index=False)

if __name__ == '__main__':
    for lb in ['gn','fg','zs'] :
        blk = gettdxblk(lb)





