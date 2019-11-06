#coding:utf-8

from __future__ import division

import sqlite3
import xlrd
import os,sys

#读取文件,返回文件名，省份，时间,文件所对应的类别
def getFileList(dir,wildcard,recursion):
    os.chdir(dir)

    fileList = []
    check_province = []
    check_time = []
    file_type = []

    exts = wildcard.split(" ")
    files = os.listdir(dir)
    for name in files:
        fullname=os.path.join(dir,name)
        if(os.path.isdir(fullname) & recursion):
            getFileList(fullname,wildcard,recursion)
        else:
            for ext in exts:
                if(name.endswith(ext)):
                    fileList.append(name)
                    check_province.append(name.split('-')[1])
                    check_time.append(name.split('-')[0])
                    file_type.append(name.split('-')[2].split('_')[1])
    return fileList,check_time,check_province,file_type

#建立数据库
def createDataBase():
    cn = sqlite3.connect('check.db')

    cn.execute('''CREATE TABLE IF NOT EXISTS TB_CHECK
           (ID integer PRIMARY KEY AUTOINCREMENT,
           ITEM TEXT,
           FIELD TEXT,
           TYPE TEXT,
           CONTENT TEXT,
           ATTRIBUTE TEXT,
           CHECKPOINT TEXT,
           REMARKS TEXT,
           ANSWER TEXT,
           DESCRIPTION TEXT,
           SUGGESTION TEXT,
           PROVINCE  TEXT,
           TIME TEXT,
           STYLE TEXT);''')
    #总分表
    cn.execute('''CREATE TABLE IF NOT EXISTS TB_SCORE
           (ID integer PRIMARY KEY AUTOINCREMENT,
           PROVINCE  TEXT,
           TIME TEXT,
           FILETYPE TEXT,
           SCORE INTEGER);''')


#存储数据
def readExcel(filename,cn,check_province,check_time,FileType):
  #读取
  workbook = xlrd.open_workbook(filename)
  # 获取sheet
  sheet_name = workbook.sheet_names()[0]
  sheet = workbook.sheet_by_name(sheet_name)

  check_Item = 'a'

  # 条目的总数
  itemCount = 0
  score = 0

  second = sheet.cell(7,1).value.encode('utf-8')

  for i in range(7,sheet.nrows):
      if sheet.cell(i,1).value.encode('utf-8') == second:
          check_Item = sheet.cell(i,0).value.encode('utf-8')
          continue

      temp = []
      for j in range(0,sheet.ncols):
          temp.append(sheet.cell(i,j).value.encode('utf-8'))

      answer = sheet.cell(i,7).value.encode('utf-8')

      # 判断分数
      if answer == "yes" or answer == "no":
          score = score + 1

      # 判断答案的输入是否正确
      if answer == "other":
          print "!!!Failed to import'%s'" % (filename)
          print "!!!Please Choose an Right Answer for '%s'--------"%(filename)
          break
      else:
          cn.execute("insert into TB_CHECK (ITEM,FIELD,TYPE,CONTENT,"
                     "ATTRIBUTE,CHECKPOINT,REMARKS,ANSWER,DESCRIPTION,"
                     "SUGGESTION,PROVINCE,TIME,STYLE) "
                     "values('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')"
                     ""%(temp[0],temp[1],temp[2],temp[3],temp[4],temp[5],temp[6],temp[7],temp[8],temp[9],check_province,check_time,check_Item))

          itemCount = itemCount + 1
  if itemCount != 0:
      score = round(score * (100 / itemCount), 2)

      cn.execute("insert into TB_SCORE (PROVINCE,TIME,FILETYPE,SCORE) "
             "values('%s','%s','%s','%.2f')"%(check_province,check_time,FileType,score))
      print "Successful for'%s'--------" % (filename)
  cn.commit()

#转码函数
def changeCode(name):
    name = name.decode('GBK')
    name = name.encode('UTF-8')
    return name

#导入数据
def importData(path):
    # 数据库
    createDataBase()
    database = sqlite3.connect("check.db")

    #文件类型
    wildcard = ".xls"

    list = getFileList(path,wildcard,1)

    nfiles = len(list[0])
    #文件名
    file = list[0]
    #时间
    time = list[1]
    #省份
    province = list[2]
    # #文件类型
    FileType = list[3]

    for count in range(0,nfiles):
        filename = file[count]
        check_province = changeCode(province[count])
        check_time = time[count]
        File_type = changeCode(FileType[count])
        readExcel(filename,database,check_province,check_time,File_type)

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print "Wrong Parameters"
    else:
        path = sys.argv[1]
        importData(path)


