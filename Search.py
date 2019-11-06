# -*- coding: utf-8 -*-
"""
Created on Sat Feb  4 20:03:09 2017
Python实现对文件夹内文本文件递归查找
@author: lenovo
"""

import os
def Search(rootDir, suffixes, arg):

    for lists in os.listdir(rootDir):
        path = os.path.join(rootDir, lists)
        if os.path.isfile(path):
            if path.endswith(suffixes):
                try:
                    with open(path, encoding='utf_8') as fh:
                        lineNum = 0
                        for line in fh:
                            lineNum += 1
                            if arg in line:
                                print(lineNum, ':', path, '\n', line)
                except:
                    print('error: ', path, '\n')
        if os.path.isdir(path):
            Search(path, suffixes, arg)


if __name__ == '__main__':
    Search('..', ('.py','.pyw'), 'dbf')