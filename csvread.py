# -*- coding: utf-8 -*-
"""
Created on Sat May 13 20:17:21 2017

@author: lenovo
"""

import csv

with open('tblfmt.csv') as f:
    datareader = csv.reader(f);
    print (list(datareader))