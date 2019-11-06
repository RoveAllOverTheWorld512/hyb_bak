# -*- coding: utf-8 -*-
"""
Created on Sun Feb 19 12:07:37 2017

@author: Lenovo
"""
import os
import sys

fn = sys.argv[0]
inifile = os.path.splitext(sys.argv[0])[0]+'.ini'
print(inifile)
