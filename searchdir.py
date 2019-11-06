# -*- coding: utf-8 -*-
"""
Created on Sat Feb  4 14:12:43 2017

@author: lenovo
"""

import os  
  
def search(path=".", name="1"):  
    for item in os.listdir(path):  
        item_path = os.path.join(path, item)  
        if os.path.isdir(item_path):  
            search(item_path, name)  
        elif os.path.isfile(item_path):  
            if name in item:  
                print(item_path)  
  
if __name__ == "__main__":  
#    search(path=r"D:\hyb",name="xls") 
    search(r"D:\hyb",'.ini')   