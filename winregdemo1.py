# -*- coding: utf-8 -*-
"""
Created on Sat Feb 18 14:40:44 2017

@author: Lenovo
"""

import winreg

key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\华西证券华彩人")
value, type = winreg.QueryValueEx(key, "InstallLocation")
print(value)
#获取该键的所有键值，因为没有方法可以获取键值的个数，所以只能用这种方法进行遍历
try:
    i = 0
    while 1:
        #EnumValue方法用来枚举键值，EnumKey用来枚举子键
        name, value, type = winreg.EnumValue(key, i)
        print(name,value,type)
        i += 1
except WindowsError:
    print

      #如果知道键的名称，也可以直接取值
    value, type = winreg.QueryValueEx(key, "InstallLocation")