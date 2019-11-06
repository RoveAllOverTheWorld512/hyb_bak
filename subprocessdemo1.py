# -*- coding: utf-8 -*-
"""
Created on Thu Sep  6 10:50:33 2018

@author: lenovo
"""
import os
import subprocess
#cmd_str = "pythonw.exe dlsyl.dy"
#p=os.system(cmd_str)
#print(p)

#ret = os.popen('dir').read()
#print(ret)

cmd_re = subprocess.run("pythonw.exe dlsyl.py", shell=True, stdout=subprocess.PIPE)
print(cmd_re.returncode)
print(cmd_re.stdout.decode('GBK'))

cmd_re = subprocess.call("python.exe dlsyl.py")
print(cmd_re)

ret = os.popen("python.exe dlsyl.py").read()
print(ret)


res = subprocess.run("dira", shell=True, stderr=subprocess.PIPE)
print(res.stderr.decode('gbk'))


ret = subprocess.getoutput('python.exe dlsyl.py')
print(ret)