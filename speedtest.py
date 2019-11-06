# -*- coding: utf-8 -*-
"""
Created on Wed Jan 24 17:31:34 2018

@author: lenovo
"""

import unittest
from selenium import webdriver
import time
class TestThree(unittest.TestCase):
 
    def setUp(self):
        self.startTime = time.time()
 
    def test_url_fire(self):
        self.driver = webdriver.Firefox()
        self.driver.get("http://www.qq.com")
        self.driver.quit()
 
    def test_url_phantom(self):
        self.driver = webdriver.PhantomJS()
        self.driver.get("http://www.qq.com")
        self.driver.quit()
 
    def tearDown(self):
        t = time.time() - self.startTime
        print("%s: %.3f" % (self.id(), t))
        self.driver.quit
 
if __name__ == '__main__':
    suite = unittest.TestLoader().loadTestsFromTestCase(TestThree)
    unittest.TextTestRunner(verbosity=0).run(suite)