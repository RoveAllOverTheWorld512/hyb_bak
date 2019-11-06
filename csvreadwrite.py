# -*- coding: utf-8 -*-
"""
Created on Thu May  4 16:35:44 2017

@author: lenovo
"""

import csv

with open('names.csv', 'w') as csvfile:
    fieldnames = ['first_name', 'last_name']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

    writer.writeheader()
    writer.writerow({'first_name': 'Baked', 'last_name': 'Beans'})
    writer.writerow({'first_name': 'Lovely'})
    writer.writerow({'first_name': 'Wonderful', 'last_name': 'Spam'})
