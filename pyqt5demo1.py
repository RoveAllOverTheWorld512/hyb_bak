#!/usr/bin/python3
# -*- coding: utf-8 -*-

"""
PyQt5 教程

这个例子使用QCalendarWidget控件创建了一个日历。

作者：我的世界你曾经来过
博客：http://blog.csdn.net/weiaitaowang
最后编辑：2016年8月4日
"""

import sys
from PyQt5.QtWidgets import QApplication, QWidget, QCalendarWidget, QLabel
from PyQt5.QtCore import QDate

class Example(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        cal = QCalendarWidget(self)
        cal.setGridVisible(True)
        cal.move(20, 20)
        cal.clicked[QDate].connect(self.showDate)

        self.lb1 = QLabel(self)
        date = cal.selectedDate()
        self.lb1.setText(date.toString())
        self.lb1.move(130, 260)

        self.setGeometry(300, 300, 350, 300)
        self.setWindowTitle('日历控件')        
        self.show()

    def showDate(self, date):

        self.lb1.setText(date.toString())

if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())