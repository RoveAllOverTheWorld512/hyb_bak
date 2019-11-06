# -*- coding: utf-8 -*-
"""
Created on Sun Dec 24 15:45:28 2017

@author: lenovo
"""

# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'example01.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!


class Ui_Form(object):                                                      #创建窗口类，继承object

    def setupUi(self, Form):

        Form.setObjectName("Form")                                          #设置窗口名

        Form.resize(400, 300)                                               #设置窗口大小

        self.quitButton = QtWidgets.QPushButton(Form)                       #创建一个按钮，并将按钮加入到窗口Form中

        self.quitButton.setGeometry(QtCore.QRect(280, 240, 75, 23))         #设置按钮大小与位置

        self.quitButton.setObjectName("quitButton")                         #设置按钮名


        self.retranslateUi(Form)

        QtCore.QMetaObject.connectSlotsByName(Form)                         #关联信号槽


    def retranslateUi(self, Form):

        _translate = QtCore.QCoreApplication.translate

        Form.setWindowTitle(_translate("Form", "Test"))                     #设置窗口标题

        self.quitButton.setText(_translate("Form", "Quit"))                 #设置按钮显示文字 