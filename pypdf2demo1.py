# -*- coding: utf-8 -*-
"""
Created on Sun Apr 15 21:27:56 2018
PyPDF2 用 Python 操作 PDF
https://zhuanlan.zhihu.com/p/26647491

Python 深入浅出 - PyPDF2 处理 PDF 文件 
https://blog.csdn.net/xingxtao/article/details/79056341

"""
from PyPDF2 import PdfFileReader, PdfFileWriter
infn = 'H3_AP201804121122335499_1.pdf'
outfn = 'outfn.pdf'
# 获取一个 PdfFileReader 对象
pdf_input = PdfFileReader(open(infn, 'rb'))
# 获取 PDF 的页数
page_count = pdf_input.getNumPages()


# 获取一个 PdfFileWriter 对象
pdf_output = PdfFileWriter()
# 将一个 PageObject 加入到 PdfFileWriter 中

for i in range(page_count):
    # 返回一个 PageObject
    page = pdf_input.getPage(i)
    print(page.extractText().encode("gbk", "ignore"))

    pdf_output.addPage(page)

# 输出到文件中
pdf_output.write(open(outfn, 'wb'))

