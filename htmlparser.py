# -*- coding: utf-8 -*-
"""
Created on Mon Feb  6 14:30:54 2017
http://stackoverflow.com/questions/6325216/parse-html-table-to-python-list
@author: lenovo
"""
from html.parser import HTMLParser


class HTMLTableParser(HTMLParser):
    """ This class serves as a html table parser. It is able to parse multiple
    tables which you feed in. You can access the result per .tables field.
    这类作为一个HTML表分析器。它能够解析你传入的多个表。你能访问结果的每个.tables字段
    """
    def __init__(self, decode_html_entities=False, data_separator=' ', ):

        HTMLParser.__init__(self)

        self._parse_html_entities = decode_html_entities
        self._data_separator = data_separator

        self._in_td = False
        self._in_th = False
        self._current_table = []
        self._current_row = []
        self._current_cell = []
        self.tables = []

    def handle_starttag(self, tag, attrs):
        """ We need to remember the opening point for the content of interest.
        The other tags (<table>, <tr>) are only handled at the closing point.
        我们需要记住感兴趣内容的开始点。其它标签(<table>, <tr>)是仅在关闭点处理。
        """
        if tag == 'td':
            self._in_td = True
        if tag == 'th':
            self._in_th = True

    def handle_data(self, data):
        """ This is where we save content to a cell 
        在这里保存内容
        """
        if self._in_td or self._in_th:
            self._current_cell.append(data.strip())

    def handle_charref(self, name):
        """ Handle HTML encoded characters 
        处理HTML编码字符
        """

        if self._parse_html_entities:
            self.handle_data(self.unescape('&#{};'.format(name)))

    def handle_endtag(self, tag):
        """ Here we exit the tags. If the closing tag is </tr>, we know that we
        can save our currently parsed cells to the current table as a row and
        prepare for a new row. If the closing tag is </table>, we save the
        current table and prepare for a new one.
        这里是退出标签。如果
        """
        if tag == 'td':
            self._in_td = False
        elif tag == 'th':
            self._in_th = False

        if tag in ['td', 'th']:
            final_cell = self._data_separator.join(self._current_cell).strip()
            self._current_row.append(final_cell)
            self._current_cell = []
        elif tag == 'tr':
            self._current_table.append(self._current_row)
            self._current_row = []
        elif tag == 'table':
            self.tables.append(self._current_table)
            self._current_table = []

def int2str(n):
    if isinstance(n,float):
        return str(int(n))
    if isinstance(n,int):
        return str(n)
    return n

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

########################################################################
#检测是不是可以转换成浮点数
########################################################################
def str2float(num):
    try:
        return float(num)
    except ValueError:
        return num

xhtml=open(r'd:\hyb\yjyg.html','rb').read().decode("UTF")
            
p = HTMLTableParser()
p.feed(xhtml)
#print(p.tables[0][0])
#print(p.tables[0][1])
zxg = ["002294","002002"]
data0 = p.tables[0][1:]
data1 = []
for da in data0 :
    if da[0][:6] in zxg:
        da[0] = da[0][:6]
        da[2] = str2float(da[2])
        da[3] = str2float(da[3])
        da[5] = str2float(da[5])
        da[6] = str2float(da[6])
        data1.append(da)

data1.insert(0,p.tables[0][0])

print(data1)
        
        
