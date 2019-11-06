# -*- coding: utf-8 -*-
"""
Created on Sat Mar 31 21:49:06 2018

@author: lenovo
"""
import requests
from contextlib import closing


class ProgressBar(object):

    def __init__(self, title,
                 count=0.0,
                 run_status=None,
                 fin_status=None,
                 total=100.0,
                 unit='', sep='/',
                 chunk_size=1.0):
        super(ProgressBar, self).__init__()
        self.info = "【%s】%s %.2f %s %s %.2f %s"
        self.title = title
        self.total = total
        self.count = count
        self.chunk_size = chunk_size
        self.status = run_status or ""
        self.fin_status = fin_status or " " * len(self.statue)
        self.unit = unit
        self.seq = sep

    def __get_info(self):
        # 【名称】状态 进度 单位 分割线 总数 单位
        _info = self.info % (self.title, self.status,
                             self.count/self.chunk_size, self.unit, self.seq, self.total/self.chunk_size, self.unit)
        return _info

    def refresh(self, count=1, status=None):
        self.count += count
        # if status is not None:
        self.status = status or self.status
        end_str = "\r"
        if self.count >= self.total:
            end_str = '\n'
            self.status = status or self.fin_status
        print(self.__get_info(), end=end_str)


filename = 'd:\\tdx\\shlday_1.zip'
url = 'http://www.tdx.com.cn/products/data/data/vipdoc/shlday.zip'

with closing(requests.get(url, stream=True)) as response:
    chunk_size = 1024 # 单次请求最大值
    content_size = int(response.headers['content-length']) # 内容体总大小
    progress = ProgressBar(file_name(), total=content_size,
                                     unit="KB", chunk_size=chunk_size, run_status="正在下载", fin_status="下载完成")
    with open(file_name, "wb") as file:
       for data in response.iter_content(chunk_size=chunk_size):
           file.write(data)
           progress.refresh(count=len(data))