#!/usr/bin/env python
# -*- coding:utf-8 -*- #

"""
@version: 0.0.1
@author: PuQiu 
@license: Apache Licence 
@contact: puqiuster@gmail.com
@software: PyCharm
@file: xlsReader.py
@time: 2017/6/9 23:20
"""
import xlrd


class XlsReader:

    def __init__(self, path):
        self._path = path
        try:
            uptPath = unicode(path, "utf8")
            self._data = xlrd.open_workbook(uptPath)
            self._table = self._data.sheets()[0]
        except Exception, e:
            print e

    @property
    def header(self):
        if self._data is None:
            return None
        try:
            table = self._data.sheets()[0]
            rowvalue = table.row_values(0)

            # 截断字段名长度为小于10，ShapeFile字段名长度限制
            heads = []
            for item in rowvalue:
                if len(item)>5:
                    tmp = item[:5]
                    heads.append(tmp)
                else:
                    heads.append(item)
            return heads
        except Exception, e:
            print self._path
            print e


    @property
    def count(self):
        return self._table.nrows

    def get_data(self):
        listrows=[]
        for rownum in range(1, self.count):
            row = self._table.row_values(rownum)
            if row:
                dictrow = {}
                for i in range(len(self.header)):
                    dictrow[self.header[i]] = row[i]
                listrows.append(dictrow)
        return listrows

if __name__ == "__main__":
    xlsr = XlsReader('D://data_process//bacth//刘雪冬(RMS00262016010300)点表.xls')
    print xlsr.header[1]
    print xlsr.count
    print xlsr.get_data()[0].keys()[0]
