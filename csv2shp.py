#!/usr/bin/env python
# -*- coding:utf-8 -*- #

"""
@version: 0.0.1
@author: PuQiu 
@license: Apache Licence 
@contact: puqiuster@gmail.com
@software: PyCharm
@file: csv2shp.py
@time: 2017/6/8 1:26
"""

from osgeo import gdal
from osgeo import ogr
import sys
import csv
import codecs
import os
from XlsReader import XlsReader


class Csv2Shape:
    def __init__(self, ):
        self._pointDic = {}
        self._pointField = []
        self._lineField = []
        self._lineList = []
        self._ProjctName = 'Temp'
        self._codeFieldName = '点号*'.decode('utf-8')  #'OLDNO'  # '\xef\xbb\xbf\xe7\x82\xb9\xe5\x8f\xb7'  # 点号
        self._codeFirstName = '管线起点点'.decode('utf-8')  # 'START_SID'  # '\xef\xbb\xbf\xe8\xb5\xb7\xe7\x82\xb9\xe7\x82\xb9\xe5\x8f\xb7'  # 起点点号
        self._codeSecondName = '管线终点点'.decode('utf-8')  # 'END_SID'  # '\xe7\xbb\x88\xe7\x82\xb9\xe7\x82\xb9\xe5\x8f\xb7'  # 终点点号
        self._X = 'X坐标*'.decode('utf-8')
        self._Y = 'Y坐标*'.decode('utf-8')

    def readCsv(self, pointfile, linefile, prjname='Temp'):

        if pointfile.endswith('.csv') and linefile.endswith('.csv'):
            uptPath = unicode(pointfile, "utf8")
            with open(uptPath, 'r') as f:
                reader = csv.DictReader(f)
                self._pointField = reader.fieldnames
                try:
                    for row in reader:
                        self._pointDic[row[self._codeFieldName]] = row
                except Exception as e:
                    print(e)
            ulinePath = unicode(linefile, "utf8")
            with open(ulinePath, 'r') as f:
                reader = csv.DictReader(f)
                self._lineField = reader.fieldnames
                try:
                    for row in reader:
                        if row[self._codeFirstName] != '' and row[self._codeSecondName] != '':
                            self._lineList.append(row)
                except Exception as e:
                    print(e)
        self._ProjctName = prjname

    def readXls(self,pointfile, linefile, prjname='Temp'):
        if pointfile.endswith('.xls') and os.path.exists(pointfile.decode('utf-8').encode('gbk')):
            xlsr = XlsReader(pointfile)
            data = xlsr.get_data()
            self._pointField = xlsr.header
            for row in data:
                try:
                    self._pointDic[row[self._codeFieldName]] = row
                except Exception,e:
                    print pointfile

        if linefile.endswith('.xls') and os.path.exists(linefile.decode('utf-8').encode('gbk')):
            xlsr = XlsReader(linefile)
            data = xlsr.get_data()
            self._lineField = xlsr.header
            try:
                for row in data:
                    if row[self._codeFirstName] != '' and row[self._codeSecondName] != '':
                        self._lineList.append(row)
            except Exception, e:
                print pointfile
                print e

        splitpath = os.path.split(pointfile)
        strbase = splitpath[1]
        strdir = splitpath[0]
        ustrbase = strbase
        ustrdir = strdir
        if strbase.find('点表') != -1 and strbase.find('.xls'):
            strbase = strbase.replace('点表', '')
            ustrbase = strbase.replace('.xls', '')
        self._ProjctName = ustrdir + '//result//' + ustrbase
        # print self._ProjctName

    def createPointsShapeFile(self):
        if not self._pointDic or not self._lineList:
            print '读入数据为空'
            return

        gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "Yes")
        gdal.SetConfigOption('SHAPE_ENCODING', 'CP936')

        driverName = "ESRI Shapefile"
        drv = gdal.GetDriverByName(driverName)

        if drv is None:
            print '获取格式驱动失败'
            return
        ptfilePath = self._ProjctName + 'pt' + '.shp'
        ds = drv.Create(ptfilePath, 0, 0, 0, gdal.GDT_Unknown)

        if ds is None:
            print "Creation of output file failed.\n"
            return
        lyr = ds.CreateLayer("point_out", None, ogr.wkbPoint)
        if lyr is None:
            print "Layer creation failed.\n"
            return
        try:
            for item in self._pointField:
                # if item[:3] == codecs.BOM_UTF8:
                #     item = item[3:]
                field_defn = ogr.FieldDefn(item.encode('utf-8'), ogr.OFTString)
                field_defn.SetWidth(32)
                if lyr.CreateField(field_defn) != 0:
                    print "Layer creation failed.\n"
                    return
            # print self._ProjctName
            for (key, value) in self._pointDic.items():
                if not key.strip() or not value:
                    continue
                pt = ogr.Geometry(ogr.wkbPoint)
                pt.SetPoint_2D(0, float(value[self._X]), float(value[self._Y]))
                feature = ogr.Feature(lyr.GetLayerDefn())
                feature.SetGeometry(pt)

                for (k, v) in value.items():
                    feature.SetField(k.encode('utf-8'), v)

                lyr.CreateFeature(feature)
                pt.Destroy()
                feature.Destroy()
        except Exception, e:
            print e
            print self._ProjctName

    def createLineShapeFile(self):

        if not self._pointDic or not self._lineList:
            print '读入数据为空'
            return

        gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "Yes")
        gdal.SetConfigOption('SHAPE_ENCODING', 'CP936')

        driverName = "ESRI Shapefile"
        drv = gdal.GetDriverByName(driverName)

        if drv is None:
            print '获取格式驱动失败'
            sys.exit(1)
        ptfilePath = self._ProjctName + 'line' + '.shp'
        ds = drv.Create(ptfilePath, 0, 0, 0, gdal.GDT_Unknown)

        if ds is None:
            print "Creation of output file failed.\n"
            return
        lyr = ds.CreateLayer("line_out", None, ogr.wkbLineString)
        if lyr is None:
            print "Layer creation failed.\n"
            return

        try:
            for item in self._lineField:
                field_defn = ogr.FieldDefn(item.encode('utf-8'), ogr.OFTString)
                field_defn.SetWidth(32)
                if lyr.CreateField(field_defn) != 0:
                    print "Layer creation failed.\n"
                    return

            for item in self._lineList:

                firestNode = item[self._codeFirstName]
                secondNode = item[self._codeSecondName]

                if not self._pointDic.has_key(firestNode) or not self._pointDic.has_key(secondNode):
                    continue

                fx = float(self._pointDic[firestNode][self._X])
                fy = float(self._pointDic[firestNode][self._Y])
                sx = float(self._pointDic[secondNode][self._X])
                sy = float(self._pointDic[secondNode][self._Y])

                line = ogr.Geometry(ogr.wkbLineString)
                line.AddPoint(fx, fy)
                line.AddPoint(sx, sy)
                feature = ogr.Feature(lyr.GetLayerDefn())
                feature.SetGeometry(line)
                for (k, v) in item.items():
                    feature.SetField(k.encode('utf-8'), v)

                lyr.CreateFeature(feature)
                line.Destroy()
                feature.Destroy()
        except Exception, e:
            print e
            print self._ProjctName

    @staticmethod
    def rb(str):
        if str[:3] == codecs.BOM_UTF8:
            return str[3:]
        else:
            return str


class Batch2Shp:
    def __init__(self, dir):
        if not os.path.exists(dir):
            print "批量文件夹不存在"
            sys.exit(1)
        self._dirpath = dir

    def walk(self):
        for path in os.listdir(self._dirpath):
            absolutely_path = os.path.join(self._dirpath, path)
            if os.path.isdir(absolutely_path) and os.path.exists(absolutely_path):
                if len(os.listdir(absolutely_path)) >= 2:
                    if os.path.exists(absolutely_path + '//pt.csv') and os.path.exists(absolutely_path + '//line.csv'):
                        upath = absolutely_path.decode('cp936').encode('utf')
                        csv2shp = Csv2Shape()
                        csv2shp.readCsv(upath + '//pt.csv', upath + '//line.csv', upath + '//')
                        csv2shp.createPointsShapeFile()
                        csv2shp.createLineShapeFile()

    def walkXls(self):
        ptLine = {}
        for path in os.listdir(self._dirpath):
            uPath = path.decode('gbk').encode('utf-8')
            if uPath.find('点表') != -1 and uPath.endswith('xls'):
                absolutely_pt_path = os.path.join(self._dirpath, uPath)
                absolutely_line_path = os.path.join(self._dirpath, uPath.replace('点表','线表'))
                # print absolutely_line_path == absolutely_pt_path
                if os.path.exists(absolutely_pt_path.decode('utf-8').encode('gbk')) and os.path.exists(absolutely_line_path.decode('utf-8').encode('gbk')):
                    ptLine[absolutely_pt_path] = absolutely_line_path
                    # print absolutely_pt_path
                    # print absolutely_line_path
                else:
                    print absolutely_pt_path + "未读取到文件"
        return ptLine


if "__main__" == __name__:
    bach = Batch2Shp('D://data_process//bacth')
    dictP = bach.walkXls()
    for (key, value) in dictP.items():
        toShp = Csv2Shape()
        toShp.readXls(key, value)
        toShp.createPointsShapeFile()
        toShp.createLineShapeFile()
