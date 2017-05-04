# -*- coding: UTF-8 -*-
import sys
import random
import datetime
import time

from openpyxl import Workbook, load_workbook
from openpyxl.styles import *
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.reader import reader
from zipfile import ZipFile, ZIP_DEFLATED
from lxml.etree import fromstring, tostring
from openpyxl.chart import (
    ScatterChart,
    Reference,
    Series,
)

from labs_utils import check_ranges_equal

reload(sys)
sys.setdefaultencoding('utf8')

response_message = {}
response_message["ws1"] = {}
response_message["ws1"]["data"] = {}
response_message["ws1"]["graphic"] = {}

response_message["ws2"] = {}
response_message["ws2"]["data"] = {}
response_message["ws2"]["graphic"] = {}

response_message["errors"] = []

def check_scatter_graphic(filename, data_x, data_y, num):
    analyze = {}
    analyze["errors"] = []
    analyze["chart_title"] = {}
    analyze["data_x"] = {}
    analyze["data_y"] = {}
    analyze["title_y"] = {}

    obj = {}

    sourceFile = ZipFile(filename, 'r')
    charts = []; [charts.append(sourceFile.read(ch)) for ch in sourceFile.namelist() if 'charts/chart' in ch]
    charts_objects = []
    for chart in charts:
        try:
            clean_chart = fromstring(chart)

            for bad in clean_chart.xpath(".//c:extLst", namespaces={'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart'}):
                bad.getparent().remove(bad)
            clean_chart = tostring(clean_chart)
            l = reader(clean_chart)
            charts_objects.append(l)
        except:
            analyze["errors"].append('График не обнаружен!')
            return analyze
    print len(charts_objects)

    try:
        obj = charts_objects[num]
    except:
            analyze["errors"].append('График не обнаружен!')
            return analyze

    if len(charts_objects) != 2:
        response_message["errors"].append('Документ должен содержать два графика')
        return analyze
    

    if obj.tagname != 'scatterChart':
        analyze["errors"].append('График не обнаружен!')
        return analyze

    if obj == {}:
        analyze["errors"].append('График не обнаружен!')
    else:
        if obj.title != None:
            analyze["chart_title"]["message"] = 'Имя графика присвоено'
            analyze["chart_title"]["status"] = True
        else:
            analyze["chart_title"]["message"] = 'Имя графика не присвоено'
            analyze["chart_title"]["status"] = False
        try:
            for s in obj.ser:
                if data_x in s.xVal.numRef.f.replace(" ", ""):
                    analyze["data_x"]["message"] = 'Данные для оси x выбраны верно'
                    analyze["data_x"]["status"] = True
                else:
                    analyze["data_x"]["message"] = 'Данные для оси x выбраны неверно'
                    analyze["data_x"]["status"] = False
        except:
            analyze["data_x"]["message"] = 'Данные для оси x выбраны неверно'
            analyze["data_x"]["status"] = False

        try:
            for s in obj.ser:
                if data_y in s.yVal.numRef.f.replace(" ", ""):
                    analyze["data_y"]["message"] = 'Данные для оси y выбраны верно'
                    analyze["data_y"]["status"] = True
                else:
                    analyze["data_y"]["message"] = 'Данные для оси y выбраны неверно'
                    analyze["data_y"]["status"] = False
        except:
            analyze["data_y"]["message"] = 'Данные для оси y выбраны неверно'
            analyze["data_y"]["status"] = False

        try:
            if obj.y_axis.title != None:
                analyze["title_y"]["message"] = 'Подпись осей выполнена'
                analyze["title_y"]["status"] = True
            else:
                analyze["title_y"]["message"] = 'Подпись осей не выполнена'
                analyze["title_y"]["status"] = False
        except:
            analyze["title_y"]["message"] = 'Подпись осей не выполнена'
            analyze["title_y"]["status"] = False

    return  analyze

def lab_2_check_answer(correct_wb, correct_wb_data_only, student_wb, student_wb_data_only, filename):
    response_message["ws1"] = {}
    response_message["ws1"]["data"] = {}
    response_message["ws1"]["graphic"] = {}

    response_message["ws2"] = {}
    response_message["ws2"]["data"] = {}
    response_message["ws2"]["graphic"] = {}

    response_message["errors"] = []

    if (len(student_wb.get_sheet_names()) == 2):
        student_ws_1 = student_wb[student_wb.get_sheet_names()[0]]
        student_ws_2 = student_wb[student_wb.get_sheet_names()[1]]
        correct_ws_1 = correct_wb[correct_wb.get_sheet_names()[0]]
        correct_ws_2 = correct_wb[correct_wb.get_sheet_names()[1]]

        ws1_data_range = 'A4:B28'

        if check_ranges_equal(correct_ws_1, student_ws_1, ws1_data_range):
            response_message["ws1"]["data"]["status"] = True
            response_message["ws1"]["data"]["message"] = "Данные для графика посчитаны верно"
        else:
            response_message["ws1"]["data"]["status"] = False
            response_message["ws1"]["data"]["message"] = "Данные для графика посчитаны неверно"

        ws2_data_range = 'A4:B34'

        if check_ranges_equal(correct_ws_2, student_ws_2, ws2_data_range):
            response_message["ws2"]["data"]["status"] = True
            response_message["ws2"]["data"]["message"] = "Данные для графика посчитаны верно"
        else:
            response_message["ws2"]["data"]["status"] = False
            response_message["ws2"]["data"]["message"] = "Данные для графика посчитаны неверно"

        data_x_1 = '$A$4:$A$28'
        data_y_1 = '$B$4:$B$28'
        response_message["ws1"]["graphic"] = check_scatter_graphic(filename, data_x_1, data_y_1, 0)

        data_x_2 = '$A$4:$A$34'
        data_y_2 = '$B$4:$B$34'
        response_message["ws2"]["graphic"] = check_scatter_graphic(filename, data_x_2, data_y_2, 1)

    else:
        response_message["errors"].append('Документ должен содержать два рабочих листа')

    return response_message