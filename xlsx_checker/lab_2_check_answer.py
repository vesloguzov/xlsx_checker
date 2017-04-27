# -*- coding: UTF-8 -*-
import sys
import random
import datetime
import time

from openpyxl import Workbook, load_workbook
from openpyxl.styles import *


reload(sys)
sys.setdefaultencoding('utf8')

response_message = {}
response_message["ws1"] = {}
response_message["ws1"]["data"] = {}
response_message["ws1"]["graphic"] = {}

response_message["ws2"] = {}
response_message["ws2"]["data"] = {}
response_message["ws2"]["graphic"] = {}


def check_ranges_equal(ws_correct, ws_student, range):
    correct_rows = ws_correct[range]
    correct_list = []
    for row in correct_rows:
        for cell in row:
            try:
                correct_list.append(round(float(cell.value), 6))
            except:
                correct_list.append(cell.value.replace(" ", "").lower().replace(".", ","))

    student_rows = ws_student[range]
    student_list = []
    for row in student_rows:
        for cell in row:
            try:
                student_list.append(round(float(cell.value), 6))
            except:
                student_list.append(cell.value.replace(" ", "").lower().replace(".", ","))

    return correct_list == student_list

def lab_2_check_answer(correct_wb, correct_wb_data_only, student_wb, student_wb_data_only):
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

    return response_message