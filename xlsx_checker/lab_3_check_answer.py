# -*- coding: UTF-8 -*-
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.reader import reader
from openpyxl.chart.layout import Layout, ManualLayout
from labs_utils import range_is_date_format, range_is_money_rub_format, formulas_is_equal

reload(sys)
sys.setdefaultencoding('utf8')

response_message = {}
response_message["formats"] = {}
response_message["errors"] = []
response_message["results"] = {}
response_message["filters"] = {}

def check_ranges_equal(ws_correct, ws_student, range):
    correct_rows = ws_correct[range]
    correct_list = []
    for row in correct_rows:
        for cell in row:
            try:
                correct_list.append(round(float(cell.value), 2))
            except:
                correct_list.append(cell.value)

    student_rows = ws_student[range]
    student_list = []
    for row in student_rows:
        for cell in row:
            try:
                student_list.append(round(float(cell.value), 2))
            except:
                student_list.append(cell.value)

    return correct_list == student_list

def check_cost(ws_student, student_range):
    student_rows = ws_student[student_range]
    for index, row in enumerate(student_rows):
        for cell in row:
            var1 = '=E'+str(index+3)+'*D'+str(index+3)
            var2 = '=D' + str(index + 3) + '*E' + str(index + 3)
            if formulas_is_equal(cell.value, var1) == False and  formulas_is_equal(cell.value, var2) == False:
                return False
    return True

def check_results(correct_ws, student_ws):
    results_range = 'A2:E26'
    check_vals = check_ranges_equal(correct_ws, student_ws, results_range)
    check_rows = [4,6,8, 11, 15, 19, 21, 25, 26]
    for r in check_rows:
        if check_ranges_equal(correct_ws, student_ws, 'A'+str(r)+':F'+str(r)) == False:
            return False
    if check_vals:
        return True
    return False

def check_sorting(correct_ws, student_ws):
    range = 'A2:E17'
    return check_ranges_equal(correct_ws, student_ws, range)

def check_formats(student_ws):
    # Проверяем правильность форматирования дат поступления
    response_message["formats"]["dates"] = {}
    if range_is_date_format(student_ws, 'C3:C17')['status']:
        response_message["formats"]["dates"]["status"] = True
        response_message["formats"]["dates"]["message"] = "Форматирование дат поступления верно"
    else:
        response_message["formats"]["dates"]["status"] = False
        response_message["formats"]["dates"]["message"] = "Форматирование дат поступления неверно"

    # Проверяем правильность форматирования цены
    response_message["formats"]["price"] = {}
    if range_is_money_rub_format(student_ws, 'E3:E17')['status']:
        response_message["formats"]["price"]["status"] = True
        response_message["formats"]["price"]["message"] = "Форматирование цены верно"
    else:
        response_message["formats"]["price"]["status"] = False
        response_message["formats"]["price"]["message"] = "Форматирование цены неверно"

    # Проверяем правильность форматирования стоимости
    response_message["formats"]["cost"] = {}
    if range_is_money_rub_format(student_ws, 'F3:F17')['status']:
        response_message["formats"]["cost"]["status"] = True
        response_message["formats"]["cost"]["message"] = "Форматирование стоимости верно"
    else:
        response_message["formats"]["cost"]["status"] = False
        response_message["formats"]["cost"]["message"] = "Форматирование стоимости неверно"

def get_date_custom_filters(ws):
    filters = {}
    filters['year'] = {}
    filters['year']['type'] = ''
    filters['year']['column'] = ''

    filters['custom'] = {}
    filters['custom']['column'] = ''
    filters['custom']['greaterThan'] = ''
    filters['custom']['lessThan'] = ''
    filters['range'] = ws.auto_filter.ref

    for Colfilter in ws.auto_filter.filterColumn:

        if Colfilter.dynamicFilter is not None:
            filters['year']['column'] = float(Colfilter.colId)
            filters['year']['type'] = Colfilter.dynamicFilter.type

        if Colfilter.customFilters is not None:
            for Colfilter1 in Colfilter.customFilters.customFilter:
                filters['custom']['column'] = float(Colfilter.colId)
                if Colfilter1.operator == 'greaterThan':
                    filters['custom']['greaterThan'] = float(Colfilter1.val)
                if Colfilter1.operator == 'lessThan':
                    filters['custom']['lessThan'] = float(Colfilter1.val)
    return filters

def check_filters(correct_ws, student_ws):

    is_data = check_ranges_equal(correct_ws, student_ws, 'D2:E17')
    response_message["filters"]["errors"] = []
    if is_data:
        response_message["filters"]["errors"].append("Таблица для фильтрации неверна")

    response_message["filters"]["year"] = {}
    if get_date_custom_filters(correct_ws)['year'] == get_date_custom_filters(student_ws)['year']:
        response_message["filters"]["year"]["message"] = "Фильтр по текущему году применен верно"
        response_message["filters"]["year"]["status"] = True
    else:
        response_message["filters"]["year"]["message"] = "Фильтр по текущему году применен неверно"
        response_message["filters"]["year"]["status"] = False

    response_message["filters"]["custom"] = {}
    if get_date_custom_filters(correct_ws)['custom'] == get_date_custom_filters(student_ws)['custom']:
        response_message["filters"]["custom"]["message"] = "Фильтр по цене применен верно"
        response_message["filters"]["custom"]["status"] = True
    else:
        response_message["filters"]["custom"]["message"] = "Фильтр по цене применен неверно"
        response_message["filters"]["custom"]["status"] = False


def lab_3_check_answer(correct_wb, correct_wb_data_only, student_wb, student_wb_data_only):
    data = [
        ["Комбайн", "19.07.2017", 100, 7800.00],
        ["Миксер", "30.05.2017", 38, 3000.00],
        ["Микровоновка", "23.08.2017", 38, 4500.00],
        ["Пылесос", "17.03.2017", 25, 3000.00],
        ["Холодильник", "03.05.2016", 56, 25000.00],
        ["Пылесос", "03.08.2017", 6, 1500.00],
        ["Телевизор", "02.03.2014", 50, 6000.00],
        ["Телевизор", "16.02.2016", 19, 12000.00],
        ["Телевизор", "13.09.2017", 32, 4500.00],
        ["Утюг", "12.07.2016", 70, 2000.00],
        ["Утюг", "20.08.2016", 15, 1000.00],
        ["Утюг", "02.08.2017", 20, 2900.00],
        ["Чайник", "15.03.2017", 25, 1540.00],
        ["Чайник", "27.07.2016", 102, 1200.00],
        ["Чайник", "04.08.2016", 45, 500.00],
    ]
    response_message["errors"] = []
    
    if (len(student_wb.get_sheet_names()) == 3):
        student_ws_1 = student_wb[student_wb.get_sheet_names()[0]]
        student_ws_2 = student_wb[student_wb.get_sheet_names()[1]]
        student_ws_3 = student_wb[student_wb.get_sheet_names()[2]]
        correct_ws_1 = correct_wb[correct_wb.get_sheet_names()[0]]
        correct_ws_2 = correct_wb[correct_wb.get_sheet_names()[1]]
        correct_ws_3 = correct_wb[correct_wb.get_sheet_names()[2]]
        # student_ws_read_only_1 = student_wb_data_only[student_wb_data_only.get_sheet_names()[0]]
        # student_ws_read_only_2 = student_wb_data_only[student_wb_data_only.get_sheet_names()[1]]
        # student_ws_read_only_3 = student_wb_data_only[student_wb_data_only.get_sheet_names()[2]]
        # correct_ws_read_only_1 = correct_wb_data_only[correct_wb_data_only.get_sheet_names()[0]]
        # correct_ws_read_only_2 = correct_wb_data_only[correct_wb_data_only.get_sheet_names()[1]]
        # correct_ws_read_only_3 = correct_wb_data_only[correct_wb_data_only.get_sheet_names()[2]]
        cost_range = 'F3:F17'

        if check_cost(student_ws_1, cost_range):
            check_formats(student_ws_1)

            response_message["sort"] = {}
            if check_sorting(correct_ws_1, student_ws_1):
                response_message["sort"]["status"] = True
                response_message["sort"]["message"] = "Сортировка товара выполнена верно"
            else:
                response_message["sort"]["status"] = False
                response_message["sort"]["message"] = "Сортировка товара выполнена неверно"

            response_message["results"] = {}
            if check_results(correct_ws_2, student_ws_2):
                response_message["results"]["status"] = True
                response_message["results"]["message"] = "Лист итогов выполнен верно"
            else:
                response_message["results"]["status"] = False
                response_message["results"]["message"] = "Лист итогов выполнен неверно"

            check_filters(correct_ws_3, student_ws_3)


        else:
            response_message["errors"].append('Стоимость не заполнена')

    else:
         response_message["errors"].append('Документ должен содержать три рабочих листа')
    return response_message