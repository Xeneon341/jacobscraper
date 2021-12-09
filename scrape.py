import os
import xlrd
import pandas as pd

from pathlib import Path

TEMPLATE_PATH = Path.cwd()

def evaluate_division_code(entry_tuple):
    if entry_tuple[2] == 0.0:
        return "DS"
    else:
        return "DTX"


def evaluate_gl_code(gl_code):
    meals_ent = 'MFOWLER CC EXP - Meals & Ent'
    travel = 'MFOWLER CC EXP - Travel'
    meals_cust = 'MFOWLER CC EXP - Meals w/ Cust'

    if gl_code == 61201:
        return meals_ent
    elif gl_code == 61251:
        return meals_cust
    elif gl_code == 61301:
        return travel
    else:
        pass


def extract_jacob_enter():
    loc = (os.path.join(TEMPLATE_PATH, "mf_config.xls"))

    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0)

    expenses_list = list(zip(sheet.col_values(0, start_rowx=1, end_rowx=None),
                             sheet.col_values(1, start_rowx=1, end_rowx=None),
                             sheet.col_values(2, start_rowx=1, end_rowx=None),
                             sheet.col_values(3, start_rowx=1, end_rowx=None),
                             sheet.col_values(4, start_rowx=1, end_rowx=None),
                             sheet.col_values(5, start_rowx=1, end_rowx=None)))

    return expenses_list


def create_excel_entries(list):
    columns = ['Type',
               'No.',
               'Description/Comment',
               'Freight Code',
               'Office Code',
               'Department Code',
               'Division Code',
               'Project Cost Center',
               'Location Code',
               'Quantity',
               'Unit of Measure Code',
               'Direct Unit Cost Excl. Tax',
               'Tax Area Code',
               'Line Amount Excl. Tax',
               'Line Discount %',
               'Qty. to Assign',
               'Qty. Assigned',
               'Item Charge No.',
               'Transport Order Doc. No.',
               'Transport Container No.',
               'Over Receive',
               'Over Receive Verified'
               ]

    df = pd.DataFrame(columns=columns)

    for i in range(len(list)):
        if list[i][2] == 0.0 or list[i][3] == 0.0:
            new_row = ['G/L Account', list[i][0], evaluate_gl_code(list[i][0]), '', list[i][4], list[i][5], evaluate_division_code(
                list[i]), '', '', 1, '', list[i][1], '', list[i][1], '', '', '', '', '', '', 'No', 'No']
            series_dict = pd.Series(dict(zip(df.columns, new_row)))
            df = df.append(series_dict, ignore_index=True)
        else:
            new_row_dtx = ['G/L Account', list[i][0], evaluate_gl_code(
                list[i][0]), '', list[i][4], list[i][5], "DTX", '', '', 1, '', list[i][1] * list[i][2], '', list[i][1] * list[i][2], '', '', '', '', '', '', 'No', 'No']
            new_row_ds = ['G/L Account', list[i][0], evaluate_gl_code(
                list[i][0]), '', list[i][4], list[i][5], "DS", '', '', 1, '', list[i][1] * list[i][3], '', list[i][1] * list[i][3], '', '', '', '', '', '', 'No', 'No']

            series_dict_dtx = pd.Series(dict(zip(df.columns, new_row_dtx)))
            df = df.append(series_dict_dtx, ignore_index=True)
            
            series_dict_ds = pd.Series(dict(zip(df.columns, new_row_ds)))
            df = df.append(series_dict_ds, ignore_index=True)

    return df


create_excel_entries(extract_jacob_enter()).to_excel(
    "output.xlsx", sheet_name='output')
