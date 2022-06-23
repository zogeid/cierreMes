import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference
import string
import datetime

def open_xls(p, f):
    return pd.read_excel(p + f)


def read_cell(tab_name, col, row):
    data = pd.read_excel(file_name, tab_name , index_col=None, usecols=col, header=row-1, nrows=0)
    return data.columns.values[0]


path = ''
file_name = 'python.xlsx'
tab = 'FRP v2.0'
open_xls(path, file_name)
read_cell(tab, "C", 7)

#PYXL
wb = load_workbook(file_name)
sheet = wb[tab]
min_column = wb.active.min_column
max_column = wb.active.max_column
min_row = wb.active.min_row
max_row = wb.active.max_row
# sheet['B7'] = '=SUM(C5:C6)'  # xls formulas
# sheet['B7'].style = 'Currency'  # format cells
# sheet['A1'] = 42  # Data can be assigned directly to cells
# sheet.append([1, 2, 3])  # Rows can also be appended
# sheet['A2'] = datetime.datetime.now()  # Python types will automatically be converted
sheet['F2'] = 'Diciembre'
wb.save(file_name)  # Save file


wb = load_workbook('Simulación.xlsx')
sheet = wb['Simulación']
if not sheet['A2'].value == None and sheet['A2'].value == 'D-01350.1.1.1':
    print('OK')
sheet['F2'] = 'Diciembre'
wb.save(file_name)  # Save file