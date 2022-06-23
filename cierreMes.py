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
    data = pd.read_excel(file_name, tab_name, index_col=None, usecols=col, header=row - 1, nrows=0)
    return data.columns.values[0]


path = ''
file_name = 'python.xlsx'
tab = 'FRP v2.0'
open_xls(path, file_name)
read_cell(tab, "C", 7)

# PYXL
wb = load_workbook(file_name)
sheet = wb[tab]

# sheet['B7'] = '=SUM(C5:C6)'  # xls formulas
# sheet['B7'].style = 'Currency'  # format cells
# sheet['A1'] = 42  # Data can be assigned directly to cells
# sheet.append([1, 2, 3])  # Rows can also be appended
# sheet['A2'] = datetime.datetime.now()  # Python types will automatically be converted
sheet['F2'] = 'Diciembre'
wb.save(file_name)  # Save file

wb = load_workbook('Simulación.xlsx')
sheet = wb['Simulación']
min_column = wb.active.min_column
max_column = wb.active.max_column
min_row = wb.active.min_row
max_row = wb.active.max_row
col = 'A'

pep = 'D-01350.1.1.1'
INGRESO_POR_RECURSO_PROPIO = 0
COSTE_POR_RECURSO_PROPIO = 0
MOB = 0
for row in range(2, max_row):
    cellPEP = col + row.__str__()
    # print(cell)
    if not sheet[cellPEP].value is None and sheet[cellPEP].value == pep:
        cellConcepto = 'AK' + row.__str__()
        cellMOB = 'AJ' + row.__str__()

        if not sheet[cellConcepto].value is None:
            if 'INGRESO POR RECURSO PROPIO' in sheet[cellConcepto].value:
                INGRESO_POR_RECURSO_PROPIO = INGRESO_POR_RECURSO_PROPIO + sheet['AR' + row.__str__()].value
            if 'COSTE POR RECURSO PROPIO' in sheet[cellConcepto].value:
                COSTE_POR_RECURSO_PROPIO = COSTE_POR_RECURSO_PROPIO + sheet['AR' + row.__str__()].value
        if not sheet[cellMOB].value is None:
            if 'MOB-' in sheet[cellMOB].value:
                MOB = MOB + sheet['AR' + row.__str__()].value

print('PEP: ', pep)
print('INGRESO_POR_RECURSO_PROPIO: ', INGRESO_POR_RECURSO_PROPIO)
print('COSTE_POR_RECURSO_PROPIO: ', COSTE_POR_RECURSO_PROPIO)
print('MOB: ', MOB)
