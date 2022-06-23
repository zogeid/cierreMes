import glob
import os
import datetime
from openpyxl import load_workbook


class Pep:
    def __init__(self, pep, ingreso, coste, mob, nombre):
        self.pep = pep
        self.coste = coste
        self.mob = mob
        self.ingreso = ingreso
        self.nombre = nombre

    def __repr__(self):
        print('PEP: ', self.pep, self.nombre)
        print('INGRESO_POR_RECURSO_PROPIO: ', self.ingreso)
        print('COSTE_POR_RECURSO_PROPIO: ', self.coste)
        print('MOB: ', self.mob, '\n')


def get_month(x):

    return{
        1: 'Enero',
        2: 'Enero',
        3: 'Enero',
        4: 'Enero',
        5: 'Enero',
        6: 'Enero',
        7: 'Enero',
        8: 'Enero',
        9: 'Enero',
        10: 'Enero',
        11: 'Enero',
        12: 'Enero',
    }[x]


filled_peps = []
# LEER LA INFORMACION DE LA SIMULACION DE OIAB
def leer_simulacion():
    wb = load_workbook('Simulación.xlsx')
    sheet = wb['Simulación']
    max_row = wb.active.max_row
    col = 'A'
    peps = []

    for row in range(2, max_row):
        cellPEP = sheet[col + row.__str__()].value
        if cellPEP is not None and cellPEP not in peps:
            peps.append(cellPEP)

    for i in peps:
        ingreso_por_recurso_propio = 0
        coste_por_recurso_propio = 0
        mob = 0
        nombre_proyecto = ''

        for row in range(2, max_row):
            cellPEP = col + row.__str__()

            if not sheet[cellPEP].value is None and sheet[cellPEP].value == i:
                cellConcepto = sheet['AK' + row.__str__()].value
                cellMOB = sheet['AJ' + row.__str__()].value
                cellNombre = sheet['B' + row.__str__()].value
                cellValor = sheet['AR' + row.__str__()].value

                if cellConcepto is not None:
                    if 'INGRESO POR RECURSO PROPIO' in cellConcepto:
                        ingreso_por_recurso_propio = ingreso_por_recurso_propio + cellValor
                    if 'COSTE POR RECURSO PROPIO' in cellConcepto:
                        coste_por_recurso_propio = coste_por_recurso_propio + cellValor
                if cellMOB is not None:
                    if 'MOB-' in cellMOB:
                        mob = mob + cellValor
                if cellNombre is not None:
                    nombre_proyecto = cellNombre

        p = Pep(i, ingreso_por_recurso_propio, coste_por_recurso_propio, mob, nombre_proyecto)
        filled_peps.append(p)


# GRABAR INFORMACION EN LOS FRPS
def grabar_frp():
    # recorrer los xlsx
    for filename in glob.glob(os.path.join('*.xlsx')):
        if 'FRP' in filename:
            if 'D-01350.1.1.1' in filename:
                print("reca")
                wb = load_workbook(filename)
                sheet = wb.active
                if 'D-01350.1.1.1' in sheet['D79'].value: #comprobacion rdundante
                    print ('ok')

                for f in filled_peps:
                    if f.pep == 'D-01350.1.1.1':
                        print(get_month(datetime.now().month))
                        #sheet['F26'].value = get_month(datetime.now().month)

            if 'D-01362.1.1.1' in filename:
                print("contsem")
                wb = load_workbook(filename)
                sheet = wb.active
                if 'D-01362.1.1.1' in sheet['C43'].value:
                    print ('ok')

            if 'D-10168.1.1.1' in filename:
                print("rgpd")
                wb = load_workbook(filename)
                sheet = wb.active
                if 'D-10168.1.1.1' in sheet['C43'].value:
                    print ('ok')

            if 'D-12330.1.1.1' in filename:
                print("webm")
                wb = load_workbook(filename)
                sheet = wb.active
                if 'D-12330.1.1.1' in sheet['C43'].value:
                    print ('ok')


leer_simulacion()
for r in filled_peps: r.__repr__()
grabar_frp()