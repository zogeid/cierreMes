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


def get_month():
    return{
        1: 'Enero',
        2: 'F',
        3: 'M',
        4: 'A',
        5: 'M',
        6: 'Junio',
        7: 'Julio',
        8: 'A',
        9: 'S',
        10: 'O',
        11: 'N',
        12: 'D',
    }[datetime.datetime.now().month]


def get_current_year():
    return datetime.datetime.now().year


filled_peps = []


# LEER LA INFORMACION DE LA SIMULACION DE OIAB
def leer_simulacion():
    wb = load_workbook('Simulación.xlsx')
    sheet = wb['Simulación']
    max_row = wb.active.max_row
    col = 'A'
    peps = []

    for row in range(2, max_row):
        cell_pep = sheet[col + row.__str__()].value
        if cell_pep is not None and cell_pep not in peps:
            peps.append(cell_pep)

    for i in peps:
        ingreso_por_recurso_propio = 0
        coste_por_recurso_propio = 0
        mob = 0
        nombre_proyecto = ''

        for row in range(2, max_row):
            cell_pep = col + row.__str__()

            if not sheet[cell_pep].value is None and sheet[cell_pep].value == i:
                cell_concepto = sheet['AK' + row.__str__()].value
                cell_mob = sheet['AJ' + row.__str__()].value
                cell_nombre = sheet['B' + row.__str__()].value
                cell_valor = sheet['AR' + row.__str__()].value

                if cell_concepto is not None:
                    if 'INGRESO POR RECURSO PROPIO' in cell_concepto:
                        ingreso_por_recurso_propio = ingreso_por_recurso_propio + cell_valor
                    if 'COSTE POR RECURSO PROPIO' in cell_concepto:
                        coste_por_recurso_propio = coste_por_recurso_propio + cell_valor
                if cell_mob is not None:
                    if 'MOB-' in cell_mob:
                        mob = mob + cell_valor
                if cell_nombre is not None:
                    nombre_proyecto = cell_nombre

        p = Pep(i, ingreso_por_recurso_propio, coste_por_recurso_propio, mob, nombre_proyecto)
        filled_peps.append(p)


# GRABAR INFORMACION EN LOS FRP'S
def grabar_frp():
    # recorrer los xlsx
    gen = (f for f in glob.glob(os.path.join('*.xlsx')) if 'FRP' in f)

    for filename in gen:
        filename_pep = filename[filename.find('D-'):28]
        for p in filled_peps:
            if p.pep == filename_pep:
        #for p in filled_peps: # en vez de for, sacar el valor concreto de filled_peps
        #   if p.pep in filename:
                if p.pep == 'D-01350.1.1.1':
                    print("reca")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    print(get_month(), get_current_year())
                    if get_current_year() == 2022:  # row26
                        sheet['F26'].value = get_month()
                    break

                if p.pep == 'D-01362.1.1.1':
                    print("contsem")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if 'D-01362.1.1.1' in sheet['C43'].value:
                        print ('ok')

                if p.pep == 'D-10168.1.1.1':
                    print("rgpd")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if 'D-10168.1.1.1' in sheet['C43'].value:
                        print ('ok')

                if p.pep == 'D-12330.1.1.1':
                    print("webm")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if 'D-12330.1.1.1' in sheet['C43'].value:
                        print ('ok')


leer_simulacion()
#for r in filled_peps: r.__repr__()
grabar_frp()