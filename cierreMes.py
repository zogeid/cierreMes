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


# devuelve el mes actual
def get_current_month():
    return 'Mayo'
    # return{
    #     1: 'Enero',
    #     2: 'Febrero',
    #     3: 'Marzo',
    #     4: 'Abril',
    #     5: 'Mayo',
    #     6: 'Junio',
    #     7: 'Julio',
    #     8: 'Agosto',
    #     9: 'Septiembre',
    #     10: 'Octubre',
    #     11: 'Noviembre',
    #     12: 'Diciembre',
    # }[datetime.datetime.now().month]


# devuelve la columna del mes actual
def get_current_month_column():
    return 'H'
    # return{
    #     1: 'D',
    #     2: 'E',
    #     3: 'F',
    #     4: 'G',
    #     5: 'H',
    #     6: 'I',
    #     7: 'J',
    #     8: 'K',
    #     9: 'L',
    #     10: 'M',
    #     11: 'N',
    #     12: 'O',
    # }[datetime.datetime.now().month]


# devuelve el año actual
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

    # recorre Simulacion.xlsx para obtener los diferentes PEPs y guardarlos en peps
    for row in range(2, max_row):
        cell_pep = sheet[col + row.__str__()].value
        if cell_pep is not None and cell_pep not in peps:
            peps.append(cell_pep)

    # Para cada PEP recoge datos de concepto, mob, nombre
    # En funcion del concepto acumula los importes de Ingreso, Coste y MOB
    # Con esta información genera un objeto PEP y lo añade a filled_peps
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
                cell_importe = sheet['AR' + row.__str__()].value

                if cell_concepto is not None:
                    if 'INGRESO POR RECURSO PROPIO' in cell_concepto:
                        ingreso_por_recurso_propio = ingreso_por_recurso_propio + cell_importe
                    if 'COSTE POR RECURSO PROPIO' in cell_concepto:
                        coste_por_recurso_propio = coste_por_recurso_propio + cell_importe
                if cell_mob is not None:
                    if 'MOB-' in cell_mob:
                        mob = mob + cell_importe
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
                if p.pep == 'D-01350.1.1.1':
                    print("reca")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if get_current_year() == 2021:  # row2
                        sheet['F2'].value = get_current_month()
                    elif get_current_year() == 2022:  # row26
                        sheet['F26'].value = get_current_month()

                        suma_mano_obra = 0
                        for i in range(29, 36):
                            try:
                                suma_mano_obra = suma_mano_obra + (sheet[f'C{i}'].value * sheet[f'{get_current_month_column()}{i}'].value)
                            except:
                                pass
                        print(sheet['H42'].internal_value)
                        if round(suma_mano_obra, 0) == round(p.coste, 0):  # Comprobamos que los costes coinciden
                            sheet[f'{get_current_month_column()}41'].value = p.mob
                            sheet[f'{get_current_month_column()}46'].value = p.ingreso

                    wb.save(filename)
                    break

                if p.pep == 'D-01362.1.1.1':
                    print("contsem")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if 'D-01362.1.1.1' in sheet['C43'].value:
                        print ('ok')
                    break

                if p.pep == 'D-10168.1.1.1':
                    print("rgpd")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if 'D-10168.1.1.1' in sheet['C43'].value:
                        print ('ok')
                    break

                if p.pep == 'D-12330.1.1.1':
                    print("webm")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if 'D-12330.1.1.1' in sheet['C43'].value:
                        print ('ok')
                    break


leer_simulacion()
for r in filled_peps: r.__repr__()
grabar_frp()