import glob
import os
import datetime
from openpyxl import load_workbook
import pandas as pd
from Pep import Pep
import Utils as U


filled_peps = {}


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
        filled_peps[p.pep] = p


# GRABAR INFORMACION EN LOS FRP'S
def grabar_frp():
    # recorrer los xlsx
    gen = (f for f in glob.glob(os.path.join('*.xlsx')) if 'FRP' in f)

    for filename in gen:
        filename_pep = filename[filename.find('D-'):28]
        for p in filled_peps:
            current_p = filled_peps[p]
            if current_p.pep == filename_pep:
                if current_p.pep == 'D-01350.1.1.1':
                    print("reca")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if U.get_current_year() == 2021:  # row2
                        sheet['F2'].value = U.get_current_month()
                    elif U.get_current_year() == 2022:  # row26
                        sheet['F26'].value = U.get_current_month()

                        suma_mano_obra = 0
                        for i in range(29, 42): # IMPORTANT! Si se añaden más perfiles/tarifas nuevas comprobar el rango
                            try:
                                suma_mano_obra = suma_mano_obra + (sheet[f'C{i}'].value * sheet[f'{U.get_current_month_column()}{i}'].value)
                            except TypeError: # Falla si hay ceros en la multiplicacion
                                pass

                        if round(suma_mano_obra, 0) == round(current_p.coste, 0):  # Comprobamos que costes coinciden
                            sheet[f'{U.get_current_month_column()}47'].value = current_p.mob
                            sheet[f'{U.get_current_month_column()}52'].value = current_p.ingreso

                    costes = suma_mano_obra + current_p.mob
                    ingreso_mob = costes / (1-sheet[f'{U.get_current_month_column()}50'].value)
                    reg = round(ingreso_mob - current_p.ingreso, 2)
                    print(f'REGULARIZAR: ', {reg})
                    wb.save(filename)
                    break

                if current_p.pep == 'D-01362.1.1.1':
                    print("contsem")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if 'D-01362.1.1.1' in sheet['C43'].value:
                        print ('ok')
                    break

                if current_p.pep == 'D-10168.1.1.1':
                    print("rgpd")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if 'D-10168.1.1.1' in sheet['C43'].value:
                        print ('ok')
                    break

                if current_p.pep == 'D-12330.1.1.1':
                    print("webm")
                    wb = load_workbook(filename)
                    sheet = wb.active
                    if 'D-12330.1.1.1' in sheet['C43'].value:
                        print ('ok')
                    break


leer_simulacion()
for r in filled_peps: r.__repr__()
grabar_frp()