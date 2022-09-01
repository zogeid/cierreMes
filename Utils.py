# devuelve el mes actual
import datetime


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