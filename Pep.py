class Pep:
    def __init__(self, pep, ingreso, coste, mob, nombre):
        self.pep = pep
        self.coste = coste
        self.mob = mob
        self.ingreso = ingreso
        self.nombre = nombre
        self.cierre_coste = 0
        self.cierre_ingreso = 0
        self.regularizacion = 0

    def __repr__(self):
        print('PEP: ', self.pep, self.nombre)
        print('INGRESO_POR_RECURSO_PROPIO: ', self.ingreso)
        print('COSTE_POR_RECURSO_PROPIO: ', self.coste)
        print('MOB: ', self.mob)
        print('CIERRE_COSTE: ', self.cierre_coste)
        print('CIERRE_INGRESO: ', self.cierre_ingreso)
        print('REGULARIZACION: ', self.regularizacion, '\n')