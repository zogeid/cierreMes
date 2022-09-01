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