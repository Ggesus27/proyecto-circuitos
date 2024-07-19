import numpy as np
import cmath
import openpyxl
def calcular_corrientes_fuente(workbook,frecuencia):
    sheet = workbook['I_fuente']
    corrientes_fuentes = []
    i=2
    while sheet[f'A{i}'].value!=None:
        nodo_referencia=sheet[f'A{i}'].value
        pico_corriente=sheet[f'C{i}'].value
        corrimiento_onda=sheet[f'D{i}'].value
        R=sheet[f'E{i}'].value
        L=sheet[f'F{i}'].value
        C=sheet[f'Z{i}'].value
        if pico_corriente is None or corrimiento_onda is None:
            continue  # O manejar el error de otra manera
        L = (L if L is not None else 0) * 1e-3  # mili
        C = (C if C is not None else 0) * 1e-6  # micro
        if R==0 and L==0 and C==0:
            R=10**6
        if C==0:
            impedancia = R + 1j * (2 * cmath.pi *frecuencia* L)
        else:  
            impedancia = R + 1j * ((2 * cmath.pi *frecuencia* L) - (1 / (2 * cmath.pi*frecuencia * C)))
        angulo_fase = 2 * cmath.pi * frecuencia* corrimiento_onda
        corriente = pico_corriente*complex(cmath.cos(angulo_fase),cmath.sin(angulo_fase))/cmath.sqrt(2)
        corrientes_fuentes.append((nodo_referencia, corriente,impedancia))
        i+=1
    return corrientes_fuentes

#cesar es gay