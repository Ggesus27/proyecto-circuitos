import numpy as np
from openpyxl import load_workbook

def calcular_impedancias(workbook,frecuencia):
    sheet = workbook['Z']
    warnings = []
    impedancia=[]
    # Iterar sobre todas las filas a partir de la segunda (asumiendo que la primera es la cabecera)
    for i in range(2, sheet.max_row + 1):
        R = sheet[f'D{i}'].value
        L = sheet[f'E{i}'].value
        C = sheet[f'F{i}'].value
        nodoi=sheet[f'A{i}'].value
        nodoj=sheet[f'B{i}'].value
        # Si todos los valores son nulos, se detiene el cálculo
        if R is None and L is None and C is None:
            break

        R = R if R is not None else 0
        L = (L if L is not None else 0)*10**-3
        C = (C if C is not None else 0)*10**-6
        if R==0 and L==0 and C==0: R=10**-6
        try:
            if C == 0:
                Z = complex(R, 2 * np.pi*frecuencia * L)  # Evitar la división por cero
            else:
                Z = complex(R,((2 * np.pi * L*frecuencia) - (1 / (2 * np.pi *frecuencia* C))))
        except ZeroDivisionError:
            warnings.append(f"Fila {i}: División por cero detectada en el cálculo de Z.")
            break
        impedancia.append([nodoi,nodoj,Z])

    return impedancia
# Cargar el archivo Excel y llamar a la función
