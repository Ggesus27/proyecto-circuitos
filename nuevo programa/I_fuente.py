import numpy as np

def calcular_corrientes_fuente(workbook):
    sheet = workbook['I_fuente']
    corrientes_fuentes = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        nodo_referencia, pico_corriente, corrimiento_onda, R, L, C = row
        if pico_corriente is None or corrimiento_onda is None:
            continue  # O manejar el error de otra manera
        L = (L if L is not None else 0) * 1e-3  # mili
        C = (C if C is not None else 0) * 1e-6  # micro
        impedancia = R + 1j * (2 * np.pi * L - 1 / (2 * np.pi * C))
        angulo_fase = 2 * np.pi * corrimiento_onda
        corriente = pico_corriente * np.exp(1j * angulo_fase)
        corrientes_fuentes.append((nodo_referencia, corriente))
    return corrientes_fuentes