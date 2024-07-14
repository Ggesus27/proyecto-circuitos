import numpy as np

def calcular_impedancias(workbook):
    sheet = workbook['Z']
    # Calcular impedancias según los valores en la hoja 'Z'
    for row in sheet.iter_rows(min_row=2, values_only=False):
        R = row[0].value if row[0].value is not None else 0
        L = (row[1].value if row[1].value is not None else 0) * 1e-3  # mili
        C = (row[2].value if row[2].value is not None else 0) * 1e-6  # micro
        Z = R + 1j * (2 * np.pi * L - 1 / (2 * np.pi * C))
        sheet.cell(row=row[0].row, column=4, value=Z)