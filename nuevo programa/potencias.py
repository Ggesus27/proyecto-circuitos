def calcular_potencias(workbook):
    sheet_fuentes = workbook['S_fuente']
    for row in sheet_fuentes.iter_rows(min_row=2, values_only=False):
        nodo, potencia = row[0].value, row[1].value
        # Calcular la potencia de las fuentes
        sheet_fuentes.cell(row=row[0].row, column=3, value=potencia)  # Ejemplo de cálculo

    sheet_impedancias = workbook['S_Z']
    for row in sheet_impedancias.iter_rows(min_row=2, values_only=False):
        nodo, impedancia = row[0].value, row[1].value
        # Calcular la potencia de las impedancias
        sheet_impedancias.cell(row=row[0].row, column=3, value=impedancia)  # Ejemplo de cálculo