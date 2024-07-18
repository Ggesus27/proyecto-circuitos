def calcular_balance(workbook):
    sheet = workbook['Balance_S']
    # Calcular el balance de potencias total del circuito
    total_potencia_fuente = 0
    total_potencia_impedancia = 0
    for row in sheet.iter_rows(min_row=2, values_only = False):
        fuente = row[0].value
        impedancia = row[1].value
        total_potencia_fuente += fuente if fuente is not None else 0
        total_potencia_impedancia += impedancia if impedancia is not None else 0
    balance_potencias = total_potencia_fuente - total_potencia_impedancia
    sheet.cell(row=2, column=3, value=balance_potencias)