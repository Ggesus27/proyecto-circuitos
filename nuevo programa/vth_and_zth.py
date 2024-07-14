def calcular_vth_zth(workbook):
    sheet = workbook['VTH_AND_ZTH']
    # Implementar el cálculo de Vth y Zth según los valores de las hojas pertinentes
    for row in sheet.iter_rows(min_row=2, values_only=False):
        # Suponiendo que los datos están en las columnas correspondientes
        nodo, vth, zth = row[0].value, row[1].value, row[2].value
        if zth is not None:
            zth = complex(zth)
            if zth.imag < 0:
                zth_type = "capacitivo"
            else:
                zth_type = "inductivo"
        sheet.cell(row=row[0].row, column=4, value=zth_type)