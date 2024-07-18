def leer_frecuencia_y_salida(workbook):
    sheet = workbook['f_and_ouput']
    frecuencia = sheet['B1'].value  # Supongo que la frecuencia está en la celda A2
    archivo_salida = sheet['B2'].value  # Supongo que el nombre del archivo de salida está en la celda B2
    return frecuencia, archivo_salida