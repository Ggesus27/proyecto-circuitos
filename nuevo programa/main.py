import openpyxl
import impedancias
import vth_and_zth as th
import potencias
import balance
import alertas

def main():
    # Especificar la ruta relativa al archivo Excel
    workbook_path = 'data_io.xlsx'
    
    try:
        # Cargar archivo Excel
        workbook = openpyxl.load_workbook(workbook_path)
    except FileNotFoundError:
        print("Error: El archivo no se encuentra en la ruta especificada.")
        return
    except PermissionError:
        print("Error: No se tienen permisos para acceder al archivo.")
        return

    # Leer y verificar datos de entrada
    f_and_output_sheet = workbook['f_and_ouput']
    alertas.verificar_datos(workbook, f_and_output_sheet)
    
    # Calcular impedancias y convertir a fasores
    Z=impedancias.calcular_impedancias(workbook)

    #calcular conductancia
    Y=[]
    for x in Z:
        Y.append(1/Z)
    
    
    # Guardar resultados en el archivo Excel
    workbook.save('data_io_output.xlsx')

if __name__ == '__main__':
    main()