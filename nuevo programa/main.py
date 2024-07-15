import openpyxl
import impedancias
import vth_and_zth as th
from v_fuente import calcular_voltajes_fuente as cal_vol
from I_fuente import calcular_corrientes_fuente as cal_co
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

    try:
        # Leer y verificar datos de entrada
        alertas.verificar_fuentes_y_impedancias(workbook)
        
        # Calcular impedancias y convertir a fasores
        Z = impedancias.calcular_impedancias(workbook)
        v_fuente = cal_vol(workbook, 60)
        i_fuente = cal_co(workbook, 60)

        # Calcular admitancia
        Y = []
        for x in Z:
            Y.append([x[0], x[1], 1 / x[2]])
        
        # Guardar resultados en el archivo Excel
        workbook.save('data_io_output.xlsx')
    except ValueError as e:
        print(e)

if __name__ == '__main__':
    main()