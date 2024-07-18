import openpyxl 
def verificar_fuentes(sheet):
    i=2
    while sheet[f'A{i}'].value!=None and sheet[f'C{i}'].value!=None and sheet[f'D{i}'].value!=None and sheet[f'E{i}'].value!=None and sheet[f'F{i}'].value!=None and sheet[f'Z{i}'].value!=None:
        if sheet[f'A{i}'].value<=0:
            sheet[f'B{i}'].value="Error en el valor del nodo, debe ser mayor o igual a 1"
            return 1
        elif sheet[f'A{i}'].value==None:
            sheet[f'B{i}'].value="falta valor del nodo"
            return 1
        elif type(sheet[f'A{i}'].value)!=type(2):
            sheet[f'B{i}'].value="Error en el valor del nodo, debe ser un numero entero"
            return 1
        elif type(sheet[f'C{i}'].value)!=type(2.5) and type(sheet[f'C{i}'].value)!=type(2):
            sheet[f'B{i}'].value="Error en el valor pico de la fuente, debe ser un numero"
            return 1
        elif type(sheet[f'D{i}'].value)!=type(2.5) and type(sheet[f'D{i}'].value)!=type(2):
            sheet[f'B{i}'].value="Error en el valor del corrimiento de onda, debe ser un numero"
            return 1
        elif type(sheet[f'E{i}'].value)!=type(2.5) and type(sheet[f'E{i}'].value)!=type(2):
            sheet[f'B{i}'].value="Error en el valor de la resistencia, debe ser un numero"
            return 1
        elif type(sheet[f'F{i}'].value)!=type(2.5) and type(sheet[f'F{i}'].value)!=type(2):
            sheet[f'B{i}'].value="Error en el valor de la inductancia, debe ser un numero"
            return 1
        elif type(sheet[f'Z{i}'].value)!=type(2.5) and type(sheet[f'Z{i}'].value)!=type(2):
            sheet[f'B{i}'].value="Error en el valor de la capacitancia, debe ser un numero"
            return 1
        elif sheet[f'E{i}'].value<0:
            sheet[f'B{i}'].value="Error en el valor de la resistencia, debe ser un valor positivo"
            return 1
        elif sheet[f'F{i}'].value<0:
            sheet[f'B{i}'].value="Error en el valor de la inductancia, debe ser un valor positivo"
            return 1
        elif sheet[f'Z{i}'].value<0:
            sheet[f'B{i}'].value="Error en el valor de la capacitancia, debe ser un valor positivo"
            return 1
        elif sheet[f'C{i}'].value<=0:
            sheet[f'B{i}'].value="Error en el valor pico de la fuente, debe ser mayor que cero"
            return 1
        elif sheet[f'C{i}'].value==None:
            sheet[f'B{i}'].value="falta valor pico de la fuente"
            return 1
        elif sheet[f'D{i}'].value==None:
            sheet[f'B{i}'].value="falta valor del corrimiento de onda"
            return 1
        
        if sheet[f'E{i}'].value==None and sheet[f'F{i}'].value==None and sheet[f'Z{i}'].value==None:
            sheet[f'E{i}'].value=0
            sheet[f'F{i}'].value=0
            sheet[f'Z{i}'].value=0
            sheet[f'B{i}'].value="se asume impedancia igual a cero"
        elif sheet[f'Z{i}'].value==None and sheet[f'F{i}'].value==None:
            sheet[f'Z{i}'].value=0
            sheet[f'F{i}'].value=0
            sheet[f'B{i}'].value="se asume inductancia y capacitancia igual a 0"
        elif sheet[f'Z{i}'].value==None and sheet[f'E{i}'].value==None:
            sheet[f'Z{i}'].value=0
            sheet[f'E{i}'].value=0
            sheet[f'B{i}'].value="se asume resistencia y capacitancia igual a 0"
        elif sheet[f'E{i}'].value==None and sheet[f'F{i}'].value==None:
            sheet[f'E{i}'].value=0
            sheet[f'F{i}'].value=0
            sheet[f'B{i}'].value="se asume resistencia e inductancia igual a 0"
        elif sheet[f'E{i}'].value==None:
            sheet[f'B{i}'].value="Se asume resistencia 0"
            sheet[f'E{i}'].value=0
        elif sheet[f'F{i}'].value==None:
            sheet[f'B{i}'].value="se asume inductancia igual a 0"
        elif sheet[f'Z{i}'].value==None:
            sheet[f'B{i}'].value="se asume capacitancia igual a 0"
            sheet[f'Z{i}'].value=0
        else:
            sheet[f'B{i}'].value="ok"
        i+=1
    return 0

def verificar_impedancias(sheet):
    i=2
    while sheet[f'A{i}'].value!=None and sheet[f'B{i}'].value!=None and sheet[f'D{i}'].value!=None and sheet[f'E{i}'].value!=None and sheet[f'F{i}'].value!=None:

        if type(sheet[f'A{i}'].value)!=type(2):
            sheet[f'C{i}'].value="Error en el valor del nodo, debe ser un numero entero"
            return 1
        elif type(sheet[f'B{i}'].value)!=type(2):
            sheet[f'C{i}'].value="Error en el valor del nodo, debe ser un numero entero"
            return 1
        elif type(sheet[f'D{i}'].value)!=type(2.5) and type(sheet[f'D{i}'].value)!=type(2):
            sheet[f'C{i}'].value="Error en el valor de la resistencia, debe ser un numero"
            return 1
        elif type(sheet[f'E{i}'].value)!=type(2.5) and type(sheet[f'E{i}'].value)!=type(2):
            sheet[f'C{i}'].value="Error en el valor de la inductanicia, debe ser un numero"
            return 1
        elif type(sheet[f'F{i}'].value)!=type(2.5) and type(sheet[f'F{i}'].value)!=type(2):
            sheet[f'C{i}'].value="Error en la capacitancia, debe ser un numero"
            return 1
        elif sheet[f'A{i}'].value==None:
            sheet[f'C{i}'].value="falta valor del nodo i"
            return 1
        elif sheet[f'B{i}'].value==None:
            sheet[f'C{i}'].value="falta valor del nodo j"
            return 1
        elif sheet[f'A{i}'].value<0:
            sheet[f'C{i}'].value="Error en el valor del nodo, debe ser mayor o igual a cero"
            return 1
        elif sheet[f'B{i}'].value<0 :
            sheet[f'C{i}'].value="Error en el valor del nodo, debe ser mayor o igual a cero"
            return 1
        elif sheet[f'D{i}'].value<0:
            sheet[f'C{i}'].value="Error en el valor de la resistencia, debe ser un valor positivo"
            return 1
        elif sheet[f'E{i}'].value<0:
            sheet[f'C{i}'].value="Error en el valor de la inductancia, debe ser un valor positivo"
            return 1
        elif sheet[f'F{i}'].value<0:
            sheet[f'C{i}'].value="Error en el valor de la capacitancia, debe ser un valor positivo"
            return 1
        
        
        if sheet[f'D{i}'].value==None and sheet[f'E{i}'].value==None and sheet[f'F{i}'].value==None:
            sheet[f'D{i}'].value=0
            sheet[f'E{i}'].value=0
            sheet[f'F{i}'].value=0
            sheet[f'C{i}'].value="se asume impedancia igual a cero"
        elif sheet[f'E{i}'].value==None and sheet[f'F{i}'].value==None:
            sheet[f'E{i}'].value=0
            sheet[f'F{i}'].value=0
            sheet[f'C{i}'].value="se asume inductancia y capacitancia igual a 0"
        elif sheet[f'D{i}'].value==None and sheet[f'F{i}'].value==None:
            sheet[f'D{i}'].value=0
            sheet[f'F{i}'].value=0
            sheet[f'C{i}'].value="se asume resistencia y capacitancia igual a 0"
        elif sheet[f'E{i}'].value==None and sheet[f'D{i}'].value==None:
            sheet[f'E{i}'].value=0
            sheet[f'D{i}'].value=0
            sheet[f'C{i}'].value="se asume resistencia e inductancia igual a 0"
        elif sheet[f'D{i}'].value==None:
            sheet[f'C{i}'].value="Se asume resistencia 0"
            sheet[f'D{i}'].value=0
        elif sheet[f'E{i}'].value==None:
            sheet[f'C{i}'].value="se asume inductancia igual a 0"
            sheet[f'E{i}'].value=0
        elif sheet[f'F{i}'].value==None:
            sheet[f'C{i}'].value="se asume capacitancia igual a 0"
            sheet[f'F{i}'].value=0
        else:
            sheet[f'C{i}'].value="ok"
        i+=1
    return 0