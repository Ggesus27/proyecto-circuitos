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
        
def Cantidad_nodos(nodos):#esta funcion devuelve la cantidad de nodos que hay en el problema
    cant_nodos=1
    for x in range(len(nodos)):
        if nodos[x][0]>cant_nodos or nodos[x][1]>cant_nodos:
            if nodos[x][0]>nodos[x][1]:
                cant_nodos=nodos[x][0]
            else:
                cant_nodos=nodos[x][1]
    return cant_nodos

def matriz(canti_nodos,Y): #calcula la matriz de admitancias
    fila=[]
    matriz_y=[]
    for x in range(1,canti_nodos):
        for n in range(1,canti_nodos):
            b=0
            for m in Y:
                if x!=n:
                    if m[0]==n and m[1]==x:
                        fila.append(-1*m[2])
                        break
                    elif m[0]==x and m[1]==n:
                        fila.append(-1*m[2])
                        break
                else:
                    fila.append(suma_matriz(n,Y))
                    break
                b+=1
            if b==len(Y):
                fila.append(0)       
        matriz_y.append(fila)
        fila=[]
    return matriz_y
                    
def suma_matriz(n,Y): #calcula los valores de la diagonal de la matriz
    Y_eq=0
    for m in Y:
        if m[0]==n or m[1]==n:
            Y_eq+=m[2]
    return Y_eq