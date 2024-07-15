import numpy as np
import openpyxl
import cmath
        
def Cantidad_nodos(nodos):#esta funcion devuelve la cantidad de nodos que hay en el problema
    cant_nodos=1
    for x in range(len(nodos)):
        if nodos[x][0]>cant_nodos or nodos[x][1]>cant_nodos:
            if nodos[x][0]>nodos[x][1]:
                cant_nodos=nodos[x][0]
            else:
                cant_nodos=nodos[x][1]
    return cant_nodos

def matriz_a(canti_nodos,Y): #calcula la matriz de admitancias
    fila=[]
    matriz_y=[]
    b=0
    for x in range(canti_nodos):
        for n in range(canti_nodos):
            valor=0
            for m in Y:
                if x!=n:
                    if m[0]==n+1 and m[1]==x+1:
                        valor-=m[2]
                        continue
                    elif m[0]==x+1 and m[1]==n+1:
                        valor-=m[2]
                        continue
                else:
                    fila.append(suma_matriz(n+1,Y))
                    break
                b+=1
            if b==len(Y):
                fila.append(0)
            elif b!=0:
                fila.append(valor)      
            b=0 
        matriz_y.append(fila)
        fila=[]
    return matriz_y
                    
def suma_matriz(n,Y): #calcula los valores de la diagonal de la matriz
    Y_eq=0
    for m in Y:
        if m[0]==n or m[1]==n:
            Y_eq+=m[2]
    return Y_eq

def matriz_b(I_fuente, cant_nodos,V_fuente=[[0,0,1],[0,0,1]]): #esta funcion calcula la matriz b del sistema de ecuaciones
    matriz=[]
    for x in range(cant_nodos):
        b=0
        for y in I_fuente:
            if y[0]==x+1:
                matriz.append(y[1])
                break
            b+=1
        if b==len(I_fuente): matriz.append(0)
    
    for x in range(len(V_fuente)):
        for y in V_fuente:
            if y[0]==x+1:
                matriz[x]+=y[1]/y[2]
    return matriz

def thevenin(matriz_a,matriz_b,canti_nodos): #calcula voltaje y resistencia de thevenin
    matriz_i=np.array(matriz_a)
    matriz_i=np.linalg.inv(matriz_i)
    matriz_b=np.array(matriz_b)
    Zth=[]
    for x in range(canti_nodos):
        Zth.append(matriz_i[x][x])
    Vth=np.dot(matriz_i,matriz_b)
    return Zth, Vth

def guardar_thevenin(Zth, Vth, book,archivo_salida):#guarda los datos en el archivo de excel
    sheet=book['VTH_AND_ZTH']
    for x in range(len(Zth)):
        sheet[f'B{x+2}'].value=abs(Vth[x])
        sheet[f'C{x+2}'].value=cmath.phase(Vth[x])*180/np.pi
        sheet[f'D{x+2}'].value=Zth[x].real
        sheet[f'E{x+2}'].value=Zth[x].imag
    book.save(archivo_salida)


