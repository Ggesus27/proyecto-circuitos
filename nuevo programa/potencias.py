import numpy as np
import openpyxl as px
import cmath
#Primero hay que llamar los valores
from impedancias import calcular_impedancias as impedancia

lista=[[1,2,impedancia],[2,3,impedancia]]
#colummna 1 nodo i, columna 2 nodo j, columna 3 impedancia


#Operaciones:
#1)Potencias de impedancias entre nodos
def Calcular_potencias_impedancias():
    vth=[1,2,3,54]
    potencias_z=[]
    for x in lista:pot=(vth[x[0]-1]-vth[x[1]-1])**2/x[2]
    potencias_z.append([x[0],x[1],pot])
    return potencias_z

  #2)Potencia de la fuente
def Calcular_potencia_voltaje(vth,vf):
   
    vth=[1,2,3,54]
    potencias_v=[]
    #columna 1 nodo, columna 2 voltaje, columna 3 impedancia
    for x in lista:pot=vth[x[0]-1]*(vf[x[0]]-vth[x[0]-1])/vf[2]
    potencias_v.append([x[0],x[1],pot])
    return potencias_v
    

from I_fuente import calcular_corrientes_fuente as I
#3)Potencia de la fuente de corriente
i=I
ic= i.conjugate() #Corriente conjugada
def calcular_potencia_corriente(vth):
    potencias_i=[]
    Pi=vth*ic 
    potencias_i.append(Pi)
    return Pi



#4)Potencias de impedancias en serie con una fuente:
  #nose :c
    