import numpy as np
import openpyxl as px
import cmath

#Operaciones:
#1)Potencias de impedancias entre nodos
def Calcular_potencias_impedancias(vth,impedancias,book,archivo_salida):
    sheet=book['S_Z']
    i=2
    potencia_total=0
    for x in impedancias:
      if x[0]!=0 and x[1]!=0:
        pot=(vth[x[0]-1]-vth[x[1]-1])*(vth[x[0]-1]-vth[x[1]-1]).conjugate()/x[2]
      elif x[0]!=0:
        pot=vth[x[0]-1]*(vth[x[0]-1]/x[2]).conjugate()
      else:
        pot=vth[x[1]-1]*(vth[x[1]-1]/x[2]).conjugate()
      sheet[f'A{i}'].value=x[0]
      sheet[f'B{i}'].value=x[1]
      sheet[f'C{i}'].value=pot.real
      sheet[f'D{i}'].value=pot.imag
      potencia_total+=pot
      i+=1
    book.save(archivo_salida)
    return potencia_total
  #2)Potencia de la fuente
def Calcular_potencia_voltaje(vth,vf):
  potencias=[]
  for x in vf:
    pot=vth[x[0]-1]*((x[1]-vth[x[0]-1])/x[2]).conjugate()
    potencias.append([x[0],pot])
  return potencias

  
  
def calcular_potencia_corriente(vth,I_fuente):
  potencias=[]
  for x in I_fuente:
    Pi=vth[x[0]-1]*x[1].conjugate() 
    potencias.append([x[0],Pi])
  return potencias

def guardar_fuentes(pot_voltaje,pot_corriente,book,archivo_salida):
  sheet=book['Sfuente']
  i=2
  potencia_total=0
  for x in pot_voltaje:
    sheet[f'A{i}'].value=x[0]
    sheet[f'B{i}'].value=x[1].real
    sheet[f'C{i}'].value=x[1].imag
    potencia_total+=x[1]
    i+=1
  for y in pot_corriente:
    sheet[f'A{i}'].value=y[0]
    sheet[f'B{i}'].value=y[1].real
    sheet[f'C{i}'].value=y[1].imag
    i+=1
    potencia_total+=y[1]
  book.save(archivo_salida)
  return potencia_total

def balance(p_t_f,p_t_z,book,archivo_salida):
  sheet=book['Balance_S']
  sheet['A2'].value=p_t_f.real
  sheet['B2'].value=p_t_f.imag
  sheet['C2'].value=p_t_z.real
  sheet['D2'].value=p_t_z.imag
  sheet['E2'].value=p_t_f.real-p_t_z.real
  sheet['F2'].value=abs(p_t_f.imag-p_t_z.imag)
  book.save(archivo_salida)
  