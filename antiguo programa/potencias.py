from impedancia import Zr,Zl,Zc 
from z_equivalente import Ze 
import openpyxl
import  cmath
#este libro debe ser el libro en que fueron guardados los valores de thevenin
book=openpyxl.load_workbook('data_io.xlsx')
#frecuencia angular
w=377#le coloqué un valor de referencia se debe colocar el valor de w
def crear_datos(book,w):
    sheet=book['VTH_AND_ZTH']
    sheet2=book['V_fuente']
    sheet3=book['I_fuente']
    sheet5=book['Z']
    v_fuente=list()
    i_fuente=list()
    v_nodos=list()
    z_nodos=list()
    i=2
    while sheet[f'B{i}'].value!=None :
        voltaje= (sheet[f'B{i}'].value)*(cmath.cos(sheet[f'C{i}'].value*cmath.pi/180)+cmath.sin(sheet[f'C{i}'].value*cmath.pi/180)*1j)
        v_nodos.append([sheet[f'A{i}'].value,voltaje])
        i+=1
    i=2
    while sheet2[f'A{i}'].value!=None :
       voltaje= (sheet2[f'C{i}'].value/(2**0.5))*(cmath.cos(sheet2[f'D{i}'].value*w)+cmath.sin(sheet2[f'D{i}'].value*w)*1j)
       v_fuente.append([sheet2[f'A{i}'].value,voltaje,Ze(Zr(sheet2[f'E{i}'].value),Zl(w,sheet2[f'F{i}'].value),Zc(w,sheet2[f'Z{i}'].value))])
       i+=1
    i=2
    while sheet3[f'A{i}'].value!=None :
        corriente= (sheet3[f'C{i}'].value/(2**0.5))*(cmath.cos(sheet3[f'D{i}'].value*w)+cmath.sin(sheet3[f'D{i}'].value*w)*1j)
        i_fuente.append([sheet3[f'A{i}'].value,corriente,Ze(Zr(sheet3[f'E{i}'].value),Zl(w,sheet3[f'F{i}'].value),Zc(w,sheet3[f'Z{i}'].value))])
        i+=1
    i=2
    while sheet5[f'A{i}'].value!=None:
        z_nodos.append([sheet5[f'A{i}'].value,sheet5[f'B{i}'].value,Ze(Zr(sheet5[f'D{i}'].value),Zl(w,sheet5[f'E{i}'].value),Zc(w,sheet5[f'F{i}'].value))])
        i+=1
    return v_fuente,v_nodos,i_fuente,z_nodos
#se llama la funcion con las variables en este orden
v_fuente,v_nodos,i_fuente,z_nodos=crear_datos(book,w)
def potencia_fuentes_voltaje(book,v_fuente,v_nodos):
    sheet4=book['Sfuente']
    sheet7=book['Balance_S']
    potencia_total=0
    c=2
    while sheet4[f'A{c}'].value!=None:c+=1
    for x in range(len(v_nodos)):
        for i in range(len(v_fuente)):
            if v_fuente[i][0]==v_nodos[x][0]:
                sheet4[f'A{c}'].value=0
                sheet4[f'B{c}'].value=v_fuente[i][0]
                I_fuente=(v_fuente[i][1]-v_nodos[x][1])/v_fuente[i][2]
                s=v_fuente[i][1]*I_fuente.conjugate()
                sheet4[f'c{c}'].value=s.real
                sheet4[f'D{c}'].value=s.imag   
                potencia_total+=s
                c+=1
    sheet7[f'A{2}'].value=potencia_total.real
    sheet7[f'B{2}'].value=potencia_total.imag  
    g=book['f_and_ouput'] 
    book.save(g['B2'].value)
potencia_fuentes_voltaje(book,v_fuente,v_nodos)
def potencia_fuentes_corriente(book,i_fuente,v_nodos):
    sheet4=book['Sfuente']
    sheet7=book['Balance_S']
    potencia_total=0
    c=2
    while sheet4[f'A{c}'].value!=None:c+=1
    for x in range(len(v_nodos)):
        for i in range(len(i_fuente)):
            if i_fuente[i][0]==v_nodos[x][0]:
                sheet4[f'A{c}'].value=0
                sheet4[f'B{c}'].value=i_fuente[i][0]
                V_fuente=v_nodos[x][1]+(i_fuente[i][2]*i_fuente[i][1])
                s=V_fuente*i_fuente[i][1]
                sheet4[f'c{c}'].value=s.real
                sheet4[f'D{c}'].value=s.imag  
                potencia_total+=s
                c+=1 
    sheet7[f'A{2}'].value=potencia_total.real
    sheet7[f'B{2}'].value=potencia_total.imag 
                
    g=book['f_and_ouput'] 
    book.save(g['B2'].value)   
potencia_fuentes_corriente(book,i_fuente,v_nodos)
def encontrar_nodo(valor,v_nodos):
    for x in range(len(v_nodos)):
        if v_nodos[x][0]==valor: return v_nodos[x]
def  potencia_z(z_nodos,v_nodos,book):
    sheet6=book['S_Z']
    sheet7=book['Balance_S']
    total_potencia=0
    for x in range(len(z_nodos)):
        sheet6[f'A{x+2}'].value=z_nodos[x][0]
        sheet6[f'B{x+2}'].value=z_nodos[x][1]
        nodo_i=encontrar_nodo(z_nodos[x][0],v_nodos)
        nodo_j=encontrar_nodo(z_nodos[x][1],v_nodos)
        v_nodo=abs(nodo_i[1]-nodo_j[1])
        potencia=(v_nodo**2)/z_nodos[x][2]
        sheet6[f'C{x+2}'].value=potencia.real
        sheet6[f'D{x+2}'].value=potencia.imag
        total_potencia+=potencia
    sheet7[f'C{2}'].value=total_potencia.real
    sheet7[f'D{2}'].value=total_potencia.imag
    g=book['f_and_ouput'] 
    book.save(g['B2'].value)
potencia_z(z_nodos,v_nodos,book)
def delta_p(book):
    sheet=book['Balance_S']
    sheet['E2'].value=abs(sheet['A2'].value-sheet['C2'].value)
    sheet['F2'].value=abs(sheet['B2'].value-sheet['D2'].value)
    g=book['f_and_ouput'] 
    book.save(g['B2'].value)
delta_p(book)