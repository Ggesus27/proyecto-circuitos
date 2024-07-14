lista=[[1,2,5],[2,3,3]]
#colummna 1 nodo i, columna 2 nodo j, columna 3 impedancia
vth=[1,2,3,54]
potencias=[]
for x in lista:
    pot=(vth[x[0]-1]-vth[x[1]-1])**2/x[2]
    potencias.append([x[0],x[1],pot])

(vf-vth)*vf/z

vf=[[1,100,5],[2,355,23]]
#columna 1 nodo, columna 2 voltaje, columna 3 impedancia
vth[x[0]-1]*(vf[x[0]]-vth[x[0]-1])/vf[2] #cmath

vth*i #conjugado
def funcion(impedancias, vth, vfuente, ifuente):
    
