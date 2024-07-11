from math import pi

def grados_radianes(grados):
    
    radianes=grados*(pi/180)
    
    return radianes
#print(grados_radianes(180))   Se hizo la prueba, si funciona, da 3.13159......osea pi

def radianes_grados(radianes):
    grados= radianes*(180/pi)
    return grados
#print(radianes_grados(pi))  se hizo la prueba, si funciona, da 180.0
