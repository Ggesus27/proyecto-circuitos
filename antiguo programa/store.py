import math

def storeNodesVI(arr, s, limit):
    for i in range(2, limit):
        if s[f'A{i}'].value == None or s[f'A{i}'].value <= 0:
            return 0 
        
        arr.append(s[f'A{i}'].value)

def storeNodesZ(arr, s, limit):
    for i in range(2, limit):
        if s[f'A{i}'].value == None or s[f'A{i}'].value < 0:
            return 0 
        elif s[f'B{i}'].value == None or s[f'B{i}'].value < 0:
            return 0 
        
        arr[0].append(s[f'A{i}'].value)
        arr[1].append(s[f'B{i}'].value)

def storeVI(arr, z, z1, z2, z3, w, s, limit):
    for i in range(2, limit):
        if s[f'A{i}'].value == None or s[f'A{i}'].value <= 0:
            return 0 

        r = z1(s[f'E{i}'].value)
        l = z2(w, s[f'F{i}'].value)
        c = z3(w, s[f'Z{i}'].value)
        arr.append(z(r, l, c))

def storeZ(arr, z, z1, z2, z3, w, s, limit):
    for i in range(2, limit):
        if s[f'A{i}'].value == None or s[f'A{i}'].value < 0:
            return 0 
        elif s[f'B{i}'].value == None or s[f'B{i}'].value < 0:
            return 0 

        r = z1(s[f'D{i}'].value)
        l = z2(w, s[f'E{i}'].value)
        c = z3(w, s[f'F{i}'].value)
        arr.append(z(r, l, c))

def storeF(arr, s, w, limit):
    for i in range(2, limit):
        if s[f'A{i}'].value == None or s[f'A{i}'].value <= 0:
            return 0 
        
        d = w * s[f'D{i}'].value * (180 / math.pi)
        m = s[f'C{i}'].value / math.sqrt(2)
        fasor = (m, d)
        arr.append(fasor)