def Zr(r):
    return r

def Zl(w, l):
    comp = 1j
    z = (w * l * (10 ** -3)) * comp
    return z

def Zc(w, c):
    comp = -1j
    z = (1 / (w * c * (10 ** -6))) * comp
    return z