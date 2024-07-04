def generateOutputVI(name, out, s1, s2):
    s1['A2'].value = 'i'
    s2['A2'].value = 'i'

    s1.merge_cells('B1:C1')
    s1.merge_cells('D1:E1')
    s2.merge_cells('B1:C1')
    s2.merge_cells('D1:E1')

    s1['B1'].value = 'Z equivalentes'
    s1['D1'].value = 'Fasores (Vf)'
    s2['B1'].value = 'Z equivalentes'
    s2['D1'].value = 'Fasores (If)'

    s1['B2'].value = 'Real'
    s1['C2'].value = 'Imag'
    s1['D2'].value = 'Module'
    s1['E2'].value = 'Degrees'
    s2['B2'].value = 'Real'
    s2['C2'].value = 'Imag'
    s2['D2'].value = 'Module'
    s2['E2'].value = 'Degrees'

    out.save(name)
    return 0

def generateOutputZ(name, out, s):
    s['A2'].value = 'i'
    s['B2'].value = 'j'

    s.merge_cells('C1:D1')

    s['C1'].value = 'Z equivalentes'

    s['C2'].value = 'Real'
    s['D2'].value = 'Imag'

    out.save(name)
    return 0

def generateOutputTH(name, out, s):
    s['A1'].value = 'i'
    s['B1'].value = 'VTH'
    s['C1'].value = 'Degrees'
    s['D1'].value = 'RTH'
    s['E1'].value = 'XTH'

    out.save(name)
    return 0

def generateOutputSZ(name, out, s):
    s['A1'].value = 'i'
    s['B1'].value = 'j'
    s['C1'].value = 'P[W]'
    s['D1'].value = 'Q[Var]'

    out.save(name)
    return 0

def generateOutputBZ(name, out, s):
    s['A1'].value = 'Pf total'
    s['B1'].value = 'Qf total'
    s['C1'].value = 'Pz total'
    s['D1'].value = 'Qz total'
    s['E1'].value = 'Delta Q total'

    out.save(name)
    return 0