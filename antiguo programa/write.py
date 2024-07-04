import math
def writeSheetVI(arr, node, f, name, out, s):
    for i in range(3, len(arr)):
        mod, deg = f[i]

        s[f'A{i}'].value = node[i]
        s[f'B{i}'].value = arr[i].real
        s[f'C{i}'].value = arr[i].imag
        s[f'D{i}'].value = mod
        s[f'E{i}'].value = deg
    out.save(name)

def writeSheetZ(arr, node, name, out, s):
    for i in range(3, len(arr)):
        s[f'A{i}'].value = node[0][i]
        s[f'B{i}'].value = node[1][i]
        s[f'C{i}'].value = arr[i].real
        s[f'D{i}'].value = arr[i].imag
    out.save(name)

def writeSheetTH(arr, node, f, name, out, s):
    for i in range(3, len(arr)):
        mod, deg = f[i]

        s[f'A{i - 1}'].value = node[i]
        s[f'B{i - 1}'].value = mod
        s[f'C{i - 1}'].value = deg
        s[f'D{i - 1}'].value = arr[i].real
        s[f'E{i - 1}'].value = math.abs(arr[i].imag)
    out.save(name)