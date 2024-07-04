def checkV(b, s, limit):
    for i in range(2, limit):
        s[f'B{i}'].value = None

        if s[f'A{i}'].value == None or s[f'A{i}'].value <= 0:
            b.save('data_io.xlsx')
            return 0

        if s[f'C{i}'].value == None:
            s[f'B{i}'].value = 'Falta de argumento'
            s[f'C{i}'].value = 0
            s[f'D{i}'].value = 0
        elif s[f'C{i}'].value == 0:
            s[f'D{i}'].value = 0

        if s[f'D{i}'].value == None:
            s[f'B{i}'].value = 'Falta de argumento'
            s[f'D{i}'].value = 0

        if s[f'E{i}'].value == None:
            s[f'B{i}'].value = 'Falta de argumento'
            s[f'E{i}'].value = 0.0001
        elif s[f'E{i}'].value == 0:
            s[f'B{i}'].value = 'Argumento nulo'
            s[f'E{i}'].value = 0.0001

        if s[f'F{i}'].value == None:
            s[f'B{i}'].value = 'Falta de argumento'
            s[f'F{i}'].value = 0.0001
        elif s[f'F{i}'].value == 0:
            s[f'B{i}'].value = 'Argumento nulo'
            s[f'F{i}'].value = 0.0001

        if s[f'Z{i}'].value == None:
            s[f'B{i}'].value = 'Falta de argumento'
            s[f'Z{i}'].value = 0.0001
        elif s[f'Z{i}'].value == 0:
            s[f'B{i}'].value = 'Argumento nulo'
            s[f'Z{i}'].value = 0.0001

        if s[f'B{i}'].value == None:
            s[f'B{i}'].value = 'OK'

def crashV(b, s, limit):
    for i in range(2, limit):
        if s[f'A{i}'].value == None or s[f'A{i}'].value <= 0:
            return 0
        
        if s[f'C{i}'].value < 0.0000:
            s[f'B{i}'].value = None
            s[f'B{i}'].value = 'Argumento invalido(!)'
            b.save('data_io.xlsx')
            return 1
        
        if s[f'E{i}'].value < 0.0000:
            s[f'B{i}'].value = None
            s[f'B{i}'].value = 'Argumento invalido(!)'
            b.save('data_io.xlsx')
            return 1
        
        if s[f'F{i}'].value < 0.0000:
            s[f'B{i}'].value = None
            s[f'B{i}'].value = 'Argumento invalido(!)'
            b.save('data_io.xlsx')
            return 1
        
        if s[f'Z{i}'].value < 0.0000:
            s[f'B{i}'].value = None
            s[f'B{i}'].value = 'Argumento invalido(!)'
            b.save('data_io.xlsx')
            return 1       