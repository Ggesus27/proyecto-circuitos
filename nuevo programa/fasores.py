import numpy as np

def convert_to_phasors(impedances):
    phasors = []
    for z in impedances:
        magnitude = np.abs(z)
        angle = np.angle(z, deg = True)  # Convertimos el ángulo a grados
        phasors.append((magnitude, angle))
    return phasors

def convert_fuentes_to_phasors(fuentes_impedances):
    phasors = []
    for nodo, valor_pico, corrimiento, impedance in fuentes_impedances:
        magnitude = np.abs(impedance)
        angle = np.angle(impedance, deg = True)  # Convertimos el ángulo a grados
        phasors.append((nodo, valor_pico, corrimiento, magnitude, angle))
    return phasors