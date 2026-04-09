"""Datos para el ejercicio CreditMetrics.

Notas de modelación:
- El espacio de estados usa 18 calificaciones: AAA, AA+, AA, AA-, A+, A, A-,
  BBB+, BBB, BBB-, BB+, BB, BB-, B+, B, B-, CCC y D.
- La tabla de transición del PDF original incluye una columna NR. Como la clase
  habla de 18^3 escenarios, por defecto el proyecto elimina NR y renormaliza las
  probabilidades restantes a 1.
- La curva libre de riesgo proviene del dato Daily Treasury Par Yield Curve Rates
  proporcionado en el enunciado. Para 4 años se interpola linealmente usando el
  promedio simple entre 3Y y 5Y, igual que en la plantilla de Excel compartida.
- Las sobretasas por rating usan la lógica de interpolación de la plantilla de
  Excel de 2 bonos, que a su vez replica la curva coarse-to-notch con spreads
  Damodaran.
"""

from dataclasses import dataclass


STATES_18 = [
    "AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-",
    "BB+", "BB", "BB-", "B+", "B", "B-", "CCC", "D",
]

# Curva libre de riesgo del enunciado (04/08/2026). 4Y se interpola como en Excel.
TREASURY_CURVE = {
    1: 0.0369,
    2: 0.0379,
    3: 0.0378,
    4: (0.0378 + 0.0392) / 2.0,
    5: 0.0392,
}

# Spreads coarse de Damodaran reconstruidos desde la plantilla de Excel.
COARSE_SPREADS = {
    "AAA": 0.0069,
    "AA": 0.0085,
    "A+": 0.0107,
    "A": 0.0118,
    "A-": 0.0133,
    "BBB": 0.0171,
    "BB+": 0.0231,
    "BB": 0.0277,
    "B+": 0.0405,
    "B": 0.0486,
    "B-": 0.0594,
    "CCC": 0.0946,
}

# Interpolación notch-a-notch siguiendo la plantilla Excel subida por el usuario.
RATING_SPREADS = {
    "AAA": COARSE_SPREADS["AAA"],
    "AA+": (COARSE_SPREADS["AAA"] + COARSE_SPREADS["AA"]) / 2.0,
    "AA": COARSE_SPREADS["AA"],
    "AA-": (COARSE_SPREADS["AA"] + COARSE_SPREADS["A+"]) / 2.0,
    "A+": COARSE_SPREADS["A+"],
    "A": COARSE_SPREADS["A"],
    "A-": COARSE_SPREADS["A-"],
    "BBB+": (COARSE_SPREADS["A-"] + COARSE_SPREADS["BBB"]) / 2.0,
    "BBB": COARSE_SPREADS["BBB"],
    "BBB-": (COARSE_SPREADS["BBB"] + COARSE_SPREADS["BB+"]) / 2.0,
    "BB+": COARSE_SPREADS["BB+"],
    "BB": COARSE_SPREADS["BB"],
    "BB-": (COARSE_SPREADS["BB"] + COARSE_SPREADS["B+"]) / 2.0,
    "B+": COARSE_SPREADS["B+"],
    "B": COARSE_SPREADS["B"],
    "B-": COARSE_SPREADS["B-"],
    "CCC": COARSE_SPREADS["CCC"],
}

# Matriz de transición 18x18 sin NR, extraída de la tabla del PDF/Excel.
TRANSITION_MATRIX_RAW_18 = {
    "AAA": [0.8709, 0.0586, 0.0250, 0.0068, 0.0016, 0.0024, 0.0013, 0.0000, 0.0005, 0.0000, 0.0003, 0.0005, 0.0003, 0.0000, 0.0003, 0.0000, 0.0005, 0.0000],
    "AA+": [0.0221, 0.7968, 0.1059, 0.0338, 0.0068, 0.0032, 0.0018, 0.0005, 0.0009, 0.0005, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000, 0.0000],
    "AA": [0.0041, 0.0131, 0.8099, 0.0842, 0.0263, 0.0112, 0.0035, 0.0038, 0.0013, 0.0008, 0.0005, 0.0003, 0.0002, 0.0002, 0.0000, 0.0002, 0.0005, 0.0002],
    "AA-": [0.0004, 0.0010, 0.0368, 0.7899, 0.0974, 0.0218, 0.0057, 0.0024, 0.0014, 0.0006, 0.0003, 0.0000, 0.0000, 0.0003, 0.0008, 0.0000, 0.0000, 0.0003],
    "A+": [0.0000, 0.0005, 0.0041, 0.0423, 0.7916, 0.0854, 0.0204, 0.0058, 0.0032, 0.0008, 0.0005, 0.0008, 0.0001, 0.0006, 0.0003, 0.0000, 0.0000, 0.0005],
    "A": [0.0003, 0.0004, 0.0021, 0.0039, 0.0519, 0.7936, 0.0661, 0.0232, 0.0081, 0.0025, 0.0009, 0.0010, 0.0006, 0.0008, 0.0002, 0.0000, 0.0001, 0.0005],
    "A-": [0.0003, 0.0001, 0.0005, 0.0014, 0.0038, 0.0614, 0.7887, 0.0727, 0.0185, 0.0053, 0.0012, 0.0012, 0.0010, 0.0010, 0.0003, 0.0001, 0.0003, 0.0005],
    "BBB+": [0.0000, 0.0001, 0.0005, 0.0006, 0.0019, 0.0069, 0.0677, 0.7665, 0.0798, 0.0150, 0.0034, 0.0026, 0.0012, 0.0014, 0.0009, 0.0002, 0.0006, 0.0009],
    "BBB": [0.0001, 0.0001, 0.0004, 0.0002, 0.0009, 0.0028, 0.0095, 0.0726, 0.7668, 0.0620, 0.0128, 0.0063, 0.0026, 0.0020, 0.0010, 0.0003, 0.0005, 0.0014],
    "BBB-": [0.0001, 0.0001, 0.0002, 0.0004, 0.0006, 0.0013, 0.0023, 0.0108, 0.0903, 0.7287, 0.0557, 0.0201, 0.0081, 0.0036, 0.0021, 0.0015, 0.0019, 0.0023],
    "BB+": [0.0004, 0.0000, 0.0000, 0.0003, 0.0003, 0.0008, 0.0008, 0.0037, 0.0149, 0.1078, 0.6610, 0.0770, 0.0257, 0.0097, 0.0051, 0.0022, 0.0033, 0.0031],
    "BB": [0.0000, 0.0000, 0.0003, 0.0001, 0.0000, 0.0005, 0.0004, 0.0015, 0.0056, 0.0187, 0.0950, 0.6534, 0.0865, 0.0236, 0.0103, 0.0035, 0.0048, 0.0046],
    "BB-": [0.0000, 0.0000, 0.0000, 0.0001, 0.0001, 0.0001, 0.0004, 0.0009, 0.0022, 0.0032, 0.0161, 0.0940, 0.6389, 0.0860, 0.0301, 0.0079, 0.0071, 0.0092],
    "B+": [0.0000, 0.0001, 0.0000, 0.0003, 0.0000, 0.0003, 0.0006, 0.0004, 0.0005, 0.0010, 0.0029, 0.0138, 0.0823, 0.6253, 0.0947, 0.0258, 0.0178, 0.0194],
    "B": [0.0000, 0.0000, 0.0001, 0.0001, 0.0000, 0.0003, 0.0003, 0.0001, 0.0005, 0.0003, 0.0009, 0.0020, 0.0105, 0.0726, 0.6160, 0.1001, 0.0396, 0.0299],
    "B-": [0.0000, 0.0000, 0.0000, 0.0000, 0.0001, 0.0003, 0.0000, 0.0005, 0.0005, 0.0008, 0.0007, 0.0016, 0.0038, 0.0202, 0.0965, 0.5528, 0.1215, 0.0589],
    "CCC": [0.0000, 0.0000, 0.0000, 0.0000, 0.0002, 0.0000, 0.0007, 0.0004, 0.0007, 0.0004, 0.0002, 0.0013, 0.0034, 0.0085, 0.0253, 0.1004, 0.4391, 0.2655],
}


@dataclass(frozen=True)
class BondSpec:
    name: str
    rating: str
    coupon_rate: float
    face_value: float = 100.0
    recovery_rate: float = 0.0
    maturity_years: int = 5


EXERCISE_BONDS = [
    BondSpec(name="Bono_A_menos", rating="A-", coupon_rate=0.075, recovery_rate=0.45),
    BondSpec(name="Bono_BBB_menos", rating="BBB-", coupon_rate=0.090, recovery_rate=0.35),
    BondSpec(name="Bono_B_menos", rating="B-", coupon_rate=0.110, recovery_rate=0.25),
]
