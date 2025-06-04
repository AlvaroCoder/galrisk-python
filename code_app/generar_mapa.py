import re
from openpyxl import load_workbook

def npv(tasa, flujos):
    return sum(f / (1 + tasa) ** i for i, f in enumerate(flujos, start=1))

def extraer_referencias(formula):
    """Extrae referencias de celdas de una fórmula de Excel."""
    return re.findall(r'\b[A-Z]{1,3}[0-9]{1,7}\b', formula)

def traducir_formula(formula, referencias):
    """Convierte fórmula Excel a una fórmula Python usando diccionario `r`."""
    formula_traducida = formula.replace("=", "")  # quitar =
    formula_traducida = formula_traducida.upper()
    formula_traducida = formula_traducida.replace("NPV", "npv")

    for ref in referencias:
        formula_traducida = re.sub(rf'\b{ref}\b', f"r['{ref}']", formula_traducida)
    
    return formula_traducida

def generar_mapa_precedentes(celda_objetivo, hoja):
    """Genera funciones para cada celda desde las raíces hasta la celda objetivo."""
    mapa = {}
    visitados = set()

    def construir(celda):
        if celda in visitados:
            return  # Evitar loops
        visitados.add(celda)

        celda_obj = hoja[celda]

        if celda_obj.data_type != 'f':  # raíz
            mapa[celda] = lambda r, val=celda_obj.value: val
            return

        formula = celda_obj.value
        referencias = extraer_referencias(formula)
        for ref in referencias:
            construir(ref)  # Recursión

        try:
            codigo = traducir_formula(formula, referencias)
            # Crear función lambda segura
            mapa[celda] = eval(f"lambda r: {codigo}", {"npv": npv})
        except Exception as e:
            print(f"Error procesando {celda}: {e}")
            mapa[celda] = lambda r: None

    construir(celda_objetivo)
    return mapa