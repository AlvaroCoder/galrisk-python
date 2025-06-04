import openpyxl
from pathlib import Path

def filtrar_valores_raiz(ruta_excel, celdas_precedentes):
    ruta_excel = Path(ruta_excel)
    if not ruta_excel.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_excel}")

    wb = openpyxl.load_workbook(ruta_excel, data_only=False)
    raices = {}

    for referencia in celdas_precedentes:
        try:
            hoja_nombre, celda_ref = referencia.split("!")
            hoja = wb[hoja_nombre]
            celda = hoja[celda_ref]

            valor = celda.value
            # Si no es fórmula y tiene algún valor → lo consideramos raíz
            if not (isinstance(valor, str) and valor.startswith("=")) and valor is not None:
                raices[referencia] = valor

        except Exception as e:
            print(f"Error al procesar {referencia}: {e}")

    return raices