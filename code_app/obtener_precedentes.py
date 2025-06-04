import openpyxl
import re
from pathlib import Path

def obtener_precedentes_completos(ruta_excel, hoja_inicial, celda_inicial):
    ruta_excel = Path(ruta_excel)
    if not ruta_excel.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {ruta_excel}")

    wb = openpyxl.load_workbook(ruta_excel, data_only=False)
    visitados = set()
    precedentes = set()

    def analizar(hoja_nombre, celda_ref):
        clave = f"{hoja_nombre}!{celda_ref}"
        if clave in visitados:
            return
        visitados.add(clave)

        if hoja_nombre not in wb.sheetnames:
            return

        hoja = wb[hoja_nombre]
        celda = hoja[celda_ref]
        valor = celda.value

        if isinstance(valor, str) and valor.startswith('='):
            # Buscar referencias en la fórmula
            patron = r"(?:'([^']+)'!)?(\$?[A-Z]{1,3}\$?[0-9]{1,7})"
            referencias = re.findall(patron, valor)
            for hoja_ref, celda_obj in referencias:
                hoja_destino = hoja_ref if hoja_ref else hoja_nombre
                celda_destino = celda_obj.upper()
                clave_destino = f"{hoja_destino}!{celda_destino}"
                precedentes.add(clave_destino)
                analizar(hoja_destino, celda_destino)
        else:
            # Es un valor literal (número, texto, etc.)
            precedentes.add(clave)

    analizar(hoja_inicial, celda_inicial.upper())
    precedentes.discard(f"{hoja_inicial}!{celda_inicial.upper()}")  # excluir la celda objetivo
    return sorted(precedentes)