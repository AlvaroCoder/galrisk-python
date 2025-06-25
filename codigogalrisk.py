import xlwings as xw
from code_app.obtener_precedentes import obtener_precedentes_completos
from code_app.filtrar_raices import filtrar_valores_raiz
from tabulate import tabulate  # pip install tabulate
from code_app.generar_tablas_distribucion import generar_tablas_aleatorias_top
from code_app.filtrar_impactos import filtrar_mayores_impactos

ruta_excel = "/Users/alvarofelipepupuchemorales/Desktop/Proyecto Economia/assets/PALTA HASS  BCP.xlsx"
hoja_nombre = "Costos Agricolas 01 ha"
celda_objetivo = "E123"

precedentes_totales = obtener_precedentes_completos(ruta_excel, hoja_nombre, celda_objetivo)
raices = filtrar_valores_raiz(ruta_excel, precedentes_totales)

wb = xw.Book(ruta_excel)
valor_objetivo_original = wb.sheets[hoja_nombre].range(celda_objetivo).value
resultados = []
for celda, valor in raices.items():
    nuevo_valor = valor * 1.1  # Aumentar el valor en un 10% como ejemplo
    
    hoja, celda_ref = celda.split("!")
    hoja_obj = wb.sheets[hoja]
    hoja_obj.range(celda_ref).value = nuevo_valor
    resultado = wb.sheets[hoja_nombre].range(celda_objetivo).value
    

    variacion = valor_objetivo_original - resultado if isinstance(resultado, (int, float)) and isinstance(valor_objetivo_original, (int, float)) else None
    resultados.append({
        "Hoja Raíz": hoja,
        "Celda Raíz": celda_ref,
        "Valor Original": valor,
        "Valor Modificado": nuevo_valor,
        "Valor Objetivo Original": valor_objetivo_original,
        "Valor Objetivo Nuevo": resultado,
        "Variación": variacion,
        "Variación Absoluta": abs(variacion) if variacion is not None else None
    })

    print(f"Hoja: {hoja}, Celda: {celda_ref}, Valor Original: {valor}, Nuevo Valor: {resultado}")

    # Restaurar el valor original
    hoja_obj.range(celda_ref).value = valor


print("="*100)
print("Resultados de la simulación:")
if resultados:
   print(tabulate(resultados, headers="keys", floatfmt=".6f", tablefmt="fancy_grid"))
else:
    print("No se encontró ninguna raíz válida para simular.")

wb.close()
# Cerrar el libro de Excel

print("="*100)
print("IMPACTOS TOP")
# Filtrar por los mayores impactos

top_impactos = filtrar_mayores_impactos(resultados, top_n=6)
print(tabulate(top_impactos, headers="keys", floatfmt=".6f", tablefmt="fancy_grid"))

print("="*100)
print("Generando tablas de simulación para los impactos top...")

numero_simulaciones = 100
tablas_simuladas, distribuciones_simuladas = generar_tablas_aleatorias_top(top_impactos, numero_simulaciones)

print("Tablas de simulación generadas con éxito.")

wb = xw.Book(ruta_excel)

import numpy as np

union_valores = [tablas_simuladas[tabla].values for tabla in tablas_simuladas.keys()]
matriz_simulada = np.hstack(union_valores)


llaves_tablas = list(tablas_simuladas.keys())
celda_refs = [llave.split("!") for llave in llaves_tablas]  # lista de [hoja, celda]

lista_resultados_van = []

for i, fila in enumerate(matriz_simulada):
    cabeceras = {}
    for (hoja, celda), valor in zip(celda_refs, fila):
        wb.sheets[hoja].range(celda).value = valor
        cabeceras[celda] = valor

    resultado = wb.sheets[hoja_nombre].range(celda_objetivo).value
    cabeceras["VAN"] = resultado
    lista_resultados_van.append(cabeceras)

wb.close()

print(f"Resultados de la simulación de {numero_simulaciones} iteraciones:")
print(tabulate(lista_resultados_van, headers="keys", floatfmt=".6f", tablefmt="fancy_grid"))
