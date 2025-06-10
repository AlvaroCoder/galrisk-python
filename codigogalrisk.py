import xlwings as xw
from code_app.obtener_precedentes import obtener_precedentes_completos
from code_app.filtrar_raices import filtrar_valores_raiz
from tabulate import tabulate  # pip install tabulate
from code_app.generar_tablas_distribucion import generar_tablas_aleatorias_top

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
from code_app.filtrar_impactos import filtrar_mayores_impactos

top_impactos = filtrar_mayores_impactos(resultados, top_n=6)
print(tabulate(top_impactos, headers="keys", floatfmt=".6f", tablefmt="fancy_grid"))

print("="*100)
print("Generando tablas de simulación para los impactos top...")
tablas_simuladas, distribuciones_simuladas = generar_tablas_aleatorias_top(top_impactos, 100000)
print("Tablas de simulación generadas con éxito.")

wb = xw.Book(ruta_excel)

"""
for tabla in list(tablas_simuladas.keys())[:2]:
    hoja, celda_ref = tabla.split("!")
    resultados_vector = []
    print(f"Tabla: {tabla}"+"="*10)
    for valor in tablas_simuladas[tabla][:100].values:
        valor_inicial = wb.sheets[hoja].range(celda_ref).value
        wb.sheets[hoja].range(celda_ref).value = valor[0]
        resultado = wb.sheets[hoja_nombre].range(celda_objetivo).value
        resultados_vector.append({
            tabla: valor[0],
            f"resultado objetivo {celda_ref}": resultado,
        })
        wb.sheets[hoja].range(celda_ref).value = valor_inicial
    print(tabulate(resultados_vector, headers="keys", floatfmt=".6f", tablefmt="fancy_grid"))

wb.close()  # Cerrar el libro de Excel
"""
import numpy as np

union_valores = [tablas_simuladas[tabla].values for tabla in tablas_simuladas.keys()]
matriz_simulada = np.hstack(union_valores)


llaves_tablas = list(tablas_simuladas.keys())
celda_refs = [llave.split("!") for llave in llaves_tablas]  # lista de [hoja, celda]
celda_columnas = [celda for _, celda in celda_refs]

lista_resultados = []

nombre_hoja_simulacion = "Simulación de Impactos"
if nombre_hoja_simulacion in [s.name for s in wb.sheets]:
    wb.sheets[nombre_hoja_simulacion].delete()


wb.sheets.add(nombre_hoja_simulacion)  # Agregar una nueva hoja para los resultados
# Escribir encabezados en la hoja de simulación

letras_encabezado = []
hoja_simulacion = wb.sheets[nombre_hoja_simulacion]
for letra in range(len(llaves_tablas)):
    letras_encabezado.append(f"{chr(65 + letra)}{1}")  # A, B, C, D, E, F, ...

for indice,letra in enumerate(letras_encabezado):
    hoja_simulacion.range(letra).value = celda_columnas[indice]  

hoja_simulacion.range("A2").value = matriz_simulada



print("Resultados de la simulación de 100 iteraciones:")
print(tabulate(lista_resultados, headers="keys", floatfmt=".6f", tablefmt="fancy_grid"))


"""
for indice, fila in enumerate(matriz_simulada):
    if indice >= 1000:
        break
    for indice_columna in range(len(fila)):
        llave = llaves_tablas[indice_columna]
        hoja, celda_ref = llave.split("!")
        wb.sheets[hoja].range(celda_ref).value = fila[indice_columna]
        
    resultado = wb.sheets[hoja_nombre].range(celda_objetivo).value
    lista_resultados.append({

        llaves_tablas[0].split("!")[1] : fila[0],
        llaves_tablas[1].split("!")[1] : fila[1],
        llaves_tablas[2].split("!")[1] : fila[2],
        llaves_tablas[3].split("!")[1] : fila[3],
        llaves_tablas[4].split("!")[1] : fila[4],
        llaves_tablas[5].split("!")[1] : fila[5],
        "resultado" : resultado,
    })
"""

"""
for i, fila in enumerate(matriz_simulada[:1000]):  # Limitamos directamente aquí
    # Escribir valores en las 6 celdas correspondientes
    for (hoja, celda), valor in zip(celda_refs, fila):
        wb.sheets[hoja].range(celda).value = valor

    # Obtener resultado de la celda objetivo
    resultado = wb.sheets[hoja_nombre].range(celda_objetivo).value

    # Armar diccionario de fila de resultados
    resultado_dict = {col: val for col, val in zip(celda_columnas, fila)}
    resultado_dict["resultado"] = resultado
    lista_resultados.append(resultado_dict)


wb.close()  # Cerrar el libro de Excel
"""