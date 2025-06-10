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
tablas_simuladas, distribuciones_simuladas = generar_tablas_aleatorias_top(top_impactos, 10000)
print("Tablas de simulación generadas con éxito.")

wb = xw.Book(ruta_excel)

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
