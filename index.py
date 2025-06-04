from code_app.obtener_precedentes import obtener_precedentes_completos
from code_app.filtrar_raices import filtrar_valores_raiz
from code_app.simular_impacto_raices import simular_impacto_raices
from code_app.filtrar_impactos import filtrar_mayores_impactos
from tabulate import tabulate  # pip install tabulate
from code_app.generar_tablas_distribucion import generar_tablas_aleatorias_top
from code_app.generar_graficas import generar_graficas_de_tablas

ruta_excel = "/Users/alvarofelipepupuchemorales/Desktop/Proyecto Economia/assets/PALTA HASS  BCP.xlsx"
hoja_nombre = "Costos Agricolas 01 ha"
celda_objetivo = "E123"

precedentes_totales = obtener_precedentes_completos(ruta_excel, hoja_nombre, celda_objetivo)
for p in precedentes_totales:
    print(p)

print("="*100)

raices = filtrar_valores_raiz(ruta_excel, precedentes_totales)

print("Raíces encontradas:")
for celda, valor in raices.items():
    print(f"{celda} → {valor}")

resultados = simular_impacto_raices(
    ruta=ruta_excel, 
    hoja_objetivo=hoja_nombre, 
    celda_objetivo=celda_objetivo, 
    raices_dict=raices)

print("="*100)
print("TABLA RESULTADOS")
# Mostrar tabla de resultados
if resultados:
    print(tabulate(resultados, headers="keys", floatfmt=".6f", tablefmt="fancy_grid"))
else:
    print("No se encontró ninguna raíz válida para simular.")

print("="*100)
print("IMPACTOS TOP")

top_impactos = filtrar_mayores_impactos(resultados, top_n=6)

for i, r in enumerate(top_impactos, 1):
    print(f"{i}. {r['Hoja Raíz']}!{r['Celda Raíz']} | Valor Original = {r['Valor Original']} | Variacion Absoluta = {r['Variación Absoluta']}")

tablas_simuladas, distribuciones_simuladas = generar_tablas_aleatorias_top(top_impactos, 10000)

generar_graficas_de_tablas(tablas_simuladas, distribuciones_simuladas)

