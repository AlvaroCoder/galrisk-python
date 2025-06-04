from funciones2 import simular_impacto_precedentes_2,mostrar_resultados_tabla,simular_impacto_precedentes, obtener_precedentes_recursivos, filtrar_raices_precedentes
from pprint import pprint
ruta_excel = "/Users/alvarofelipepupuchemorales/Desktop/Proyecto Economia/assets/PALTA HASS  BCP.xlsx"
hoja_nombre = "Costos Agricolas 01 ha"
celda_objetivo = "E123"

precedentes = obtener_precedentes_recursivos(ruta_excel, hoja_nombre, celda_objetivo)
pprint(precedentes)
print("="*20)
raices = filtrar_raices_precedentes(precedentes)

pprint(raices)

for r in raices:
    print(f'Celda: {r["celda"]}, Valor: {r["valor"]}, Hoja: {r["hoja"]}')


resultados = simular_impacto_precedentes_2(ruta=ruta_excel, hoja_objetivo=hoja_nombre, celda_objetivo=celda_objetivo, raices=raices)
pprint(resultados)
print(mostrar_resultados_tabla(resultados))


