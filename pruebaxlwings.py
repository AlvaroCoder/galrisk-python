import xlwings as xw

data_prueba = [
    24.76940368445477,
    24.766220319071348,
    24.9689227006509,
    25.27792474052166,
    24.725574246063168,
    25.480831137175375,
    25.143797591078226,
    25.61041031070497,
    24.87838976593049,
    24.49826915811586
]

ruta_excel = "/Users/alvarofelipepupuchemorales/Desktop/Proyecto Economia/assets/PALTA HASS  BCP.xlsx"
hoja_nombre = "Costos Agricolas 01 ha"
celda_objetivo = "E123"

wb = xw.Book(ruta_excel)

for data in data_prueba:
    wb.sheets[hoja_nombre].range('E45').value = data
    
    resultado = wb.sheets[hoja_nombre].range(celda_objetivo).value

    print(f"Valor actualizado en {celda_objetivo}: resultado {resultado}")