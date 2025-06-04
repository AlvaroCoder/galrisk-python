import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import string

def letra_columna(index):
    """Convierte índice numérico a letra como Excel (A-Z, AA, AB, etc.)."""
    letras = ''
    while index >= 0:
        letras = chr(index % 26 + ord('A')) + letras
        index = index // 26 - 1
    return letras

def cargar_hojas():
    global ruta_excel
    ruta_excel = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if ruta_excel:
        xls = pd.ExcelFile(ruta_excel)
        listbox_hojas.delete(0, tk.END)
        for hoja in xls.sheet_names:
            listbox_hojas.insert(tk.END, hoja)

def cargar_hoja(nombre_hoja):
    df = pd.read_excel(ruta_excel, sheet_name=nombre_hoja)

    # Limpiar tabla
    for i in tree.get_children():
        tree.delete(i)
    tree["columns"] = []

    num_columnas = 26 # A-Z máximo
    columnas_excel = [letra_columna(i) for i in range(num_columnas)]
    print(columnas_excel)

    tree["columns"] = columnas_excel
    tree["show"] = "tree headings"  # Muestra columna #0 como índice (números de fila)

    # Configurar encabezados
    for i, col in enumerate(columnas_excel):
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor='center', stretch=False)
    tree.heading("#0", text="#")  # Número de fila
    tree.column("#0", width=80, anchor='center')

    # Rellenar celdas
    for i in range(0, max(len(df), 400)):
        valores = []
        for j in range(num_columnas):
            try:
                val = df.iloc[i, j]
            except:
                val = ""
            valores.append("" if pd.isna(val) else val)
        tree.insert("", "end", text=str(i+1), values=valores)

def on_seleccionar_hoja(event):
    seleccion = listbox_hojas.curselection()
    if seleccion:
        hoja = listbox_hojas.get(seleccion[0])
        cargar_hoja(hoja)

# GUI
root = tk.Tk()
root.title("Visor tipo Excel")
root.geometry("1200x700")

# Frame superior (botón cargar)
frame_top = tk.Frame(root)
frame_top.pack(fill="x", pady=5, padx=5)
tk.Label(frame_top, text="Selecciona un archivo Excel:").pack(side="left")
tk.Button(frame_top, text="Cargar archivo", command=cargar_hojas).pack(side="left", padx=10)

# Frame lateral izquierdo (hojas)
frame_left = tk.Frame(root)
frame_left.pack(side="left", fill="y", padx=5)
tk.Label(frame_left, text="Hojas del archivo").pack()
listbox_hojas = tk.Listbox(frame_left, width=25)
listbox_hojas.pack(fill="y")
listbox_hojas.bind("<<ListboxSelect>>", on_seleccionar_hoja)

# Frame principal para tabla
frame_main = tk.Frame(root)
frame_main.pack(side="left", fill="both", expand=True)

scroll_y = tk.Scrollbar(frame_main, orient="vertical")
scroll_x = tk.Scrollbar(frame_main, orient="horizontal")

tree = ttk.Treeview(frame_main, yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

scroll_y.config(command=tree.yview)
scroll_x.config(command=tree.xview)

scroll_y.pack(side="right", fill="y")
scroll_x.pack(side="bottom", fill="x")
tree.pack(fill="both", expand=True)

root.mainloop()