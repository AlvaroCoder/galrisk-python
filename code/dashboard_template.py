import tkinter as tk
from tkinter import ttk 
from tkinter import filedialog, messagebox
from funciones_dashboard import seleccionar_archivo
import openpyxl
import os
import pandas as pd 
import string

def letra_columnda(index):
    """"""
    letras = ""
    while index >= 0:
        letras = chr()

def abrir_excel(ruta, nombre_hoja="",num_filas=400, columnas_hasta='Z'):
    # Leemos el archivo 
    df = pd.read_excel(ruta, nombre_hoja)

    columnas = list(string.ascii_uppercase[:string.ascii_uppercase.index(columnas_hasta)+1])

    # Limpiamos el visor de excel
    for i in tree.get_children():
        tree.delete(i)

    tree['columns'] = columnas
    tree['show'] = "headings"

    for col in columnas:
        tree.heading(col, text=col)
        tree.column(col, width=200)
    for i in range(1, num_filas + 1):
        empty_row = ['' for _ in columnas]
        tree.insert("", "end", text=str(i),values=empty_row)

def cargar_archivo():
    archivo_info = {"ruta":"", "nombre":""}
    ruta = seleccionar_archivo()
    if ruta:
        try:
            # Cargar archivo
            wb = openpyxl.load_workbook(ruta)
            hojas = wb.sheetnames

            # Actualizar titulo superior
            nombre_archivo = os.path.basename(ruta)
            archivo_info["nombre"] = nombre_archivo
            archivo_info["ruta"] = ruta

            label_archivo.config(text=f"📁 Archivo: {nombre_archivo}")

            # Limpiar panel de hojas
            for widget in left_frame.winfo_children():
                if isinstance(widget, ttk.Button) and widget != btn_archivo:
                    widget.destroy()
            
            for hoja in hojas:
                btn = ttk.Button(left_frame, text=hoja, width=25,
                                command=lambda h=hoja: abrir_excel(ruta, h)
                                 )
                btn.pack(padx=5, pady=2, anchor="w")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo analizar la hoja:\n{e}")

# Crear ventana principal
root = tk.Tk()
root.title("Visor de Excel")
root.geometry("800x500")
root.configure(bg="#f5f5f5")

# Estilo general
style = ttk.Style()
style.theme_use("clam")
style.configure("TButton", font=("Segoe UI", 10), padding=6)
style.configure("TLabel", font=("Segoe UI", 11))

# Frame superior (barra superior)
top_frame = tk.Frame(root, bg="white", height=100)
top_frame.pack(side="top", fill="x")

# Titulo superior
label_archivo = ttk.Label(top_frame, text="📁 Ningún archivo seleccionado", background="#ffffff", anchor="w", padding=10, font=("Segoe UI", 12, "bold"))
label_archivo.pack(side="left")

# Relleno entre titulo y boton
label_espacio = ttk.Label(top_frame, background="white")
label_espacio.pack(side="left", expand=True)

# Boton de archivo
btn_archivo = ttk.Button(top_frame, text="Seleccionar archivo", command=cargar_archivo)
btn_archivo.pack(side="right", padx=20)


# Frame lateral izquierdo (barra lateral)
left_frame = tk.Frame(root, bg="#eaeaea", width=200, bd=1, relief="solid", pady=10)
left_frame.pack(side="left", fill="y")

# Frame central (contenido principal)
center_frame = tk.Frame(root, bg="gray")
center_frame.pack(side="left", fill="both", expand=True)

tree = ttk.Treeview(center_frame)
tree.pack(fill=tk.BOTH, expand=True)

root.mainloop()