import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import string

root = tk.Tk()
root.title("Visor de Excel estilo Excel")

# ---------------- Frame superior ----------------
frame_top = tk.Frame(root)
frame_top.pack(fill="x", pady=5, padx=5)
tk.Label(frame_top, text="Selecciona un archivo Excel:").pack(side="left")
tk.Button(frame_top, text="Cargar archivo", command=lambda: cargar_hojas()).pack(side="left", padx=10)

# ---------------- Frame izquierdo ----------------
frame_left = tk.Frame(root)
frame_left.pack(side="left", fill="y", padx=5)
tk.Label(frame_left, text="Hojas del archivo").pack()
listbox_hojas = tk.Listbox(frame_left, width=25)
listbox_hojas.pack(fill="y")
listbox_hojas.bind("<<ListboxSelect>>", lambda event: mostrar_hoja())

# ---------------- Frame principal para tabla ----------------
frame_main = tk.Frame(root)
frame_main.pack(side="left", fill="both", expand=True)

scroll_y = tk.Scrollbar(frame_main, orient="vertical")
scroll_y.pack(side="right", fill="y")

scroll_x = tk.Scrollbar(root, orient="horizontal")
scroll_x.pack(side="bottom", fill="x")

# Treeview con columnas A-Z
columnas_excel = list(string.ascii_uppercase)  # A-Z
tree = ttk.Treeview(frame_main, columns=columnas_excel, show="headings",
                    yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
tree.pack(fill="both", expand=True)

scroll_y.config(command=tree.yview)
scroll_x.config(command=tree.xview)

# Estilo tipo Excel
style = ttk.Style()
style.theme_use("default")
style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background="#e6e6e6")
style.configure("Treeview", font=("Segoe UI", 10), rowheight=25, borderwidth=1, relief="solid")

# ---------------- Funciones ----------------
archivo_actual = None
excel_data = None

def cargar_hojas():
    global archivo_actual, excel_data
    archivo_actual = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx *.xls")])
    if archivo_actual:
        excel_data = pd.read_excel(archivo_actual, sheet_name=None)  # Lee todas las hojas
        listbox_hojas.delete(0, tk.END)
        for nombre in excel_data.keys():
            listbox_hojas.insert(tk.END, nombre)

def mostrar_hoja():
    if not excel_data:
        return
    seleccion = listbox_hojas.curselection()
    if not seleccion:
        return
    hoja = listbox_hojas.get(seleccion[0])
    df = excel_data[hoja]

    # Limpiar treeview
    for item in tree.get_children():
        tree.delete(item)
    

    # Mostrar columnas A-Z
    for col in columnas_excel:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor="center", stretch=False)
    
    tree.heading("#0", text="#")  # Número de fila
    tree.column("#0", width=50, anchor='center')
    # Mostrar filas numeradas
    for i in range(400):  # 1 a 400
        values = []
        for j in range(26):  # A-Z (0-25)
            try:
                val = df.iat[i, j]
            except IndexError:
                val = ""
            values.append(val)
        tree.insert("", "end", values=values)

# ---------------- Iniciar App ----------------
root.mainloop()