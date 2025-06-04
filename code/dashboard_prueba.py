import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd

class ExcelViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Visor de Excel")
        self.root.geometry("800x500")

        # Panel superior
        top_frame = tk.Frame(root, bg="skyblue", height=50)
        top_frame.pack(side="top", fill="x")

        tk.Label(top_frame, text="Visor de Excel", bg="skyblue", font=("Arial", 12, "bold")).pack(side="left", padx=10)
        tk.Button(top_frame, text="Cargar archivo", command=self.cargar_archivo).pack(side="right", padx=10)

        # Panel izquierdo (hojas)
        self.left_frame = tk.Frame(root, bg="lightblue", width=150)
        self.left_frame.pack(side="left", fill="y")

        self.hojas_listbox = tk.Listbox(self.left_frame)
        self.hojas_listbox.pack(fill="both", expand=True, padx=5, pady=5)
        self.hojas_listbox.bind("<<ListboxSelect>>", self.mostrar_hoja)

        # Panel central (datos)
        self.center_frame = tk.Frame(root, bg="gray")
        self.center_frame.pack(side="left", fill="both", expand=True)

        self.tree = ttk.Treeview(self.center_frame)
        self.tree.pack(fill="both", expand=True)

        self.excel_data = {}  # Aquí se almacenan las hojas

    def cargar_archivo(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        self.excel_data = pd.read_excel(file_path, sheet_name=None)  # Carga todas las hojas
        self.hojas_listbox.delete(0, tk.END)

        for hoja in self.excel_data.keys():
            self.hojas_listbox.insert(tk.END, hoja)

    def mostrar_hoja(self, event):
        selection = self.hojas_listbox.curselection()
        if not selection:
            return

        hoja_nombre = self.hojas_listbox.get(selection[0])
        df = self.excel_data[hoja_nombre]

        # Limpiar el Treeview
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for _, row in df.iterrows():
            self.tree.insert("", tk.END, values=list(row))

# Ejecutar la aplicación
root = tk.Tk()
app = ExcelViewerApp(root)
root.mainloop()