import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from code_app.simular_impacto_raices import simular_impacto_raices
from code_app.filtrar_impactos import filtrar_mayores_impactos as filtrar_top_variaciones
# Puedes importar más funciones según necesidad

class MainPanel:
    def __init__(self, root, usuario):
        self.root = root
        self.usuario = usuario
        self.df_excel = None

        self.frame_main = tk.Frame(self.root)
        self.frame_main.pack(fill="both", expand=True)

        # Topbar
        topbar = tk.Frame(self.frame_main, bg="#1e90ff", height=50)
        topbar.pack(fill="x")
        tk.Label(topbar, text=f"Usuario: {usuario}", bg="#1e90ff", fg="white", font=("Arial", 12)).pack(side="right", padx=10)

        # Panel de bienvenida
        bienvenida = tk.Label(self.frame_main, text=f"Bienvenido, {usuario}", font=("Arial", 16))
        bienvenida.pack(pady=20)

        # Botón para subir Excel
        boton_subir = tk.Button(self.frame_main, text="Subir archivo Excel", command=self.subir_archivo)
        boton_subir.pack(pady=10)

        # Botón para calcular raíces
        self.boton_calcular = tk.Button(self.frame_main, text="Calcular raíces", state="disabled", command=self.calcular_raices)
        self.boton_calcular.pack(pady=10)

    def subir_archivo(self):
        archivo = filedialog.askopenfilename(
            title="Selecciona un archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx")]
        )
        if archivo:
            try:
                self.df_excel = pd.read_excel(archivo)
                messagebox.showinfo("Éxito", "Archivo cargado correctamente.")
                self.boton_calcular.config(state="normal")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar el archivo: {e}")

    def calcular_raices(self):
        if self.df_excel is not None:
            try:
                resultado = simular_impacto_raices(self.df_excel)
                top_resultado = filtrar_top_variaciones(resultado)
                messagebox.showinfo("Análisis completo", f"Top variables con mayor impacto:\n{top_resultado['Variable'].tolist()}")
                # Aquí puedes llamar a nuevas pantallas, gráficos, etc.
            except Exception as e:
                messagebox.showerror("Error en análisis", str(e))