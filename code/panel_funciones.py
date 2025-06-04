import tkinter as tk   
from tkinter import filedialog, messagebox
from tkinter import ttk
import openpyxl
import os

from funciones import obtener_precedentes_con_valor

def crear_dashboard_excel():
    archivo_info = {"ruta" : "","nombre":""}
    def seleccionar_archivo():
        ruta = filedialog.askopenfilename(
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Selecciona un archivo Excel"
        )

        if ruta:
            try:
                # Cargar archivo
                wb = openpyxl.load_workbook(ruta)
                hojas = wb.sheetnames

                # Actualizar título superior
                nombre_archivo = os.path.basename(ruta)
                archivo_info["ruta"] = ruta
                archivo_info["nombre"] = nombre_archivo                
                label_archivo.config(text=f"📁 Archivo: {nombre_archivo}")

                # Limpiar panel de hojas
                for widget in frame_hojas.winfo_children():
                    if isinstance(widget, ttk.Button) and widget != btn_archivo:
                        widget.destroy()

                # Mostrar hojas en el panel izquierdo
                for hoja in hojas:
                    btn = ttk.Button(frame_hojas, text=hoja, width=25,
                                     command=lambda h=hoja: mostrar_precedentes(h))
                    btn.pack(pady=2, padx=5, anchor="w")

                limpiar_panel_contenido()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    def mostrar_precedentes(nombre_hoja=""):
        limpiar_panel_contenido()
        ruta = archivo_info.get("ruta")
        if not ruta:
            return
    
        try:
            precedentes = obtener_precedentes_con_valor(ruta, nombre_hoja)
            if not precedentes:
                tk.Label(frame_contenido, text="No se encontraron precedentes.",
                         bg="#ffffff", font=("Segoe UI", 11)).pack(pady=10)
                return

            # Título de hoja seleccionada
            tk.Label(frame_contenido, text=f"📄 Hoja: {nombre_hoja}", bg="#ffffff",
                     font=("Segoe UI", 13, "bold")).pack(pady=10)

            for p in precedentes:
                print(p['valor'])
                texto = f"📌 Celda: {p['celda']} | Fórmula: {p['formula']} | Referencias: {', '.join(p['referencias'])} | Valor : {p['valor']}"
                tk.Label(frame_contenido, text=texto, bg="#ffffff", anchor="w",
                         justify="left", wraplength=500).pack(fill="x", padx=20, pady=4)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo analizar la hoja:\n{e}")

        return
    
    def limpiar_panel_contenido():
        for widget in frame_contenido.winfo_children():
            widget.destroy()
    # Crear ventana principal
    root = tk.Tk()
    root.title("Dashboard de Excel")
    root.geometry("800x500")
    root.configure(bg="#f5f5f5")

    # Estilo general
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TButton", font=("Segoe UI", 10), padding=6)
    style.configure("TLabel", font=("Segoe UI", 11))

    # --- Título superior ---
    label_archivo = ttk.Label(root, text="📁 Ningún archivo seleccionado", background="#ffffff", anchor="w", padding=10, font=("Segoe UI", 12, "bold"))
    label_archivo.pack(fill="x", side="top")

    # --- Contenedor principal ---
    main_frame = tk.Frame(root, bg="#f5f5f5")
    main_frame.pack(fill="both", expand=True)

    # --- Panel lateral (hojas) ---
    frame_hojas = tk.Frame(main_frame, width=200, bg="#eaeaea", bd=1, relief="solid")
    frame_hojas.pack(side="left", fill="y")

    # --- Botón para abrir archivo ---
    btn_archivo = ttk.Button(frame_hojas, text="📂 Seleccionar archivo", command=seleccionar_archivo)
    btn_archivo.pack(pady=10, padx=5)

    # --- Panel contenido con Scroll ---
    contenido_canvas = tk.Canvas(main_frame, bg="#ffffff", bd=1, relief="solid", highlightthickness=0)
    contenido_scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=contenido_canvas.yview)
    contenido_scrollbar.pack(side="right", fill="y")

    contenido_canvas.configure(yscrollcommand=contenido_scrollbar.set)
    contenido_canvas.pack(side="right", fill="both", expand=True)

    # --- Panel principal (contenido) ---
    frame_contenido = tk.Frame(contenido_canvas, bg="#ffffff")
    canvas_window = contenido_canvas.create_window((0, 0), window=frame_contenido, anchor="nw")
    
    # Función para redimensionar automáticamente el scroll
    def actualizar_scroll(event):
        contenido_canvas.configure(scrollregion=contenido_canvas.bbox("all"))

    frame_contenido.bind("<Configure>", actualizar_scroll)

    # Habilitar scroll con la rueda del mouse
    def on_mousewheel(event):
        contenido_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    contenido_canvas.bind_all("<MouseWheel>", on_mousewheel)

    # Aquí puedes seguir agregando widgets al frame_contenido

    root.mainloop()


crear_dashboard_excel()