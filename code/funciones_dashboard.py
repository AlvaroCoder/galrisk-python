from tkinter import filedialog, messagebox

def seleccionar_archivo():
    return filedialog.askopenfilename(
            filetypes=[("Archivos Excel", "*.xlsx")],
            title="Selecciona un archivo Excel")
