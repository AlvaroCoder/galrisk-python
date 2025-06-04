import tkinter as tk
from tkinter import messagebox
from interfaz_economia.main_panel import MainPanel

class LoginApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inicio de Sesión")
        self.usuario = None

        self.frame_login = tk.Frame(root, padx=20, pady=20)
        self.frame_login.pack()

        tk.Label(self.frame_login, text="Nombre de usuario:").grid(row=0, column=0)
        self.entry_usuario = tk.Entry(self.frame_login)
        self.entry_usuario.grid(row=0, column=1)

        tk.Label(self.frame_login, text="Contraseña:").grid(row=1, column=0)
        self.entry_password = tk.Entry(self.frame_login, show="*")
        self.entry_password.grid(row=1, column=1)

        self.btn_login = tk.Button(self.frame_login, text="Iniciar sesión", command=self.validar_login)
        self.btn_login.grid(row=2, columnspan=2, pady=10)

    def validar_login(self):
        nombre = self.entry_usuario.get()

        if not nombre.strip():
            messagebox.showerror("Error", "Debes ingresar un nombre de usuario.")
            return

        self.usuario = nombre
        self.frame_login.destroy()
        MainPanel(self.root, self.usuario)