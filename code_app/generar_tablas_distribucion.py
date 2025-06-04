import numpy as np
import pandas as pd
from scipy.stats import triang, beta

def generar_tablas_aleatorias_top(top_variables, n=1000):
    """
    Genera tablas de valores aleatorios para las variables top más influyentes.

    Args:
        top_variables (list): Lista de variables top con 'Celda Raíz' y 'Valor Original'.
        n (int): Número de muestras aleatorias a generar para cada variable.

    Returns:
        dict: Diccionario con DataFrames para cada variable.
    """
    tablas = {}
    distribuciones = {}
    for var in top_variables:
        nombre = f"{var['Hoja Raíz']}!{var['Celda Raíz']}"
        moda = var["Valor Original"]

        print(f"\nConfiguración para variable: {nombre} (Moda: {moda})")
        minimo = float(input(f"→ Ingrese el valor mínimo de {nombre} -> {moda}: "))
        maximo = float(input(f"→ Ingrese el valor máximo de {nombre} -> {moda}: "))

        while not (minimo <= moda <= maximo):
            print("⚠️  La moda debe estar entre mínimo y máximo.")
            minimo = float(input("→ Ingrese el valor mínimo: "))
            maximo = float(input("→ Ingrese el valor máximo: "))

        distribucion = input("→ Tipo de distribución (PERT / Normal / Triangular): ").strip().lower()
        while distribucion not in ["pert", "normal", "triangular"]:
            distribucion = input("⚠️  Distribución inválida. Ingrese (PERT / Normal / Triangular): ").strip().lower()

        # Generar muestras
        if distribucion == "normal":
            mu = moda
            sigma = (maximo - minimo) / 6  # Aprox. 99.7% de valores en rango
            muestras = np.random.normal(mu, sigma, n)
            muestras = np.clip(muestras, minimo, maximo)  # Limitar entre min y max

        elif distribucion == "triangular":
            c = (moda - minimo) / (maximo - minimo)
            muestras = triang.rvs(c, loc=minimo, scale=(maximo - minimo), size=n)

        elif distribucion == "pert":
            # Distribución PERT (modificada beta): más realista para simulaciones
            def pert_random(a, b, c, lamb=4, size=1):
                alpha = 1 + lamb * (c - a) / (b - a)
                beta_ = 1 + lamb * (b - c) / (b - a)
                samples = beta.rvs(alpha, beta_, size=size)
                return a + samples * (b - a)
            muestras = pert_random(minimo, maximo, moda, size=n)

        df = pd.DataFrame({nombre: muestras})
        tablas[nombre] = df
        distribuciones[nombre] = distribucion

    return tablas, distribuciones