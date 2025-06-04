import matplotlib.pyplot as plt

def generar_graficas_de_tablas(tablas, distribuciones, bins=30):
    """
    Genera histogramas para los dataframes de variables top.

    Args:
        tablas (dict): Diccionario con nombres de variables como claves y DataFrames como valores.
        bins (int): Número de bins para el histograma.
    """
    for nombre_variable, df in tablas.items():
        plt.figure(figsize=(8, 4))
        plt.hist(df[nombre_variable], bins=bins, alpha=0.7, color='skyblue', edgecolor='black', density=True)
        plt.title(f"Distribución simulada - {nombre_variable} - {distribuciones[nombre_variable].title()}")
        plt.xlabel("Valor")
        plt.ylabel("Densidad")
        plt.grid(True)
        plt.tight_layout()
        plt.show()