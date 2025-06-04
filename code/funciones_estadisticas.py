import numpy as np

def generar_aleatorio(minimo, maximo, modelo="uniforme", media=None, desviacion=None):
    """
    Genera un número aleatorio entre minimo y maximo basado en un modelo estadístico.

    Args:
        minimo (float): Valor mínimo del rango.
        maximo (float): Valor máximo del rango.
        modelo (str): Modelo estadístico a usar. Puede ser: "uniforme", "normal", "triangular", etc.
        media (float): Media (para distribuciones que la usen como la normal).
        desviacion (float): Desviación estándar (solo para modelo normal).

    Returns:
        float: Número aleatorio generado.
    """
    if modelo == "uniforme":
        return np.random.uniform(minimo, maximo)

    elif modelo == "normal":
        if media is None:
            media = (minimo + maximo) / 2
        if desviacion is None:
            desviacion = (maximo - minimo) / 6  # 99.7% dentro de 3σ
        valor = np.random.normal(media, desviacion)
        return np.clip(valor, minimo, maximo)

    elif modelo == "triangular":
        moda = media if media is not None else (minimo + maximo) / 2
        return np.random.triangular(minimo, moda, maximo)

    elif modelo == "beta":
        # Distribución beta para valores entre [0, 1], escalamos después
        a, b = 2, 5  # Puedes parametrizar estos
        valor = np.random.beta(a, b)
        return minimo + valor * (maximo - minimo)

    else:
        raise ValueError(f"Modelo '{modelo}' no soportado. Usa 'uniforme', 'normal', 'triangular' o 'beta'.")



