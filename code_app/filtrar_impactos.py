def filtrar_mayores_impactos(resultados, top_n=6):
    """
    Filtra y ordena los resultados según la variación absoluta, de mayor a menor.

    Args:
        resultados (list): Lista de diccionarios con las variaciones calculadas.
        top_n (int): Número de resultados a retornar (por defecto 6).

    Returns:
        list: Lista de los 'top_n' elementos con mayor variación absoluta.
    """

    if not resultados:
        return []

    # Filtrar resultados válidos con variación absoluta definida
    resultados_filtrados = [
        r for r in resultados
        if r.get("Variación Absoluta") is not None
    ]
    # Ordenar de mayor a menor según variación absoluta
    resultados_ordenados = sorted(
        resultados_filtrados,
        key=lambda x: x["Variación Absoluta"],
        reverse=True
    )

    # Retornar solo los primeros 'top_n'
    return resultados_ordenados[:top_n]