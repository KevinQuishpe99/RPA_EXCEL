# __init__.py para módulo utils
from .busqueda import (
    buscar_plantilla,
    extraer_numero_poliza,
    obtener_hojas_validas,
    validar_patron_regex,
    limpiar_nombre_archivo,
)

from .polizas import (
    detectar_poliza_desde_plantilla,
    generar_nombre_archivo,
    obtener_hojas_procesables,
)

__all__ = [
    'buscar_plantilla',
    'extraer_numero_poliza',
    'obtener_hojas_validas',
    'validar_patron_regex',
    'limpiar_nombre_archivo',
    'detectar_poliza_desde_plantilla',
    'generar_nombre_archivo',
    'obtener_hojas_procesables',
]
