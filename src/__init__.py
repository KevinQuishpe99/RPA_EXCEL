# __init__.py para módulo src
from .config import (
    CONFIGURACION_POLIZAS,
    CONFIG_SISTEMA,
    ENCABEZADOS_CRITICOS,
    PALABRAS_CLAVE_TOTALES,
    TRANSFORMACIONES,
    MESES_ESPANOL,
)

from .core import TransformadorDatos

from .utils import (
    buscar_plantilla,
    extraer_numero_poliza,
    detectar_poliza_desde_plantilla,
    generar_nombre_archivo,
)

__all__ = [
    # Config
    'CONFIGURACION_POLIZAS',
    'CONFIG_SISTEMA',
    'ENCABEZADOS_CRITICOS',
    'PALABRAS_CLAVE_TOTALES',
    'TRANSFORMACIONES',
    'MESES_ESPANOL',
    
    # Core
    'TransformadorDatos',
    
    # Utils
    'buscar_plantilla',
    'extraer_numero_poliza',
    'detectar_poliza_desde_plantilla',
    'generar_nombre_archivo',
]
