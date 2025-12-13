# src/utils/archivos.py
"""
Utilidades para manejo de archivos
"""

import os
from pathlib import Path


def buscar_archivo_recursivo(nombre_archivo, ruta_inicio=None):
    """
    Busca un archivo de manera recursiva desde una ruta.
    
    Args:
        nombre_archivo (str): Nombre del archivo a buscar
        ruta_inicio (Path): Ruta donde comenzar la búsqueda
    
    Returns:
        Path: Ruta del archivo encontrado o None
    """
    if ruta_inicio is None:
        ruta_inicio = Path.cwd()
    else:
        ruta_inicio = Path(ruta_inicio)
    
    for archivo in ruta_inicio.rglob(nombre_archivo):
        return archivo
    
    return None


def obtener_ruta_plantilla(ubicaciones_alternativas=None):
    """
    Obtiene la ruta de la plantilla desde múltiples ubicaciones.
    
    Args:
        ubicaciones_alternativas (list): Ubicaciones adicionales donde buscar
    
    Returns:
        Path: Ruta de la plantilla o None
    """
    nombre = 'plantilla5852.xlsx'
    
    # Ubicaciones por defecto
    ubicaciones = [
        Path.cwd() / nombre,
        Path.cwd() / 'src' / 'plantillas' / nombre,
        Path.cwd() / 'src' / nombre,
        Path.cwd() / 'resources' / nombre,
        Path.cwd() / 'data' / nombre,
    ]
    
    # Agregar ubicaciones alternativas
    if ubicaciones_alternativas:
        ubicaciones.extend([Path(u) / nombre for u in ubicaciones_alternativas])
    
    # Buscar en ubicaciones
    for ruta in ubicaciones:
        if ruta.exists():
            return ruta
    
    # Búsqueda recursiva
    return buscar_archivo_recursivo(nombre)


def validar_archivo_excel(ruta):
    """
    Valida que un archivo sea un Excel válido.
    
    Args:
        ruta (str): Ruta del archivo
    
    Returns:
        bool: True si es válido, False en caso contrario
    """
    ruta = Path(ruta)
    
    # Verificar existencia
    if not ruta.exists():
        return False
    
    # Verificar extensión
    if ruta.suffix.lower() not in ['.xlsx', '.xls']:
        return False
    
    # Verificar que sea un archivo (no directorio)
    if not ruta.is_file():
        return False
    
    return True


def crear_directorio_si_no_existe(ruta):
    """Crea un directorio si no existe"""
    Path(ruta).mkdir(parents=True, exist_ok=True)
