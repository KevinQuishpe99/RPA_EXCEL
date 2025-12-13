# src/modelo/poliza.py
"""
Modelo de Póliza
Maneja toda la lógica relacionada con pólizas
"""

class Poliza:
    """Representa una póliza del sistema"""
    
    def __init__(self, tipo, config):
        self.tipo = tipo
        self.numero = None
        self.nombre_hoja = None
        self.config = config
    
    def establecer_numero(self, numero):
        """Establece el número de póliza"""
        self.numero = numero
    
    def establecer_nombre_hoja(self, nombre):
        """Establece el nombre de la hoja"""
        self.nombre_hoja = nombre
    
    def obtener_prefijo(self):
        """Retorna el prefijo de la póliza"""
        return self.config.get('prefijo', 'Póliza')
    
    def obtener_nombre_archivo(self):
        """Retorna el nombre base para el archivo"""
        return self.config.get('nombre_archivo', 'Facturación')
    
    def obtener_descripcion(self):
        """Retorna descripción de la póliza"""
        return self.config.get('descripcion', 'Sin descripción')
    
    def obtener_patrones(self):
        """Retorna patrones de búsqueda"""
        return self.config.get('patrones_hoja', [])
    
    def a_dict(self):
        """Convierte póliza a diccionario"""
        return {
            'tipo': self.tipo,
            'numero': self.numero,
            'nombre_hoja': self.nombre_hoja,
            'prefijo': self.obtener_prefijo(),
            'descripcion': self.obtener_descripcion(),
        }
