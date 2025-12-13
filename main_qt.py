import sys
from PySide6.QtWidgets import QApplication
from src.vista_qt.principal_qt import VentanaPrincipalQt
from src.controlador.coordinador import CoordinadorPrincipal


def main():
    app = QApplication(sys.argv)
    ventana = VentanaPrincipalQt()
    
    # Conectar el coordinador
    coordinador = CoordinadorPrincipal(ventana)
    
    # Mensaje inicial
    ventana.add_message("✓ Sistema listo para transformar archivos")
    ventana.add_message("✓ Selecciona un archivo 413 para comenzar")
    
    ventana.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()