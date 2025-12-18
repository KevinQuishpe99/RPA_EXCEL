import sys
import os
import traceback
from PySide6.QtWidgets import QApplication, QMessageBox
from src.vista_qt.principal_qt import VentanaPrincipalQt
from src.controlador.coordinador import CoordinadorPrincipal


def main():
    try:
        # Suprimir warnings menores de Qt
        os.environ['QT_LOGGING_RULES'] = '*.debug=false;qt.qpa.*=false'
        
        app = QApplication(sys.argv)
        ventana = VentanaPrincipalQt()
        ventana.show()
        
        # Conectar el coordinador después de mostrar la ventana
        coordinador = CoordinadorPrincipal(ventana)
        
        # Mensaje inicial (después de mostrar para evitar congelamiento)
        from PySide6.QtCore import QTimer
        QTimer.singleShot(100, lambda: ventana.add_message("✓ Sistema listo para transformar archivos"))
        QTimer.singleShot(200, lambda: ventana.add_message("✓ Selecciona archivo 413 (DV) o 455 (TC)"))
        
        return app.exec()
    except Exception as e:
        error_msg = f"Error al iniciar la aplicación:\n{str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        try:
            QMessageBox.critical(None, "Error de inicio", error_msg)
        except:
            pass
        return 1


if __name__ == "__main__":
    sys.exit(main())