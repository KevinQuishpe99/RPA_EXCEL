from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QComboBox,
    QProgressBar, QFileDialog, QTextEdit, QFrame
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QLinearGradient, QPalette, QColor, QBrush


class VentanaPrincipalQt(QWidget):
    # Signals to bridge with controller
    solicitar_transformacion = Signal()
    archivo_seleccionado = Signal(str)
    descargar_resultado = Signal(str, str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Transformador Excel RPA - Qt")
        self.resize(760, 640)

        self.archivo_origen = None
        self.ruta_temporal_resultado = None
        self.nombre_archivo_sugerido = "archivo_transformado.xlsx"

        self._build_ui()
        # Aplicar fondo blanco al final para no interferir con gradientes
        self.setStyleSheet("VentanaPrincipalQt{background:#FFFFFF;}")

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(16)

        # Header con gradiente naranja corporativo
        header = QFrame()
        header.setStyleSheet(
            "QFrame{"
            "background: qlineargradient(spread:pad, x1:0, y1:0, x2:0, y2:1, stop:0 #FF6B35, stop:1 #FF8C42);"
            "border-radius:12px;"
            "padding:15px 20px;"
            "}"
        )
        h_layout = QHBoxLayout(header)
        h_layout.setSpacing(15)
        h_layout.setContentsMargins(0, 0, 0, 0)
        
        # Logo
        logo_label = QLabel()
        from PySide6.QtGui import QPixmap
        import os
        logo_path = os.path.join(os.path.dirname(__file__), "..", "img", "logo.png")
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            logo_label.setPixmap(pixmap.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        logo_label.setStyleSheet("background:transparent;")
        h_layout.addWidget(logo_label)
        
        # Textos
        text_layout = QVBoxLayout()
        text_layout.setSpacing(4)
        titulo = QLabel("Transformador Raúl Coka Barriga")
        titulo.setStyleSheet("color:white;font-weight:bold;font-size:20px;background:transparent;")
        subtitulo = QLabel("Convierte tus archivos con seguridad")
        subtitulo.setStyleSheet("color:white;font-size:11px;background:transparent;")
        text_layout.addWidget(titulo)
        text_layout.addWidget(subtitulo)
        h_layout.addLayout(text_layout)
        h_layout.addStretch()
        
        root.addWidget(header)

        # Archivo de origen
        sec_archivo = QVBoxLayout()
        lbl_archivo_titulo = QLabel("ARCHIVO DE ORIGEN")
        lbl_archivo_titulo.setStyleSheet("font-weight:bold;color:#333333;")
        sec_archivo.addWidget(lbl_archivo_titulo)

        archivo_row = QHBoxLayout()
        self.lbl_archivo = QLabel("No seleccionado")
        self.lbl_archivo.setStyleSheet("background:#F5F5F5;padding:12px 14px;border-radius:8px;color:#333333;border:1px solid #CCCCCC")
        # Balance row width: give the path label a bit more stretch
        archivo_row.addWidget(self.lbl_archivo, 1)

        self.btn_seleccionar = QPushButton("SELECCIONAR")
        # Make the select button slightly wider and taller to be comfortable
        self.btn_seleccionar.setStyleSheet(
            "QPushButton{border-radius:8px;padding:8px 14px;background:#FFFFFF;color:#333333;border:1px solid #CCCCCC;min-width: 120px;}"
            "QPushButton:hover{background:#F5F5F5;}"
        )
        try:
            # Harmonize button dimensions
            self.btn_seleccionar.setFixedWidth(130)
            self.btn_seleccionar.setFixedHeight(36)
        except Exception:
            pass
        self.btn_seleccionar.clicked.connect(self._seleccionar_archivo)
        archivo_row.addWidget(self.btn_seleccionar)
        sec_archivo.addLayout(archivo_row)
        root.addLayout(sec_archivo)

        # Tipo de archivo
        sec_tipo = QVBoxLayout()
        lbl_tipo = QLabel("TIPO DE ARCHIVO")
        lbl_tipo.setStyleSheet("font-weight:bold;color:#333333;")
        sec_tipo.addWidget(lbl_tipo)

        self.combo_tipo = QComboBox()
        self.combo_tipo.setEnabled(False)
        # Set a comfortable width/height for the combo box
        try:
            self.combo_tipo.setFixedWidth(260)
            self.combo_tipo.setFixedHeight(34)
        except Exception:
            pass
        self.combo_tipo.setStyleSheet("""
            QComboBox{
                padding:8px 10px;
                border:1px solid #CCCCCC;
                border-radius:8px;
                background:#FFFFFF;
                color:#333333;
                font-size:11px; /* Revert to original size */
                min-width: 220px; /* ensure enough width */
            }
            QComboBox::drop-down{
                width:20px;
                border:none;
                background:#FFFFFF;
            }
            QComboBox::down-arrow{
                image:none;
                width:0px;
            }
            QComboBox QAbstractItemView{
                background:#FFFFFF;
                color:#333333;
                selection-background-color:#FF6347; /* tomato */
                selection-color:#333333; /* keep text dark so it's readable */
                border:1px solid #CCCCCC;
                padding:0px;
                margin:0px;
                font-size:11px;
            }
            QComboBox QAbstractItemView::item{
                padding:8px;
                height:25px;
            }
            QComboBox QAbstractItemView::item:hover{
                background:#FFE4E1; /* light tomato hover */
            }
        """)
        # Agregar mensaje de placeholder
        self.combo_tipo.addItem("Seleccionar tipo de archivo...")
        self.combo_tipo.currentIndexChanged.connect(self._on_tipo_cambiado)
        sec_tipo.addWidget(self.combo_tipo)
        root.addLayout(sec_tipo)

        # INFORMACIÓN DEL PROCESO - más grande
        info_title = QLabel("INFORMACIÓN DEL PROCESO")
        info_title.setStyleSheet("font-weight:bold;color:#333333;font-size:12pt;")
        root.addWidget(info_title)

        self.text_estado = QTextEdit()
        self.text_estado.setReadOnly(True)
        self.text_estado.setMinimumHeight(180)  # Reducido para ver botones
        self.text_estado.setMaximumHeight(220)  # Altura máxima controlada
        self.text_estado.setStyleSheet(
            "QTextEdit{"
            "background: qlineargradient(spread:pad,x1:0,y1:0,x2:1,y2:1, stop:0 #0d0d0d, stop:1 #1a1a1a);"
            "color:#00FF00;"
            "border-radius:10px;"
            "padding:10px;"
            "border:2px solid #FF6B35;"
            "font-family:Consolas,Monaco,monospace;"
            "font-size:10pt;"
            "}"
        )
        root.addWidget(self.text_estado, 1)  # Peso reducido

        # Progreso
        prog_row = QHBoxLayout()
        self.barra = QProgressBar()
        self.barra.setRange(0, 100)
        self.barra.setStyleSheet(
            "QProgressBar{border:1px solid #CCCCCC;border-radius:8px;background:#F5F5F5;text-align:center;color:#333333;}"
            "QProgressBar::chunk{background:qlineargradient(spread:pad,x1:0,y1:0,x2:1,y2:0,stop:0 #FF6B35, stop:1 #FF8C42);border-radius:6px;}"
        )
        prog_row.addWidget(self.barra, 1)
        self.lbl_porcentaje = QLabel("0%")
        self.lbl_porcentaje.setStyleSheet("color:#666666;font-weight:bold;")
        prog_row.addWidget(self.lbl_porcentaje)
        root.addLayout(prog_row)

        # Botones
        btn_row = QHBoxLayout()
        self.btn_transformar = QPushButton("TRANSFORMAR")
        # Gradient button style
        self.btn_transformar.setStyleSheet(
            "QPushButton:disabled{background:#E5E7EB;color:#9CA3AF;border:1px solid #D1D5DB;border-radius:10px;padding:12px 20px;font-weight:bold;}"
            "QPushButton{border-radius:10px;padding:12px 20px;color:white;font-weight:bold;"
            "background:qlineargradient(spread:pad,x1:0,y1:0,x2:1,y2:0,stop:0 #FF6B35, stop:1 #FF8C42);"
            "border:2px solid #FF6B35;}"
            "QPushButton:hover{background:qlineargradient(spread:pad,x1:0,y1:0,x2:1,y2:0,stop:0 #FF8C42, stop:1 #FF6B35);}"
        )
        self.btn_transformar.setEnabled(False)
        self.btn_transformar.clicked.connect(lambda: self.solicitar_transformacion.emit())
        btn_row.addWidget(self.btn_transformar)

        self.btn_descargar = QPushButton("DESCARGAR")
        self.btn_descargar.setEnabled(False)
        self.btn_descargar.setStyleSheet(
            "QPushButton{border-radius:10px;padding:12px 20px;"
            "background:qlineargradient(spread:pad,x1:0,y1:0,x2:1,y2:0,stop:0 #FF6B35, stop:1 #FF8C42);"
            "color:white;font-weight:bold;border:2px solid #FF6B35;}"
            "QPushButton:hover{background:qlineargradient(spread:pad,x1:0,y1:0,x2:1,y2:0,stop:0 #FF8C42, stop:1 #FF6B35);}"
            "QPushButton:disabled{background:#F5F5F5;color:#999999;border:1px solid #CCCCCC;}"
        )
        self.btn_descargar.clicked.connect(self._descargar)
        btn_row.addWidget(self.btn_descargar)

        self.btn_analizar_otro = QPushButton("ANALIZAR OTRO")
        self.btn_analizar_otro.setEnabled(False)
        self.btn_analizar_otro.setStyleSheet(
            "QPushButton{border-radius:10px;padding:12px 20px;"
            "background:qlineargradient(spread:pad,x1:0,y1:0,x2:1,y2:0,stop:0 #FF6B35, stop:1 #FF8C42);"
            "color:white;font-weight:bold;border:2px solid #FF6B35;}"
            "QPushButton:hover{background:qlineargradient(spread:pad,x1:0,y1:0,x2:1,y2:0,stop:0 #FF8C42, stop:1 #FF6B35);}"
            "QPushButton:disabled{background:#F5F5F5;color:#999999;border:1px solid #CCCCCC;}"
        )
        self.btn_analizar_otro.clicked.connect(self._analizar_otro)
        btn_row.addWidget(self.btn_analizar_otro)
        root.addLayout(btn_row)

    # ===== Public API for controller =====
    def set_polizas(self, nombres):
        self.combo_tipo.clear()
        self.combo_tipo.addItem("Seleccionar tipo de archivo...")  # Placeholder
        self.combo_tipo.addItems(nombres)
        self.combo_tipo.setCurrentIndex(0)  # Mostrar placeholder por defecto

    def set_archivo_resultado_temp(self, ruta, nombre_sugerido=None):
        self.ruta_temporal_resultado = ruta
        self.nombre_archivo_sugerido = nombre_sugerido if nombre_sugerido else "archivo_transformado.xlsx"
        self.btn_descargar.setEnabled(True)
        self.btn_transformar.setEnabled(False)

    def add_message(self, msg):
        # Prefix like console prompt
        self.text_estado.append(f"> {msg}")

    def set_progress(self, value):
        v = max(0, min(100, int(value)))
        self.barra.setValue(v)
        self.lbl_porcentaje.setText(f"{v}%")

    def highlight_descargar(self):
        self.btn_descargar.setEnabled(True)
        self.btn_transformar.setEnabled(False)

    def highlight_analizar(self):
        self.btn_analizar_otro.setEnabled(True)
        self.btn_descargar.setEnabled(False)
    
    def obtener_poliza_seleccionada(self):
        """Retorna el nombre de la póliza seleccionada"""
        return self.combo_tipo.currentText()
    
    def mostrar_error(self, titulo, mensaje):
        """Muestra un diálogo de error"""
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.critical(self, titulo, mensaje)
    
    def mostrar_exito(self, titulo, mensaje):
        """Muestra un diálogo de éxito"""
        from PySide6.QtWidgets import QMessageBox
        QMessageBox.information(self, titulo, mensaje)

    # ===== Internal slots =====
    def _seleccionar_archivo(self):
        ruta, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo de origen (413 o 455)", filter="Excel (*.xlsx)")
        if ruta:
            self.archivo_origen = ruta
            self.lbl_archivo.setText(ruta.split('/')[-1])
            self.btn_seleccionar.setEnabled(False)
            self.combo_tipo.setEnabled(True)
            self.archivo_seleccionado.emit(ruta)

    def _on_tipo_cambiado(self, idx):
        # idx=0 es el placeholder "Seleccionar tipo de archivo..."
        if idx > 0:
            self.btn_transformar.setEnabled(True)
        else:
            self.btn_transformar.setEnabled(False)

    def _descargar(self):
        if not self.ruta_temporal_resultado:
            return
        # Usar nombre sugerido como nombre inicial
        destino, _ = QFileDialog.getSaveFileName(
            self, 
            "Guardar archivo transformado",
            self.nombre_archivo_sugerido,  # Nombre inicial sugerido
            "Excel (*.xlsx)"
        )
        if destino:
            self.descargar_resultado.emit(self.ruta_temporal_resultado, destino)

    def _analizar_otro(self):
        # reset simple
        self.archivo_origen = None
        self.ruta_temporal_resultado = None
        self.lbl_archivo.setText("No seleccionado")
        self.btn_seleccionar.setEnabled(True)
        self.combo_tipo.setEnabled(False)
        self.combo_tipo.setCurrentIndex(-1)
        self.btn_transformar.setEnabled(False)
        self.btn_descargar.setEnabled(False)
        self.btn_analizar_otro.setEnabled(False)
        self.set_progress(0)
        # Limpiar consola y porcentaje
        self.text_estado.clear()
        if hasattr(self, 'lbl_porcentaje'):
            self.lbl_porcentaje.setText("0%")

    def habilitar_controles(self, habilitado=True):
        """Habilita/deshabilita los controles principales durante transformación"""
        self.btn_seleccionar.setEnabled(habilitado)
        # Solo habilitar transformar si hay póliza seleccionada
        if habilitado and self.combo_tipo.currentIndex() >= 0:
            self.btn_transformar.setEnabled(True)
        else:
            self.btn_transformar.setEnabled(False)
        self.combo_tipo.setEnabled(habilitado)

    # ===== UI helpers =====
    # Métodos _make_pill y _set_pill_state eliminados (sección ESTADO removida)