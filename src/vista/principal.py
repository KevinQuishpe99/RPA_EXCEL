# src/vista/principal.py
"""
Vista Principal
Contiene la interfaz gráfica principal
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from queue import Queue


class VistaPrincipal:
    """Interfaz gráfica principal de la aplicación"""
    
    def __init__(self, raiz):
        self.raiz = raiz
        self.raiz.title("Transformador Excel RPA")
        self.raiz.geometry("650x700")
        self.raiz.resizable(False, False)

        # Estilo moderno
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except:
            pass
        
        # Colores principales
        COLOR_HEADER = '#FF7F50'  # Naranja
        COLOR_BG = '#FFFFFF'
        COLOR_TEXT = '#2F3B52'
        COLOR_LIGHT = '#F5F5F5'
        COLOR_PRIMARY = '#5B7EFF'  # Azul
        COLOR_SECONDARY = '#D3D3D3'  # Gris
        
        # Estilos
        style.configure('TFrame', background=COLOR_BG)
        style.configure('Header.TFrame', background=COLOR_HEADER)
        style.configure('Header.TLabel', background=COLOR_HEADER, foreground='white', font=('Segoe UI', 18, 'bold'))
        style.configure('Subtitle.TLabel', background=COLOR_HEADER, foreground='white', font=('Segoe UI', 10))
        style.configure('TLabel', background=COLOR_BG, foreground=COLOR_TEXT, font=('Segoe UI', 10))
        style.configure('Title.TLabel', background=COLOR_BG, foreground=COLOR_TEXT, font=('Segoe UI', 11, 'bold'))
        style.configure('TButton', font=('Segoe UI', 10, 'bold'), padding=10)
        style.configure('Primary.TButton', font=('Segoe UI', 11, 'bold'), padding=12, background=COLOR_PRIMARY, foreground='white')
        # Ajuste de altura: igualar al alto del área de "Seleccionar archivo"
        # El label usa padding=12, así que alineamos el botón secundario al mismo padding
        style.configure('Secondary.TButton', font=('Segoe UI', 10), padding=12)
        style.configure('TCombobox', padding=8, font=('Segoe UI', 10))
        style.configure('Horizontal.TProgressbar', thickness=8)
        
        # Cola para mensajes thread-safe
        self.cola_mensajes = Queue()
        
        # Callbacks de controlador
        self.callback_transformar = None
        self.callback_seleccionar_archivo = None
        self.callback_descargar = None
        
        # Archivo resultado
        self.archivo_resultado = None
        
        self._crear_ui()
    
    def _crear_ui(self):
        """Crea los elementos de la interfaz"""
        # Configurar grid para expandir
        self.raiz.columnconfigure(0, weight=1)
        self.raiz.rowconfigure(0, weight=1)
        
        # ===== HEADER NARANJA =====
        frame_header = ttk.Frame(self.raiz, style='Header.TFrame')
        frame_header.grid(row=0, column=0, sticky="ew")
        frame_header.columnconfigure(0, weight=1)
        
        label_titulo = ttk.Label(frame_header, text="Transformador Raúl Coka", style='Header.TLabel')
        label_titulo.pack(pady=(20, 5), padx=20)
        
        label_subtitulo = ttk.Label(frame_header, text="Barriga", style='Header.TLabel')
        label_subtitulo.pack(pady=(0, 5), padx=20)
        
        label_descripcion = ttk.Label(frame_header, text="Convierte tus archivos con seguridad", style='Subtitle.TLabel')
        label_descripcion.pack(pady=(0, 20), padx=20)
        
        # ===== CONTENIDO PRINCIPAL =====
        frame_contenido = ttk.Frame(self.raiz)
        frame_contenido.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        frame_contenido.columnconfigure(0, weight=1)
        self.raiz.rowconfigure(1, weight=1)
        
        # --- ARCHIVO DE ORIGEN ---
        label_archivo_titulo = ttk.Label(frame_contenido, text="ARCHIVO DE ORIGEN", style='Title.TLabel')
        label_archivo_titulo.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        
        frame_archivo = ttk.Frame(frame_contenido)
        frame_archivo.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 20))
        frame_archivo.columnconfigure(0, weight=1)
        
        self.label_archivo = ttk.Label(frame_archivo, text="No seleccionado", foreground="#999999", 
                                       font=('Segoe UI', 10), background='#F5F5F5', padding=12)
        self.label_archivo.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        self.btn_seleccionar = ttk.Button(frame_archivo, text="SELECCIONAR", command=self.seleccionar_archivo, style='Secondary.TButton')
        self.btn_seleccionar.grid(row=0, column=1)
        
        # --- TIPO DE ARCHIVO ---
        label_tipo_titulo = ttk.Label(frame_contenido, text="TIPO DE ARCHIVO", style='Title.TLabel')
        label_tipo_titulo.grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 8))
        
        self.combo_poliza = ttk.Combobox(frame_contenido, width=25, state="disabled", font=('Segoe UI', 10))
        self.combo_poliza.grid(row=3, column=0, columnspan=2, sticky="w", pady=(0, 20))
        self.combo_poliza.bind("<<ComboboxSelected>>", self._on_poliza_selected)
        
        # --- ESTADO ---
        label_estado_titulo = ttk.Label(frame_contenido, text="ESTADO", style='Title.TLabel')
        label_estado_titulo.grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 8))
        
        # Frame para los pasos
        frame_pasos = tk.Frame(frame_contenido, background='#F5F5F5', bd=1, relief='solid')
        frame_pasos.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=(0, 20))
        frame_pasos.columnconfigure(0, weight=1)
        
        self.lista_pasos = tk.Label(frame_pasos, text="1   Selecciona archivo 413\n\n2   Selecciona formato\n\n3   Haz clic en Transformar",
                                   background='#F5F5F5', foreground='#666666', font=('Segoe UI', 10), 
                                   justify='left', padx=20, pady=20)
        self.lista_pasos.grid(row=0, column=0, sticky="nsew")
        
        # --- BARRA DE PROGRESO ---
        barra_container = ttk.Frame(frame_contenido)
        barra_container.grid(row=6, column=0, columnspan=2, sticky='ew', pady=(0, 20))
        barra_container.columnconfigure(0, weight=1)
        
        self.barra_progreso = ttk.Progressbar(barra_container, mode='determinate', maximum=100)
        self.barra_progreso.grid(row=0, column=0, sticky="ew")
        
        self.label_progreso = ttk.Label(barra_container, text="0%", font=("Segoe UI", 9), foreground='#999999')
        self.label_progreso.grid(row=0, column=1, sticky="e", padx=(8, 0))
        
        # --- BOTONES DE ACCIÓN ---
        frame_botones = ttk.Frame(frame_contenido)
        frame_botones.grid(row=7, column=0, columnspan=2, sticky="ew")
        frame_botones.columnconfigure(0, weight=1)
        
        self.btn_transformar = ttk.Button(frame_botones, text="TRANSFORMAR", state="disabled", 
                                         style='Primary.TButton', command=self.transformar)
        self.btn_transformar.pack(side="left", padx=(0, 10), fill="x", expand=True)
        
        self.btn_descargar = ttk.Button(frame_botones, text="DESCARGAR", state="disabled", 
                                       style='Secondary.TButton', command=self.descargar_resultado)
        self.btn_descargar.pack(side="left", padx=(0, 10))
        
        self.btn_analizar_otro = ttk.Button(frame_botones, text="ANALIZAR OTRO", state="disabled", 
                                           style='Secondary.TButton', command=self.analizar_otro_archivo)
        self.btn_analizar_otro.pack(side="left")
    
    def seleccionar_archivo(self):
        """Abre diálogo para seleccionar archivo"""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo 413",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos", "*.*")]
        )
        if archivo:
            self.establecer_archivo_origen(archivo)
            if self.callback_seleccionar_archivo:
                self.callback_seleccionar_archivo(archivo)
    
    def establecer_archivo_origen(self, ruta):
        """Establece el nombre del archivo seleccionado"""
        import os
        self.label_archivo.config(text=os.path.basename(ruta), foreground="black")
        # Deshabilitar botón de seleccionar después de elegir archivo
        self.btn_seleccionar.config(state="disabled")
        # Habilitar combobox de tipo de archivo
        self.combo_poliza.config(state="readonly")
        # Reset states for new session
        self.btn_transformar.config(state="disabled", style="")
        self.btn_descargar.config(state="disabled", style="")
        self.btn_analizar_otro.config(state="disabled", style="")
    
    def establecer_polizas(self, lista_polizas):
        """Establece las pólizas disponibles"""
        self.combo_poliza.config(values=lista_polizas)
        if len(lista_polizas) > 0:
            self.combo_poliza.current(0)
    
    def obtener_poliza_seleccionada(self):
        """Obtiene la póliza seleccionada"""
        return self.combo_poliza.get()
    
    def _on_poliza_selected(self, event):
        """Callback cuando se selecciona un tipo de archivo"""
        if self.combo_poliza.get():
            self.btn_transformar.config(state="normal")
        else:
            self.btn_transformar.config(state="disabled")
    
    def agregar_mensaje(self, mensaje):
        """Agrega un mensaje a la cola thread-safe"""
        self.cola_mensajes.put(mensaje)
    
    def procesar_mensajes(self):
        """Procesa mensajes de la cola (llamado desde GUI thread)"""
        try:
            while True:
                mensaje = self.cola_mensajes.get_nowait()
                # Simplemente procesa sin mostrar en interfaz
                # Los mensajes se envían pero la interfaz nueva es más simple
        except:
            pass
        
        # Programar próxima llamada
        self.raiz.after(100, self.procesar_mensajes)
    
    def iniciar_progreso(self):
        """Inicia animación de progreso"""
        self.barra_progreso.start()
    
    def detener_progreso(self):
        """Detiene animación de progreso"""
        self.barra_progreso.stop()
    
    def establecer_progreso(self, porcentaje):
        """Establece el progreso a un porcentaje específico (0-100)"""
        porcentaje = max(0, min(100, porcentaje))
        self.barra_progreso['value'] = porcentaje
        self.label_progreso.config(text=f"{int(porcentaje)}%")
        self.raiz.update_idletasks()
    
    def reiniciar_progreso(self):
        """Reinicia la barra de progreso a 0%"""
        self.barra_progreso['value'] = 0
        self.label_progreso.config(text="0%")
    
    def habilitar_controles(self, habilitado=True):
        """Habilita/deshabilita los controles principales"""
        estado = "normal" if habilitado else "disabled"
        self.btn_seleccionar.config(state=estado)
        self.btn_transformar.config(state=estado if self.combo_poliza.get() else "disabled")
        self.combo_poliza.config(state="readonly" if habilitado else "disabled")

    def resaltar_descargar(self):
        """Resalta el botón de descargar y deshabilita transformar"""
        self.btn_transformar.config(state="disabled", style="")
        self.btn_descargar.config(state="normal", style="Primary.TButton")
        self.btn_analizar_otro.config(state="disabled", style="")

    def resaltar_analizar_otro(self):
        """Resalta Analizar Otro y deshabilita descargar"""
        self.btn_descargar.config(state="disabled", style="")
        self.btn_analizar_otro.config(state="normal", style="Primary.TButton")
    
    def mostrar_exito(self, titulo, mensaje):
        """Muestra diálogo de éxito"""
        messagebox.showinfo(titulo, mensaje)
    
    def mostrar_error(self, titulo, mensaje):
        """Muestra diálogo de error"""
        messagebox.showerror(titulo, mensaje)
    
    def limpiar(self):
        """Limpia la interfaz"""
        self.label_archivo.config(text="No seleccionado", foreground="#999999")
        self.archivo_resultado = None
        self.btn_descargar.config(state="disabled")
        self.btn_descargar.config(state="disabled", style="")
        self.btn_analizar_otro.config(state="disabled", style="")
        self.reiniciar_progreso()
    
    def transformar(self):
        """Botón para transformar"""
        if self.callback_transformar:
            self.callback_transformar()
    
    def establecer_archivo_resultado(self, ruta, nombre_sugerido=None):
        """Establece el archivo resultado disponible para descarga"""
        self.archivo_resultado = ruta
        self.nombre_archivo_sugerido = nombre_sugerido if nombre_sugerido else "archivo_transformado.xlsx"
        self.resaltar_descargar()
    
    def descargar_resultado(self):
        """Descarga el archivo con diálogo de selección"""
        if not self.archivo_resultado:
            messagebox.showerror("Error", "No hay archivo para descargar")
            return
        
        # Usar nombre sugerido
        nombre_archivo = getattr(self, 'nombre_archivo_sugerido', 'archivo_transformado.xlsx')
        
        # Mostrar diálogo para seleccionar ubicación
        ruta_descarga = filedialog.asksaveasfilename(
            title="Guardar archivo transformado",
            defaultextension=".xlsx",
            initialfile=nombre_archivo,
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos", "*.*")]
        )
        
        if ruta_descarga and self.callback_descargar:
            self.callback_descargar(self.archivo_resultado, ruta_descarga)
    
    def analizar_otro_archivo(self):
        """Reinicia el flujo para analizar otro archivo"""
        # Limpiar interfaz
        self.limpiar()
        # Rehabilitar botón de seleccionar
        self.btn_seleccionar.config(state="normal")
        # Deshabilitar botón analizar otro
        self.btn_analizar_otro.config(state="disabled", style="")
        # Resetear barra de progreso
        self.reiniciar_progreso()
