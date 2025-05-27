# -*- coding: utf-8 -*-
from qgis.PyQt.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton
from qgis.PyQt.QtGui import QPixmap
from qgis.PyQt.QtCore import Qt
import os

class AboutDialog(QDialog):
    def __init__(self, iface, parent=None):
        super(AboutDialog, self).__init__(parent)
        self.iface = iface
        self.setWindowTitle("Acerca de Topografía")
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.setMinimumWidth(400)
        
        # Obtener la ruta del directorio del plugin
        # Se asume que about.py está en la carpeta 'tools' dentro del directorio del plugin
        plugin_dir = os.path.dirname(os.path.dirname(__file__))
        icon_path = os.path.join(plugin_dir, 'icons', 'icon.png')
        
        # Crear layout
        layout = QVBoxLayout()
        
        # Añadir logo si existe
        if os.path.exists(icon_path):
            logo_label = QLabel()
            pixmap = QPixmap(icon_path)
            logo_label.setPixmap(pixmap.scaled(100, 100, Qt.KeepAspectRatio))
            logo_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(logo_label)
        
        # Texto descriptivo del plugin
        about_text = """
        <h2>Plugin Topografía</h2>
        <p><b>Versión:</b> 1.1a</p>
        <p><b>Descripción:</b> Herramientas topográficas avanzadas para QGIS,
        incluyendo cálculo de azimuts, rumbos, distancias, áreas y generación de reportes.</p>
        <p><b>Desarrollado por:</b> Omar Ruelas Santa Cruz</p>
        <p><b>Contacto:</b> omarruelassantacruz@gmail.com</p>
        <p><b>Licencia:</b> GPL v3</p>
        <p><b>Página de inicio:</b> <a href="https://github.com/Hikuritamete/topografia-qgis">
        https://github.com/Hikuritamete/topografia-qgis</a></p>
        """
        text_label = QLabel(about_text)
        text_label.setWordWrap(True)
        text_label.setOpenExternalLinks(True) # Permite abrir enlaces en el texto
        layout.addWidget(text_label)
        
        # Botón de cierre
        close_button = QPushButton("Cerrar")
        close_button.clicked.connect(self.close)
        layout.addWidget(close_button, alignment=Qt.AlignCenter)
        
        self.setLayout(layout)
