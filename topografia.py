# -*- coding: utf-8 -*-
from qgis.PyQt.QtWidgets import QAction, QMessageBox
from qgis.PyQt.QtGui import QIcon
import os
import importlib.util
from qgis.utils import iface

def import_module_from_path(path, module_name):
    """
    Importa un módulo de Python desde una ruta de archivo específica.
    Esto es útil para cargar módulos que no están en el PYTHONPATH.
    """
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

class TopografiaPlugin:
    """
    Clase principal del plugin Topografía.
    Gestiona la inicialización y descarga del plugin,
    así como el lanzamiento de sus herramientas.
    """
    def __init__(self, iface):
        """
        Constructor del plugin.
        :param iface: Una referencia a la interfaz de QGIS.
        """
        self.iface = iface
        self.plugin_dir = os.path.dirname(__file__)
        self.actions = []  # Lista para almacenar todas las acciones del plugin
        self.menu = "&Topografía"  # Nombre del menú principal del plugin
        self.icon_paths = {} # Diccionario para almacenar las rutas de los iconos

        # Módulos de herramientas que se cargarán dinámicamente
        self.poligonos_module = None
        self.lineas_module = None
        self.curvas_nivel_module = None
        self.about_module = None

        self.load_modules()

    def load_modules(self):
        """Carga dinámicamente los módulos de las herramientas del plugin."""
        # Ruta a la subcarpeta 'tools'
        tools_dir = os.path.join(self.plugin_dir, 'tools')

        module_paths = {
            'poligonos': os.path.join(tools_dir, 'poligonos.py'),
            'lineas': os.path.join(tools_dir, 'lineas.py'),
            'curvas_nivel': os.path.join(tools_dir, 'curvas_nivel.py'),
            'about': os.path.join(tools_dir, 'about.py')
        }

        for module_name, path in module_paths.items():
            try:
                if not os.path.exists(path):
                    raise FileNotFoundError(f"El archivo '{os.path.basename(path)}' no se encontró en la ruta: {path}")

                setattr(self, f"{module_name}_module", import_module_from_path(path, module_name))
            except Exception as e:
                QMessageBox.critical(self.iface.mainWindow(), f"Error al cargar módulo '{module_name}'",
                                     f"No se pudo cargar el módulo '{module_name}'.\nDetalles: {e}\n"
                                     f"Asegúrate de que el archivo '{os.path.basename(path)}' exista en:\n{os.path.dirname(path)}")


    def initGui(self):
        """
        Carga el plugin, añadiendo sus acciones al menú de complementos.
        """
        # Definir rutas de iconos según la estructura de archivos actualizada
        icon_dir = os.path.join(self.plugin_dir, 'icons')
        self.icon_paths['icon'] = os.path.join(icon_dir, 'icon.png')
        self.icon_paths['poligonos'] = os.path.join(icon_dir, 'poligonos.png')
        self.icon_paths['lineas'] = os.path.join(icon_dir, 'lineas.png')
        # Se corrige el nombre del icono para curvas de nivel
        self.icon_paths['curvas_nivel'] = os.path.join(icon_dir, 'curvas.png')
        self.icon_paths['about'] = os.path.join(icon_dir, 'about.png')

        # Crear acciones del menú y añadirlas al menú de complementos de QGIS
        # Acción para Cálculos de Polígonos
        action_poligonos = QAction(
            QIcon(self.icon_paths['poligonos']),
            "Cálculos de Polígonos",
            self.iface.mainWindow()
        )
        action_poligonos.triggered.connect(self.run_poligonos)
        self.iface.addPluginToMenu(self.menu, action_poligonos)
        self.actions.append(action_poligonos)

        # Acción para Cálculos de Líneas
        action_lineas = QAction(
            QIcon(self.icon_paths['lineas']),
            "Cálculos de Líneas",
            self.iface.mainWindow()
        )
        action_lineas.triggered.connect(self.run_lineas)
        self.iface.addPluginToMenu(self.menu, action_lineas)
        self.actions.append(action_lineas)

        # Acción para Generar Curvas de Nivel
        action_curvas_nivel = QAction(
            QIcon(self.icon_paths['curvas_nivel']),
            "Generar Curvas de Nivel",
            self.iface.mainWindow()
        )
        action_curvas_nivel.triggered.connect(self.run_curvas_nivel)
        self.iface.addPluginToMenu(self.menu, action_curvas_nivel)
        self.actions.append(action_curvas_nivel)

        # Acción "Acerca de"
        action_about = QAction(
            QIcon(self.icon_paths['about']),
            "Acerca de Topografía",
            self.iface.mainWindow()
        )
        action_about.triggered.connect(self.run_about)
        self.iface.addPluginToMenu(self.menu, action_about)
        self.actions.append(action_about)

    def unload(self):
        """
        Descarga el plugin, eliminando sus acciones del menú de complementos.
        """
        for action in self.actions:
            self.iface.removePluginMenu(self.menu, action)
            self.iface.removeToolBarIcon(action)

    def run_poligonos(self):
        """Lanza el diálogo de cálculo de polígonos."""
        if self.poligonos_module:
            dialog = self.poligonos_module.CalculosPoligonosDialog(self.iface)
            dialog.exec_()
        else:
            QMessageBox.warning(self.iface.mainWindow(), "Módulo no cargado",
                                 "El módulo de cálculos de polígonos no está disponible.")

    def run_lineas(self):
        """Lanza el diálogo de cálculo de líneas."""
        if self.lineas_module:
            dialog = self.lineas_module.CalculosLineasDialog(self.iface)
            dialog.exec_()
        else:
            QMessageBox.warning(self.iface.mainWindow(), "Módulo no cargado",
                                 "El módulo de cálculos de líneas no está disponible.")

    def run_curvas_nivel(self):
        """Lanza el diálogo de generación de curvas de nivel."""
        if self.curvas_nivel_module:
            dialog = self.curvas_nivel_module.CurvasNivelDialog(self.iface)
            dialog.exec_()
        else:
            QMessageBox.warning(self.iface.mainWindow(), "Módulo no cargado",
                                 "El módulo de curvas de nivel no está disponible.")

    def run_about(self):
        """Lanza el diálogo Acerca de."""
        if self.about_module:
            dialog = self.about_module.AboutDialog(self.iface)
            dialog.exec_()
        else:
            QMessageBox.warning(self.iface.mainWindow(), "Módulo no cargado",
                                 "El módulo 'Acerca de' no está disponible.")
