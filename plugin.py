# -*- coding: utf-8 -*-
from qgis.PyQt.QtWidgets import QAction
from qgis.PyQt.QtGui import QIcon
import os
import importlib.util
from qgis.utils import iface

def import_module_from_path(path, module_name):
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

class TopografiaPlugin:
    def __init__(self, iface):
        self.iface = iface
        self.plugin_dir = os.path.dirname(__file__)
        self.actions = []
        self.menu = "&Topografía"
        self.toolbar = None

    def initGui(self):
        # Cargar módulos
        poligonos_path = os.path.join(self.plugin_dir, 'tools', 'poligonos.py')
        self.poligonos_module = import_module_from_path(poligonos_path, 'poligonos')
        
        lineas_path = os.path.join(self.plugin_dir, 'tools', 'lineas.py')
        self.lineas_module = import_module_from_path(lineas_path, 'lineas')
        
        about_path = os.path.join(self.plugin_dir, 'tools', 'about.py')
        self.about_module = import_module_from_path(about_path, 'about')
        
        # Configurar rutas de iconos
        self.icon_paths = {
            'about': os.path.join(self.plugin_dir, 'icons', 'about.png'),
            'icon': os.path.join(self.plugin_dir, 'icons', 'icon.png'),
            'lineas': os.path.join(self.plugin_dir, 'icons', 'lineas.png'),
            'poligonos': os.path.join(self.plugin_dir, 'icons', 'poligonos.png')
        }
        
        # Crear acciones
        self.poligonos_action = QAction(
            QIcon(self.icon_paths['poligonos']),
            "Cálculos de polígonos",
            self.iface.mainWindow()
        )
        self.poligonos_action.triggered.connect(self.run_poligonos)
        self.poligonos_action.setWhatsThis("Calcula ángulos, azimut, rumbo, área y perímetro de polígonos")
        
        self.lineas_action = QAction(
            QIcon(self.icon_paths['lineas']),
            "Cálculos de líneas",
            self.iface.mainWindow()
        )
        self.lineas_action.triggered.connect(self.run_lineas)
        self.lineas_action.setWhatsThis("Calcula azimut, rumbo y distancias de líneas")
        
        self.about_action = QAction(
            QIcon(self.icon_paths['about']),
            "Acerca de",
            self.iface.mainWindow()
        )
        self.about_action.triggered.connect(self.run_about)
        self.about_action.setWhatsThis("Muestra información sobre el plugin")
        
        # Añadir al menú
        self.iface.addPluginToMenu("&Topografía", self.poligonos_action)
        self.iface.addPluginToMenu("&Topografía", self.lineas_action)
        self.iface.addPluginToMenu("&Topografía", self.about_action)
        self.actions.extend([self.poligonos_action, self.lineas_action, self.about_action])
        
        # Añadir a la barra de herramientas
        self.toolbar = self.iface.addToolBar("Topografía")
        self.toolbar.setIconSize(iface.iconSize(True))
        self.toolbar.setObjectName("TopografiaToolbar")
        
        # Añadir acciones a la barra de herramientas (excepto About)
        self.toolbar.addAction(self.poligonos_action)
        self.toolbar.addAction(self.lineas_action)
        
        # Configurar icono del plugin
        if os.path.exists(self.icon_paths['icon']):
            self.iface.pluginToolBar().setIconSize(iface.iconSize(True))
            for action in self.iface.pluginToolBar().actions():
                if action.text() == "Topografía":
                    action.setIcon(QIcon(self.icon_paths['icon']))
                    break

    def unload(self):
        for action in self.actions:
            self.iface.removePluginMenu("&Topografía", action)
            self.iface.removeToolBarIcon(action)
        if hasattr(self, 'toolbar'):
            del self.toolbar

    def run_poligonos(self):
        """Lanza el diálogo de cálculo de polígonos"""
        dialog = self.poligonos_module.CalculosPoligonosDialog(self.iface)
        dialog.exec_()

    def run_lineas(self):
        """Lanza el diálogo de cálculo de líneas"""
        dialog = self.lineas_module.CalculosLineasDialog(self.iface)
        dialog.exec_()
        
    def run_about(self):
        """Lanza el diálogo Acerca de"""
        dialog = self.about_module.AboutDialog(self.iface)
        dialog.exec_()

def classFactory(iface):
    return TopografiaPlugin(iface)