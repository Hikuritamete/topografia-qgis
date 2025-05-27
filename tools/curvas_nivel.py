# -*- coding: utf-8 -*-

from qgis.PyQt.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QSpinBox, 
    QDoubleSpinBox, QDialogButtonBox, QGroupBox, QWidget, QInputDialog, QFileDialog,
    QCheckBox, QMessageBox, QAction 
)
from qgis.PyQt.QtGui import QColor, QIcon
from qgis.PyQt.QtCore import Qt, QVariant
from qgis.core import (
    QgsRasterLayer, QgsVectorLayer, QgsField, QgsFields, QgsFeature, 
    QgsGeometry, QgsWkbTypes, QgsProject, QgsDistanceArea,
    QgsMapLayerProxyModel, QgsMapLayer, QgsVectorFileWriter,
    QgsCoordinateTransformContext, Qgis, QgsExpression
)
from qgis.gui import QgsMapLayerComboBox
from qgis import processing
import os
import tempfile
import uuid

class CurvasNivelDialog(QDialog):
    """
    Diálogo para la herramienta de generación de curvas de nivel.
    """
    def __init__(self, iface, parent=None):
        super().__init__(parent)
        self.iface = iface
        self.setWindowTitle("Generar Curvas de Nivel")
        self.setMinimumWidth(550)
        try:
            self.setup_ui()
        except Exception as e:
            QMessageBox.critical(self, "Error de Inicialización", f"Ocurrió un error al configurar la interfaz: {e}")
        
    def setup_ui(self):
        """Configura la interfaz de usuario del diálogo."""
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Grupo de Entrada
        input_group = QGroupBox("Capa de Entrada")
        input_layout = QVBoxLayout()
        input_layout.setSpacing(5)

        input_layout.addWidget(QLabel("Seleccionar capa (Raster o Puntos Vectoriales):"))
        self.input_layer_combo = QgsMapLayerComboBox()
        
        # Solución compatible con todas las versiones de QGIS
        # Configuración básica sin usar proxy model
        self.input_layer_combo.setFilters(QgsMapLayerProxyModel.RasterLayer | QgsMapLayerProxyModel.PointLayer)
        self.input_layer_combo.setExcludedProviders(['postgres', 'oracle', 'wms', 'wfs'])
        
        input_layout.addWidget(self.input_layer_combo)

        self.height_field_label = QLabel("Campo de Altura (solo para puntos vectoriales):")
        input_layout.addWidget(self.height_field_label)
        self.height_field_combo = QComboBox()
        input_layout.addWidget(self.height_field_combo)

        # Nuevo: Selector de método de interpolación
        self.interpolation_method_label = QLabel("Método de Interpolación (solo para puntos):")
        input_layout.addWidget(self.interpolation_method_label)
        self.interpolation_method_combo = QComboBox()
        self.interpolation_method_combo.addItems(["TIN Interpolation", "IDW Interpolation"])
        input_layout.addWidget(self.interpolation_method_combo)

        self.input_layer_combo.layerChanged.connect(self.update_input_options)
        input_group.setLayout(input_layout)
        layout.addWidget(input_group)

        # Grupo de Parámetros de Curvas
        params_group = QGroupBox("Parámetros de Generación de Curvas")
        params_layout = QVBoxLayout()
        params_layout.setSpacing(5)

        grid_layout = QHBoxLayout()
        grid_layout.addWidget(QLabel("Intervalo entre curvas:"))
        self.interval_spinbox = QDoubleSpinBox()
        self.interval_spinbox.setMinimum(0.1)
        self.interval_spinbox.setMaximum(100000.0)
        self.interval_spinbox.setSingleStep(0.5)
        self.interval_spinbox.setValue(10.0)
        grid_layout.addWidget(self.interval_spinbox)
        params_layout.addLayout(grid_layout)

        grid_layout = QHBoxLayout()
        grid_layout.addWidget(QLabel("Contorno base:"))
        self.base_contour_spinbox = QDoubleSpinBox()
        self.base_contour_spinbox.setMinimum(-99999.0)
        self.base_contour_spinbox.setMaximum(99999.0)
        self.base_contour_spinbox.setSingleStep(1.0)
        self.base_contour_spinbox.setValue(0.0)
        grid_layout.addWidget(self.base_contour_spinbox)
        params_layout.addLayout(grid_layout)

        grid_layout = QHBoxLayout()
        grid_layout.addWidget(QLabel("Factor Z:"))
        self.z_factor_spinbox = QDoubleSpinBox()
        self.z_factor_spinbox.setMinimum(0.001)
        self.z_factor_spinbox.setMaximum(1000.0)
        self.z_factor_spinbox.setSingleStep(0.1)
        self.z_factor_spinbox.setValue(1.0)
        grid_layout.addWidget(self.z_factor_spinbox)
        params_layout.addLayout(grid_layout)

        grid_layout = QHBoxLayout()
        grid_layout.addWidget(QLabel("Intervalo de curvas maestras (cada N curvas):"))
        self.major_interval_spinbox = QSpinBox()
        self.major_interval_spinbox.setMinimum(1)
        self.major_interval_spinbox.setMaximum(100)
        self.major_interval_spinbox.setValue(5)
        grid_layout.addWidget(self.major_interval_spinbox)
        params_layout.addLayout(grid_layout)

        params_group.setLayout(params_layout)
        layout.addWidget(params_group)

        # La sección de estilo de curvas de nivel ha sido eliminada.
        # Grupo de Estilo de Curvas (predefinido)
        # style_group = QGroupBox("Estilo de Curvas de Nivel (predefinido)")
        # style_layout = QVBoxLayout()
        # style_layout.setSpacing(5)
        # style_layout.addWidget(QLabel("Curvas Secundarias: Línea negra delgada (0.2 mm)"))
        # style_layout.addWidget(QLabel("Curvas Maestras: Línea roja más gruesa (0.6 mm)"))
        # style_group.setLayout(style_layout)
        # layout.addWidget(style_group)

        # Grupo de Opciones de Salida
        output_group = QGroupBox("Opciones de Salida")
        output_layout = QVBoxLayout()

        self.add_to_project_checkbox = QCheckBox("Añadir curvas al proyecto")
        self.add_to_project_checkbox.setChecked(True)
        output_layout.addWidget(self.add_to_project_checkbox)

        self.export_to_file_checkbox = QCheckBox("Exportar a archivo")
        self.export_to_file_checkbox.setChecked(False)
        self.export_to_file_checkbox.stateChanged.connect(self.toggle_export_options)
        output_layout.addWidget(self.export_to_file_checkbox)

        self.export_options_widget = QWidget()
        export_options_layout = QVBoxLayout()
        self.output_format_combo = QComboBox(self)
        self.output_format_combo.addItems(["ESRI Shapefile", "GeoPackage", "KML", "DXF"])
        export_options_layout.addWidget(QLabel("Formato de Salida:"))
        export_options_layout.addWidget(self.output_format_combo)
        self.export_options_widget.setLayout(export_options_layout)
        self.export_options_widget.setVisible(False)
        output_layout.addWidget(self.export_options_widget)

        output_group.setLayout(output_layout)
        layout.addWidget(output_group)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.process_contours)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)
        self.update_input_options() # Llamar al inicio para configurar el estado inicial

    def update_input_options(self):
        """Actualiza las opciones de entrada (campo de altura y método de interpolación)
        basándose en el tipo de capa seleccionada."""
        layer = self.input_layer_combo.currentLayer()
        is_point_layer = (layer and 
                          layer.type() == QgsMapLayer.VectorLayer and 
                          layer.geometryType() == QgsWkbTypes.PointGeometry)

        self.height_field_label.setEnabled(is_point_layer)
        self.height_field_combo.setEnabled(is_point_layer)
        self.interpolation_method_label.setEnabled(is_point_layer)
        self.interpolation_method_combo.setEnabled(is_point_layer)

        if is_point_layer:
            self.populate_height_field_combo(layer)
        else:
            self.height_field_combo.clear()
            # No es necesario limpiar el combo de interpolación, solo deshabilitarlo

    def populate_height_field_combo(self, layer):
        """Rellena el QComboBox del campo de altura con campos numéricos de la capa."""
        self.height_field_combo.clear()
        if layer:
            for field in layer.fields():
                if field.type() in (QVariant.Double, QVariant.Int, QVariant.LongLong):
                    self.height_field_combo.addItem(field.name())

    def toggle_export_options(self, state):
        """Muestra u oculta las opciones de exportación según el estado de la casilla de verificación."""
        self.export_options_widget.setVisible(state == Qt.Checked)

    def process_contours(self):
        """Procesa la generación de curvas de nivel."""
        input_layer = self.input_layer_combo.currentLayer()
        interval = self.interval_spinbox.value()
        base_contour = self.base_contour_spinbox.value()
        z_factor = self.z_factor_spinbox.value()
        major_interval = self.major_interval_spinbox.value()

        if not input_layer:
            QMessageBox.warning(self, "Error", "Debe seleccionar una capa de entrada.")
            return

        # Variables para archivos temporales
        temp_raster_path = None
        temp_contour_vector_path = None

        try:
            # Para capas raster
            if input_layer.type() == QgsMapLayer.RasterLayer:
                # Definir una ruta de archivo temporal para la salida de gdal:contour
                temp_contour_vector_path = os.path.join(tempfile.gettempdir(), f"contours_raster_{uuid.uuid4().hex}.gpkg")

                params_contour = {
                    'INPUT': input_layer.source(), 
                    'BAND': 1, # Asumimos la primera banda para el ráster
                    'INTERVAL': interval,
                    'BASE_CONTOUR': base_contour,
                    'Z_FACTOR': z_factor,
                    'FIELD_NAME': 'ELEV', # Nombre del campo de elevación en las curvas de salida
                    'OUTPUT': temp_contour_vector_path # Salida a archivo temporal
                }
                result = processing.run("gdal:contour", params_contour)
                
                if result and 'OUTPUT' in result:
                    # Cargar la capa desde el archivo temporal
                    contours_layer = QgsVectorLayer(result['OUTPUT'], "Curvas de Nivel", "ogr")
                    if not contours_layer.isValid():
                        QMessageBox.critical(self, "Error", "No se pudo cargar la capa de curvas de nivel generada.")
                        return
                    # self.apply_style(contours_layer, interval, base_contour, major_interval) # Llamada a apply_style eliminada
                    QgsProject.instance().addMapLayer(contours_layer)
                    self.iface.messageBar().pushMessage("Éxito", "Curvas de nivel generadas y añadidas al proyecto.", level=Qgis.Success)
                    if self.export_to_file_checkbox.isChecked():
                        self.export_layer_to_file(contours_layer)
                    self.accept()
                else:
                    QMessageBox.warning(self, "Error", "No se pudieron generar las curvas de nivel desde el ráster.")

            # Para capas vectoriales de puntos
            elif input_layer.type() == QgsMapLayer.VectorLayer and input_layer.geometryType() == QgsWkbTypes.PointGeometry:
                height_field = self.height_field_combo.currentText()
                if not height_field:
                    QMessageBox.warning(self, "Error", "Debe seleccionar un campo de altura para la interpolación de puntos.")
                    return

                if input_layer.featureCount() < 3:
                    QMessageBox.warning(self, "Error", "Se necesitan al menos 3 puntos para la interpolación.")
                    return

                # Asegurarse de que la capa sigue siendo válida y está en el proyecto
                resolved_point_layer = QgsProject.instance().mapLayer(input_layer.id())
                if not resolved_point_layer:
                    QMessageBox.critical(self, "Error de Capa", "No se pudo encontrar la capa de puntos seleccionada en el proyecto. Por favor, asegúrate de que esté cargada.")
                    return
                
                # Crear nombres de archivos temporales únicos para el ráster interpolado y el vector de contornos
                temp_dir = tempfile.gettempdir()
                temp_raster_path = os.path.join(temp_dir, f"interpolated_raster_{uuid.uuid4().hex}.tif")
                temp_contour_vector_path = os.path.join(temp_dir, f"contours_points_{uuid.uuid4().hex}.gpkg") # Nuevo temporal para la salida de contornos
                
                selected_interpolation_method = self.interpolation_method_combo.currentText()
                field_index = resolved_point_layer.fields().indexOf(height_field)

                if selected_interpolation_method == "TIN Interpolation":
                    # Formato para TIN: 'layer_id::~::field_index::~::use_z_bool::~::source_type'
                    # use_z_bool: 0 (False) si se usa un campo de atributo, 1 (True) si se usan coordenadas Z reales
                    # source_type: 0 para fuente de puntos
                    interpolation_data_string = f"{resolved_point_layer.id()}::~::{field_index}::~::{0}::~::{0}"
                    
                    params_interpolation = {
                        'INTERPOLATION_DATA': interpolation_data_string,
                        'METHOD': 0, # Método de interpolación TIN (Linear)
                        'EXTENT': resolved_point_layer.extent(),
                        'PIXEL_SIZE': 10, # Tamaño de píxel predeterminado
                        'OUTPUT': temp_raster_path
                    }
                    result_interpolation = processing.run("qgis:tininterpolation", params_interpolation)

                elif selected_interpolation_method == "IDW Interpolation":
                    # Formato para IDW: 'layer_id::~::field_index::~::use_z_bool::~::source_type'
                    interpolation_data_string = f"{resolved_point_layer.id()}::~::{field_index}::~::{0}::~::{0}"

                    params_interpolation = {
                        'INTERPOLATION_DATA': interpolation_data_string,
                        'POWER': 2.0, # Potencia para IDW
                        'RADIUS': 0.0, # Radio de búsqueda (0.0 para global)
                        'EXTENT': resolved_point_layer.extent(),
                        'PIXEL_SIZE': 10, # Tamaño de píxel predeterminado
                        'OUTPUT': temp_raster_path
                    }
                    result_interpolation = processing.run("qgis:idwinterpolation", params_interpolation)
                else:
                    QMessageBox.warning(self, "Error", "Método de interpolación no válido seleccionado.")
                    return

                if not result_interpolation or not result_interpolation['OUTPUT']:
                    QMessageBox.critical(self, "Error", f"No se pudo generar el ráster interpolado a partir de los puntos con {selected_interpolation_method}.")
                    return

                interpolated_raster = QgsRasterLayer(result_interpolation['OUTPUT'], "Raster Interpolado Temporal")
                if not interpolated_raster.isValid():
                    QMessageBox.critical(self, "Error", "No se pudo cargar el ráster interpolado temporal.")
                    return

                # Paso 2: Generar líneas de contorno a partir del ráster interpolado utilizando gdal:contour
                params_contour = {
                    'INPUT': interpolated_raster.source(),
                    'BAND': 1,
                    'INTERVAL': interval,
                    'BASE_CONTOUR': base_contour,
                    'Z_FACTOR': z_factor,
                    'FIELD_NAME': 'ELEV',
                    'OUTPUT': temp_contour_vector_path # Salida a archivo temporal
                }
                result_contour = processing.run("gdal:contour", params_contour)

                if result_contour and result_contour['OUTPUT']:
                    # Cargar la capa de contornos desde el archivo temporal
                    contours_layer = QgsVectorLayer(result_contour['OUTPUT'], "Curvas de Nivel", "ogr")
                    if not contours_layer.isValid():
                        QMessageBox.critical(self, "Error", "No se pudo cargar la capa de curvas de nivel generada.")
                        return

                    # self.apply_style(contours_layer, interval, base_contour, major_interval) # Llamada a apply_style eliminada
                    QgsProject.instance().addMapLayer(contours_layer)
                    self.iface.messageBar().pushMessage("Éxito", "Curvas de nivel generadas y añadidas al proyecto.", level=Qgis.Success)
                    if self.export_to_file_checkbox.isChecked():
                        self.export_layer_to_file(contours_layer)
                    self.accept()
                else:
                    QMessageBox.critical(self, "Error", "El algoritmo de contorno no generó una salida válida.")

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error al generar las curvas de nivel: {e}")
        finally:
            # Limpieza de archivos temporales
            try:
                if temp_raster_path and os.path.exists(temp_raster_path):
                    os.remove(temp_raster_path)
                if temp_contour_vector_path and os.path.exists(temp_contour_vector_path):
                    os.remove(temp_contour_vector_path)
            except Exception as e:
                self.iface.messageBar().pushMessage("Advertencia", 
                    f"No se pudieron eliminar archivos temporales: {str(e)}", 
                    level=Qgis.Warning, duration=5)

    # El método apply_style ha sido eliminado.

    def export_layer_to_file(self, layer):
        """Exporta la capa de contorno a un archivo."""
        selected_format = self.output_format_combo.currentText()
        file_extension = {
            "ESRI Shapefile": "shp",
            "GeoPackage": "gpkg",
            "KML": "kml",
            "DXF": "dxf"
        }.get(selected_format, "gpkg") # Por defecto a GeoPackage

        file_name, _ = QFileDialog.getSaveFileName(
            self, "Guardar Curvas de Nivel", "",
            f"{selected_format} (*.{file_extension});;Todos los archivos (*.*)"
        )

        if file_name:
            # Asegurarse de que la extensión sea correcta
            if not file_name.lower().endswith(f".{file_extension}"):
                file_name += f".{file_extension}"

            options = QgsVectorFileWriter.SaveVectorOptions()
            options.driverName = selected_format
            options.fileEncoding = "UTF-8"

            # Para DXF, se pueden añadir opciones específicas si es necesario
            if selected_format == "DXF":
                options.datasourceOptions = ['DXF_LAYER_MODE=FEATURE_ATTRIBUTES']

            transform_context = QgsProject.instance().transformContext()

            error = QgsVectorFileWriter.writeAsVectorFormatV2(
                layer,
                file_name,
                transform_context,
                options
            )

            if error[0] == QgsVectorFileWriter.NoError:
                self.iface.messageBar().pushMessage(
                    "Éxito", f"Curvas de nivel exportadas a: {file_name}", level=Qgis.Success
                )
            else:
                QMessageBox.critical(self, "Error de Exportación", f"No se pudo exportar la capa: {error[1]}")


class CurvasNivelTool:
    """
    Clase principal de la herramienta.
    """
    def __init__(self, iface):
        self.iface = iface
        self.dialog = None

    def run(self):
        """Ejecuta la herramienta mostrando el diálogo."""
        if not self.dialog:
            self.dialog = CurvasNivelDialog(self.iface, self.iface.mainWindow())
        self.dialog.exec_()


def classFactory(iface):
    """
    Función llamada cuando QGIS carga el plugin.
    :param iface: Una referencia a la interfaz de QGIS.
    :return: Una instancia de la clase CurvasNivelTool.
    """
    return CurvasNivelTool(iface)
