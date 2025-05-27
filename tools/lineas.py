from qgis.PyQt.QtWidgets import (QDialog, QVBoxLayout, QLabel, QComboBox, 
                                QSpinBox, QDialogButtonBox, QHBoxLayout,
                                QCheckBox, QMessageBox, QGroupBox, 
                                QTabWidget, QWidget, QApplication, QTextEdit)
from qgis.PyQt.QtCore import Qt, QVariant
from qgis.gui import QgsMapLayerComboBox, QgsFileWidget
from qgis.core import (QgsProject, QgsVectorLayer, QgsMapLayerProxyModel,
                      QgsWkbTypes, QgsFields, QgsField, QgsFeature,
                      QgsGeometry, QgsPointXY, QgsCoordinateTransformContext,
                      QgsVectorFileWriter, QgsDistanceArea)
import math
import os
import csv
from datetime import datetime
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from openpyxl import Workbook

class CalculosLineasDialog(QDialog):
    def __init__(self, iface):
        super().__init__(iface.mainWindow())
        self.iface = iface
        self.setWindowTitle("Cálculos de líneas")
        self.setup_ui()
        self.setMinimumWidth(500)
        self.distance_area = QgsDistanceArea()
        self.distance_area.setEllipsoid('WGS84')
        self.report_data = []

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)

        # Crear pestañas
        self.tabs = QTabWidget()
        
        # Pestaña de Parámetros
        self.param_tab = QWidget()
        self.param_layout = QVBoxLayout()
        self.param_layout.setSpacing(10)
        
        # Grupo de entrada
        input_group = QGroupBox("Capa de entrada")
        input_layout = QVBoxLayout()
        
        input_layout.addWidget(QLabel("Capa de líneas:"))
        self.layer_combo = QgsMapLayerComboBox()
        self.layer_combo.setFilters(QgsMapLayerProxyModel.LineLayer)
        input_layout.addWidget(self.layer_combo)
        
        self.selected_only = QCheckBox("Usar solo objetos seleccionados")
        self.selected_only.setEnabled(False)
        input_layout.addWidget(self.selected_only)
        
        self.layer_combo.layerChanged.connect(self.actualizar_opcion_seleccionados)
        input_group.setLayout(input_layout)
        self.param_layout.addWidget(input_group)

        # Grupo de configuración
        config_group = QGroupBox("Configuración de cálculo")
        config_layout = QVBoxLayout()
        
        # Opciones de cálculo
        config_layout.addWidget(QLabel("Opciones de cálculo:"))
        self.azimut_cb = QCheckBox("Calcular azimut")
        self.azimut_cb.setChecked(True)
        self.rumbo_cb = QCheckBox("Calcular rumbo")
        self.rumbo_cb.setChecked(True)
        self.distancia_cb = QCheckBox("Calcular distancias")
        self.distancia_cb.setChecked(True)
        self.dist_acum_cb = QCheckBox("Calcular distancias acumuladas")
        self.dist_acum_cb.setChecked(True)
        self.reporte_cb = QCheckBox("Generar reporte resumen")
        self.reporte_cb.setChecked(True)
        
        config_layout.addWidget(self.azimut_cb)
        config_layout.addWidget(self.rumbo_cb)
        config_layout.addWidget(self.distancia_cb)
        config_layout.addWidget(self.dist_acum_cb)
        config_layout.addWidget(self.reporte_cb)
        
        # Unidades angulares
        config_layout.addWidget(QLabel("Unidades angulares:"))
        self.unit_combo = QComboBox()
        self.unit_combo.addItems(["Grados Decimales", "Grados/Minutos/Segundos"])
        config_layout.addWidget(self.unit_combo)
        
        # Precisión decimal
        config_layout.addWidget(QLabel("Precisión decimal:"))
        self.decimal_spin = QSpinBox()
        self.decimal_spin.setRange(0, 10)
        self.decimal_spin.setValue(4)
        config_layout.addWidget(self.decimal_spin)
        
        config_group.setLayout(config_layout)
        self.param_layout.addWidget(config_group)

        # Grupo de salida
        output_group = QGroupBox("Salida")
        output_layout = QVBoxLayout()
        
        self.temp_rb = QCheckBox("Crear capa temporal")
        self.temp_rb.setChecked(True)
        output_layout.addWidget(self.temp_rb)
        
        self.file_rb = QCheckBox("Guardar en archivo:")
        output_layout.addWidget(self.file_rb)
        
        self.file_widget = QgsFileWidget()
        self.file_widget.setFilter("Shapefile (*.shp);;GeoPackage (*.gpkg);;Todos los archivos (*)")
        self.file_widget.setEnabled(False)
        self.file_widget.setStorageMode(QgsFileWidget.SaveFile)
        output_layout.addWidget(self.file_widget)
        
        self.reporte_file_rb = QCheckBox("Guardar reporte en archivo:")
        output_layout.addWidget(self.reporte_file_rb)
        
        self.reporte_file_widget = QgsFileWidget()
        self.reporte_file_widget.setFilter("Archivo de texto (*.txt);;Todos los archivos (*)")
        self.reporte_file_widget.setEnabled(False)
        self.reporte_file_widget.setStorageMode(QgsFileWidget.SaveFile)
        output_layout.addWidget(self.reporte_file_widget)
        
        # Opciones de exportación
        self.export_csv_rb = QCheckBox("Exportar a CSV:")
        output_layout.addWidget(self.export_csv_rb)
        
        self.csv_file_widget = QgsFileWidget()
        self.csv_file_widget.setFilter("CSV (*.csv);;Todos los archivos (*)")
        self.csv_file_widget.setEnabled(False)
        self.csv_file_widget.setStorageMode(QgsFileWidget.SaveFile)
        output_layout.addWidget(self.csv_file_widget)
        
        self.export_excel_rb = QCheckBox("Exportar a Excel:")
        output_layout.addWidget(self.export_excel_rb)
        
        self.excel_file_widget = QgsFileWidget()
        self.excel_file_widget.setFilter("Excel (*.xlsx);;Todos los archivos (*)")
        self.excel_file_widget.setEnabled(False)
        self.excel_file_widget.setStorageMode(QgsFileWidget.SaveFile)
        output_layout.addWidget(self.excel_file_widget)
        
        self.export_pdf_rb = QCheckBox("Exportar a PDF:")
        output_layout.addWidget(self.export_pdf_rb)
        
        self.pdf_file_widget = QgsFileWidget()
        self.pdf_file_widget.setFilter("PDF (*.pdf);;Todos los archivos (*)")
        self.pdf_file_widget.setEnabled(False)
        self.pdf_file_widget.setStorageMode(QgsFileWidget.SaveFile)
        output_layout.addWidget(self.pdf_file_widget)
        
        # Conectar señales
        self.file_rb.toggled.connect(self.file_widget.setEnabled)
        self.reporte_file_rb.toggled.connect(self.reporte_file_widget.setEnabled)
        self.export_csv_rb.toggled.connect(self.csv_file_widget.setEnabled)
        self.export_excel_rb.toggled.connect(self.excel_file_widget.setEnabled)
        self.export_pdf_rb.toggled.connect(self.pdf_file_widget.setEnabled)
        
        output_group.setLayout(output_layout)
        self.param_layout.addWidget(output_group)

        self.param_tab.setLayout(self.param_layout)
        self.tabs.addTab(self.param_tab, "Parámetros")

        # Pestaña de Registro
        self.log_tab = QWidget()
        self.log_layout = QVBoxLayout()
        self.log_label = QLabel("Preparado para realizar cálculos sobre líneas...")
        self.log_label.setWordWrap(True)
        
        # Añadir área para el reporte resumen
        self.reporte_text = QTextEdit()
        self.reporte_text.setReadOnly(True)
        
        self.log_layout.addWidget(self.log_label)
        self.log_layout.addWidget(self.reporte_text)
        self.log_tab.setLayout(self.log_layout)
        self.tabs.addTab(self.log_tab, "Registro y Reporte")

        layout.addWidget(self.tabs)

        # Botones
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.setCenterButtons(True)
        self.button_box.accepted.connect(self.calcular_azimut_rumbo)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)
        
        self.setLayout(layout)

    def actualizar_opcion_seleccionados(self):
        """Habilita/deshabilita la opción de usar solo seleccionados"""
        layer = self.layer_combo.currentLayer()
        self.selected_only.setEnabled(layer and layer.selectedFeatureCount() > 0)

    def calcular_azimut(self, punto_inicio, punto_fin):
        """Calcula el azimut entre dos puntos"""
        dx = punto_fin.x() - punto_inicio.x()
        dy = punto_fin.y() - punto_inicio.y()
        
        azimut_rad = math.atan2(dx, dy)
        if azimut_rad < 0:
            azimut_rad += 2 * math.pi
            
        return azimut_rad

    def azimut_a_rumbo(self, azimut_rad):
        """Convierte azimut a rumbo"""
        decimales = self.decimal_spin.value()
        azimut_grados = math.degrees(azimut_rad)
        
        if 0 <= azimut_grados < 90:
            return f"N {azimut_grados:.{decimales}f}° E"
        elif 90 <= azimut_grados < 180:
            return f"S {180 - azimut_grados:.{decimales}f}° E"
        elif 180 <= azimut_grados < 270:
            return f"S {azimut_grados - 180:.{decimales}f}° W"
        else:
            return f"N {360 - azimut_grados:.{decimales}f}° W"

    def formatear_angulo(self, angulo_rad, es_rumbo=False):
        """Formatea el ángulo según las unidades seleccionadas"""
        decimales = self.decimal_spin.value()
        angulo_grados = math.degrees(angulo_rad) % 360
        
        if es_rumbo:
            return self.azimut_a_rumbo(angulo_rad)
        
        if self.unit_combo.currentIndex() == 0:  # Grados decimales
            return f"{angulo_grados:.{decimales}f}°"
        else:  # Grados/Minutos/Segundos
            grados = int(angulo_grados)
            minutos_float = abs(angulo_grados - grados) * 60
            minutos = int(minutos_float)
            segundos = round((minutos_float - minutos) * 60, decimales)
            
            # Ajustar redondeo
            if segundos >= 60:
                segundos -= 60
                minutos += 1
            if minutos >= 60:
                minutos -= 60
                grados += 1
            
            return f"{grados}° {minutos}' {segundos}\""

    def calcular_azimut_rumbo(self):
        """Método principal para realizar cálculos sobre líneas"""
        try:
            if not (self.azimut_cb.isChecked() or self.rumbo_cb.isChecked() or 
                    self.distancia_cb.isChecked() or self.dist_acum_cb.isChecked()):
                QMessageBox.warning(self, "Error", "Seleccione al menos un tipo de cálculo")
                return

            self.log_label.setText("Iniciando cálculo...")
            self.button_box.setEnabled(False)
            QApplication.processEvents()
            
            layer = self.layer_combo.currentLayer()
            if not layer:
                QMessageBox.warning(self, "Error", "Seleccione una capa de líneas")
                return

            crs = layer.crs()
            self.distance_area.setSourceCrs(crs, QgsCoordinateTransformContext())
            
            if self.selected_only.isChecked() and self.selected_only.isEnabled():
                features = layer.selectedFeatures()
            else:
                features = layer.getFeatures()

            # Configurar campos de salida
            fields = QgsFields()
            fields.append(QgsField("id_linea", QVariant.Int))
            fields.append(QgsField("segmento", QVariant.Int))
            fields.append(QgsField("x_ini", QVariant.Double))
            fields.append(QgsField("y_ini", QVariant.Double))
            fields.append(QgsField("x_fin", QVariant.Double))
            fields.append(QgsField("y_fin", QVariant.Double))
            
            if self.distancia_cb.isChecked():
                fields.append(QgsField("longitud", QVariant.Double))
            
            if self.dist_acum_cb.isChecked():
                fields.append(QgsField("long_acum", QVariant.Double))
            
            if self.azimut_cb.isChecked():
                fields.append(QgsField("azimut_txt", QVariant.String))
                fields.append(QgsField("azimut_num", QVariant.Double))
            
            if self.rumbo_cb.isChecked():
                fields.append(QgsField("rumbo_txt", QVariant.String))
                fields.append(QgsField("rumbo_num", QVariant.Double))

            # Lista para capas de salida
            output_layers = []
            reporte_lines = []
            reporte_lines.append("REPORTE DE CÁLCULOS DE LÍNEAS")
            reporte_lines.append(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            reporte_lines.append(f"Capa de origen: {layer.name()}")
            reporte_lines.append("")

            # Crear capas según selección
            if self.temp_rb.isChecked():
                temp_layer = QgsVectorLayer(
                    f"Point?crs={crs.authid()}",
                    "calculos_lineas",
                    "memory"
                )
                temp_layer.dataProvider().addAttributes(fields)
                temp_layer.updateFields()
                output_layers.append(temp_layer)

            if self.file_rb.isChecked():
                output_path = self.file_widget.filePath()
                if not output_path:
                    QMessageBox.warning(self, "Error", "Seleccione un archivo de salida")
                    return
                
                file_layer = QgsVectorLayer(
                    f"Point?crs={crs.authid()}",
                    "calculos_lineas",
                    "memory"
                )
                file_layer.dataProvider().addAttributes(fields)
                file_layer.updateFields()
                output_layers.append(file_layer)

            # Procesar cada línea (siempre para todos los segmentos)
            total = layer.selectedFeatureCount() if self.selected_only.isChecked() else layer.featureCount()
            processed = 0
            
            for feature in features:
                geom = feature.geometry()
                if geom.isEmpty():
                    processed += 1
                    continue
                    
                vertices = geom.asPolyline()
                if len(vertices) < 2:
                    processed += 1
                    continue
                
                reporte_lines.append(f"Línea ID: {feature.id()}")
                reporte_lines.append(f"Número de segmentos: {len(vertices)-1}")
                
                # Calcular para cada segmento
                distancia_acumulada = 0.0
                for i in range(len(vertices)-1):
                    punto_inicio = QgsPointXY(vertices[i])
                    punto_fin = QgsPointXY(vertices[i+1])
                    
                    # Calcular longitud del segmento
                    longitud = self.distance_area.measureLine(punto_inicio, punto_fin)
                    distancia_acumulada += longitud
                    
                    # Punto medio para la geometría
                    punto_medio = QgsPointXY(
                        (punto_inicio.x() + punto_fin.x()) / 2,
                        (punto_inicio.y() + punto_fin.y()) / 2
                    )
                    
                    # Calcular azimut
                    azimut_rad = self.calcular_azimut(punto_inicio, punto_fin) if self.azimut_cb.isChecked() else 0
                    
                    # Crear feature para cada capa de salida
                    for output_layer in output_layers:
                        feat = QgsFeature(fields)
                        feat.setGeometry(QgsGeometry.fromPointXY(punto_medio))
                        
                        # Atributos comunes
                        atributos = [
                            feature.id(),
                            i + 1,
                            punto_inicio.x(),
                            punto_inicio.y(),
                            punto_fin.x(),
                            punto_fin.y()
                        ]
                        
                        # Agregar atributos según selección
                        if self.distancia_cb.isChecked():
                            atributos.append(longitud)
                        
                        if self.dist_acum_cb.isChecked():
                            atributos.append(distancia_acumulada)
                        
                        if self.azimut_cb.isChecked():
                            azimut_txt = self.formatear_angulo(azimut_rad)
                            atributos.extend([azimut_txt, math.degrees(azimut_rad)])
                        
                        if self.rumbo_cb.isChecked():
                            rumbo_txt = self.formatear_angulo(azimut_rad, es_rumbo=True)
                            atributos.extend([rumbo_txt, math.degrees(azimut_rad)])
                        
                        feat.setAttributes(atributos)
                        output_layer.dataProvider().addFeature(feat)
                    
                    # Agregar al reporte
                    reporte_lines.append(f"\nSegmento {i+1}:")
                    reporte_lines.append(f"  Punto inicio: ({punto_inicio.x():.4f}, {punto_inicio.y():.4f})")
                    reporte_lines.append(f"  Punto fin: ({punto_fin.x():.4f}, {punto_fin.y():.4f})")
                    if self.distancia_cb.isChecked():
                        reporte_lines.append(f"  Longitud: {longitud:.4f} m")
                    if self.dist_acum_cb.isChecked():
                        reporte_lines.append(f"  Longitud acumulada: {distancia_acumulada:.4f} m")
                    if self.azimut_cb.isChecked():
                        reporte_lines.append(f"  Azimut: {self.formatear_angulo(azimut_rad)}")
                    if self.rumbo_cb.isChecked():
                        reporte_lines.append(f"  Rumbo: {self.formatear_angulo(azimut_rad, es_rumbo=True)}")
                
                reporte_lines.append("\n" + "="*50 + "\n")
                processed += 1
                self.log_label.setText(f"Procesando... {processed}/{total} líneas")
                QApplication.processEvents()

            # Actualizar reporte en la interfaz
            if self.reporte_cb.isChecked():
                self.reporte_text.setPlainText("\n".join(reporte_lines))

            # Guardar y cargar capas
            for layer in output_layers:
                if self.file_rb.isChecked() and layer == file_layer:
                    options = QgsVectorFileWriter.SaveVectorOptions()
                    options.driverName = "ESRI Shapefile" if self.file_widget.filePath().endswith('.shp') else "GPKG"
                    transform_context = QgsCoordinateTransformContext()
                    
                    QgsVectorFileWriter.writeAsVectorFormatV2(
                        layer,
                        self.file_widget.filePath(),
                        transform_context,
                        options
                    )
                    result_layer = QgsVectorLayer(self.file_widget.filePath(), os.path.basename(self.file_widget.filePath()), "ogr")
                    QgsProject.instance().addMapLayer(result_layer)
                elif self.temp_rb.isChecked() and layer == temp_layer:
                    QgsProject.instance().addMapLayer(layer)

            # Guardar reporte en archivo
            if self.reporte_file_rb.isChecked() and self.reporte_file_widget.filePath():
                try:
                    with open(self.reporte_file_widget.filePath(), 'w', encoding='utf-8') as f:
                        f.write("\n".join(reporte_lines))
                    QMessageBox.information(self, "Éxito", f"Reporte guardado en:\n{self.reporte_file_widget.filePath()}")
                except Exception as e:
                    QMessageBox.warning(self, "Error", f"No se pudo guardar el reporte:\n{str(e)}")

            # Exportaciones adicionales
            if self.export_csv_rb.isChecked() and self.csv_file_widget.filePath():
                self.exportar_a_csv(self.csv_file_widget.filePath(), output_layers[0])
            
            if self.export_excel_rb.isChecked() and self.excel_file_widget.filePath():
                self.exportar_a_excel(self.excel_file_widget.filePath(), output_layers[0])
            
            if self.export_pdf_rb.isChecked() and self.pdf_file_widget.filePath():
                self.exportar_a_pdf(self.pdf_file_widget.filePath())

            self.log_label.setText("Proceso completado con éxito")
            self.tabs.setCurrentIndex(1)  # Mostrar pestaña de registro/reporte
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error:\n{str(e)}")
            self.log_label.setText(f"Error: {str(e)}")
        finally:
            self.button_box.setEnabled(True)
            QApplication.processEvents()

    def exportar_a_csv(self, file_path, layer):
        """Exporta los resultados a CSV"""
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                
                # Escribir encabezados
                headers = [field.name() for field in layer.fields()]
                writer.writerow(headers)
                
                # Escribir datos con manejo de valores NULL
                for feature in layer.getFeatures():
                    row = []
                    for field in headers:
                        value = feature[field]
                        # Convertir NULL/None a cadena vacía
                        if value is None or str(value).upper() == 'NULL':
                            row.append('')
                        else:
                            row.append(value)
                    writer.writerow(row)
                
            QMessageBox.information(self, "Éxito", f"Datos exportados a CSV:\n{file_path}")
            return True
        except Exception as e:
            QMessageBox.warning(self, "Error", f"No se pudo exportar a CSV:\n{str(e)}")
            return False

    def exportar_a_excel(self, file_path, layer):
        """Exporta los resultados a Excel"""
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Cálculos de Líneas"
            
            # Escribir encabezados
            headers = [field.name() for field in layer.fields()]
            ws.append(headers)
            
            # Escribir datos con manejo de valores NULL
            for feature in layer.getFeatures():
                row = []
                for field in headers:
                    value = feature[field]
                    # Convertir NULL/None a cadena vacía
                    if value is None or str(value).upper() == 'NULL':
                        row.append('')
                    else:
                        row.append(value)
                ws.append(row)
            
            wb.save(file_path)
            QMessageBox.information(self, "Éxito", f"Datos exportados a Excel:\n{file_path}")
            return True
        except Exception as e:
            QMessageBox.warning(self, "Error", f"No se pudo exportar a Excel:\n{str(e)}")
            return False

    def exportar_a_pdf(self, file_path):
        """Exporta el reporte a PDF"""
        try:
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(file_path)
            
            # Configurar página
            printer.setPageSize(QPrinter.A4)
            printer.setPageMargins(15, 15, 15, 15, QPrinter.Millimeter)
            
            # Imprimir
            self.reporte_text.document().print_(printer)
            
            QMessageBox.information(self, "Éxito", f"Reporte exportado a PDF:\n{file_path}")
            return True
        except Exception as e:
            QMessageBox.warning(self, "Error", f"No se pudo exportar a PDF:\n{str(e)}")
            return False

def classFactory(iface):
    return CalculosLineasDialog(iface)