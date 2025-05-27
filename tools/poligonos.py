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

class CalculosPoligonosDialog(QDialog):
    def __init__(self, iface):
        super().__init__(iface.mainWindow())
        self.iface = iface
        self.setWindowTitle("Cálculos de Polígonos")
        self.setup_ui()
        self.setMinimumWidth(600)
        self.distance_area = QgsDistanceArea()
        self.distance_area.setEllipsoid('WGS84')
        self.report_data = []
        self.output_layer = None

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
        
        input_layout.addWidget(QLabel("Capa de polígonos:"))
        self.layer_combo = QgsMapLayerComboBox()
        self.layer_combo.setFilters(QgsMapLayerProxyModel.PolygonLayer)
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
        self.internal_cb = QCheckBox("Calcular ángulos internos")
        self.internal_cb.setChecked(True)
        self.external_cb = QCheckBox("Calcular ángulos externos")
        self.azimut_cb = QCheckBox("Calcular azimut")
        self.rumbo_cb = QCheckBox("Calcular rumbo")
        self.distancia_cb = QCheckBox("Calcular distancias entre vértices")
        self.dist_acum_cb = QCheckBox("Calcular distancias acumuladas")
        self.area_cb = QCheckBox("Calcular área")
        self.area_cb.setChecked(True)
        self.perimetro_cb = QCheckBox("Calcular perímetro")
        self.perimetro_cb.setChecked(True)
        self.reporte_cb = QCheckBox("Generar reporte resumen")
        self.reporte_cb.setChecked(True)
        
        config_layout.addWidget(self.internal_cb)
        config_layout.addWidget(self.external_cb)
        config_layout.addWidget(self.azimut_cb)
        config_layout.addWidget(self.rumbo_cb)
        config_layout.addWidget(self.distancia_cb)
        config_layout.addWidget(self.dist_acum_cb)
        config_layout.addWidget(self.area_cb)
        config_layout.addWidget(self.perimetro_cb)
        config_layout.addWidget(self.reporte_cb)
        
        # Formato de ángulo
        config_layout.addWidget(QLabel("Formato de salida:"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["Grados Decimales", "Grados/Minutos/Segundos", "Radianes"])
        config_layout.addWidget(self.format_combo)
        
        # Unidades de área
        config_layout.addWidget(QLabel("Unidades de área:"))
        self.area_unit_combo = QComboBox()
        self.area_unit_combo.addItems(["Metros cuadrados", "Hectáreas", "Kilómetros cuadrados"])
        config_layout.addWidget(self.area_unit_combo)
        
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
        
        # Nuevas opciones de exportación
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
        self.log_label = QLabel("Preparado para realizar cálculos sobre polígonos...")
        self.log_label.setWordWrap(True)
        self.log_layout.addWidget(self.log_label)
        
        # Añadir área para el reporte resumen
        self.reporte_text = QTextEdit()
        self.reporte_text.setReadOnly(True)
        self.log_layout.addWidget(self.reporte_text)
        
        self.log_tab.setLayout(self.log_layout)
        self.tabs.addTab(self.log_tab, "Registro y Reporte")

        layout.addWidget(self.tabs)

        # Botones
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.setCenterButtons(True)
        self.button_box.accepted.connect(self.calcular_y_guardar)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)
        
        self.setLayout(layout)

    def actualizar_opcion_seleccionados(self):
        """Habilita/deshabilita la opción de usar solo seleccionados"""
        layer = self.layer_combo.currentLayer()
        self.selected_only.setEnabled(layer and layer.selectedFeatureCount() > 0)

    def calcular_distancia(self, p1, p2):
        """Calcula la distancia entre dos puntos"""
        return math.sqrt((p2.x() - p1.x())**2 + (p2.y() - p1.y())**2)

    def calcular_perimetro(self, geometry):
        """Calcula el perímetro de una geometría de polígono"""
        perimetro = 0.0
        if geometry.isMultipart():
            for part in geometry.asMultiPolygon():
                for ring in part:
                    for i in range(len(ring)-1):
                        perimetro += self.calcular_distancia(QgsPointXY(ring[i]), QgsPointXY(ring[i+1]))
        else:
            ring = geometry.asPolygon()[0]
            for i in range(len(ring)-1):
                perimetro += self.calcular_distancia(QgsPointXY(ring[i]), QgsPointXY(ring[i+1]))
        return perimetro

    def calcular_y_guardar(self):
        """Método principal para cálculo y guardado"""
        try:
            # Validación de opciones seleccionadas
            if not (self.internal_cb.isChecked() or self.external_cb.isChecked() or 
                    self.azimut_cb.isChecked() or self.rumbo_cb.isChecked() or
                    self.distancia_cb.isChecked() or self.dist_acum_cb.isChecked() or
                    self.area_cb.isChecked() or self.perimetro_cb.isChecked()):
                QMessageBox.warning(self, "Error", "Seleccione al menos un tipo de cálculo")
                return

            self.log_label.setText("Iniciando cálculo...")
            self.button_box.setEnabled(False)
            QApplication.processEvents()
            
            layer = self.layer_combo.currentLayer()
            if not layer:
                QMessageBox.warning(self, "Error", "Seleccione una capa de polígonos")
                return

            crs = layer.crs()
            self.distance_area.setSourceCrs(crs, QgsCoordinateTransformContext())
            
            # Configurar campos de salida
            fields = QgsFields()
            fields.append(QgsField("id_pol", QVariant.Int))
            fields.append(QgsField("vertice", QVariant.Int))
            fields.append(QgsField("x", QVariant.Double))
            fields.append(QgsField("y", QVariant.Double))
            
            if self.distancia_cb.isChecked():
                fields.append(QgsField("distancia", QVariant.Double))
            
            if self.dist_acum_cb.isChecked():
                fields.append(QgsField("dist_acum", QVariant.Double))
            
            if self.area_cb.isChecked():
                fields.append(QgsField("area", QVariant.Double))
            
            if self.perimetro_cb.isChecked():
                fields.append(QgsField("perimetro", QVariant.Double))
            
            if self.internal_cb.isChecked():
                fields.append(QgsField("ang_int_txt", QVariant.String))
                fields.append(QgsField("ang_int_num", QVariant.Double))
            
            if self.external_cb.isChecked():
                fields.append(QgsField("ang_ext_txt", QVariant.String))
                fields.append(QgsField("ang_ext_num", QVariant.Double))
            
            if self.azimut_cb.isChecked():
                fields.append(QgsField("azimut_txt", QVariant.String))
                fields.append(QgsField("azimut_num", QVariant.Double))
            
            if self.rumbo_cb.isChecked():
                fields.append(QgsField("rumbo_txt", QVariant.String))
                fields.append(QgsField("rumbo_num", QVariant.Double))

            # Procesar cada polígono
            features = layer.selectedFeatures() if self.selected_only.isChecked() and self.selected_only.isEnabled() else layer.getFeatures()
            total = layer.selectedFeatureCount() if self.selected_only.isChecked() else layer.featureCount()
            processed = 0
            
            # Crear capa temporal para procesamiento
            temp_layer = QgsVectorLayer(
                f"Point?crs={crs.authid()}",
                "temp_calculos",
                "memory"
            )
            temp_layer.dataProvider().addAttributes(fields)
            temp_layer.updateFields()
            
            reporte_lines = []
            reporte_lines.append("REPORTE DE CÁLCULOS TOPOGRÁFICOS")
            reporte_lines.append(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            reporte_lines.append(f"Capa de origen: {layer.name()}")
            reporte_lines.append("")
            
            for feature in features:
                geom = feature.geometry()
                if not geom.isEmpty():
                    vertices = self.extraer_vertices(geom)
                    if len(vertices) >= 3:
                        # Calcular área y perímetro
                        area = self.convertir_area(self.distance_area.measureArea(geom)) if self.area_cb.isChecked() else None
                        perimetro = self.calcular_perimetro(geom) if self.perimetro_cb.isChecked() else None
                        
                        # Agregar al reporte
                        reporte_lines.append(f"Polígono ID: {feature.id()}")
                        if self.area_cb.isChecked():
                            reporte_lines.append(f"Área: {area:.4f} {self.area_unit_combo.currentText()}")
                        if self.perimetro_cb.isChecked():
                            reporte_lines.append(f"Perímetro: {perimetro:.4f} metros")
                        reporte_lines.append("Vértices:")
                        
                        distancia_acumulada = 0.0
                        for i in range(len(vertices)):
                            p1 = vertices[i]
                            p0 = vertices[(i - 1) % len(vertices)]
                            p2 = vertices[(i + 1) % len(vertices)]
                            
                            # Calcular distancias
                            distancia = self.calcular_distancia(p1, p2) if self.distancia_cb.isChecked() else None
                            if distancia and self.dist_acum_cb.isChecked():
                                distancia_acumulada += distancia
                            
                            # Calcular ángulos y azimut
                            interno_rad, externo_rad = self.calcular_angulo(p0, p1, p2)
                            azimut_rad = self.calcular_azimut(p1, p2) if self.azimut_cb.isChecked() else 0
                            
                            # Crear feature
                            feat = QgsFeature(fields)
                            feat.setGeometry(QgsGeometry.fromPointXY(p1))
                            
                            # Atributos básicos
                            atributos = [
                                feature.id(),  # id_pol
                                i + 1,        # vertice
                                p1.x(),       # x
                                p1.y()        # y
                            ]
                            
                            # Agregar distancias
                            if self.distancia_cb.isChecked():
                                atributos.append(distancia)
                            if self.dist_acum_cb.isChecked():
                                atributos.append(distancia_acumulada)
                            
                            # Agregar área y perímetro (solo en el primer vértice)
                            if i == 0:
                                if self.area_cb.isChecked():
                                    atributos.append(area)
                                if self.perimetro_cb.isChecked():
                                    atributos.append(perimetro)
                            else:
                                if self.area_cb.isChecked():
                                    atributos.append(None)
                                if self.perimetro_cb.isChecked():
                                    atributos.append(None)
                            
                            # Resto de atributos
                            if self.internal_cb.isChecked():
                                interno_txt, interno_num = self.formatear_angulo(interno_rad)
                                atributos.extend([interno_txt, interno_num])
                                reporte_lines.append(f"  Vértice {i+1}: Ángulo interno: {interno_txt}")
                            
                            if self.external_cb.isChecked():
                                externo_txt, externo_num = self.formatear_angulo(externo_rad)
                                atributos.extend([externo_txt, externo_num])
                                reporte_lines.append(f"  Vértice {i+1}: Ángulo externo: {externo_txt}")
                            
                            if self.azimut_cb.isChecked():
                                azimut_txt, azimut_num = self.formatear_angulo(azimut_rad)
                                atributos.extend([azimut_txt, azimut_num])
                                reporte_lines.append(f"  Lado {i+1}-{i+2 if i+2 <= len(vertices) else 1}: Azimut: {azimut_txt}")
                            
                            if self.rumbo_cb.isChecked():
                                rumbo_txt = self.azimut_a_rumbo(azimut_rad)
                                atributos.extend([rumbo_txt, math.degrees(azimut_rad)])
                                reporte_lines.append(f"  Lado {i+1}-{i+2 if i+2 <= len(vertices) else 1}: Rumbo: {rumbo_txt}")
                            
                            if self.distancia_cb.isChecked():
                                reporte_lines.append(f"  Lado {i+1}-{i+2 if i+2 <= len(vertices) else 1}: Distancia: {distancia:.4f} m")
                            
                            if self.dist_acum_cb.isChecked():
                                reporte_lines.append(f"  Distancia acumulada hasta vértice {i+1}: {distancia_acumulada:.4f} m")
                            
                            feat.setAttributes(atributos)
                            temp_layer.dataProvider().addFeature(feat)
                        
                        reporte_lines.append("")
                
                processed += 1
                self.log_label.setText(f"Procesando... {processed}/{total} polígonos")
                QApplication.processEvents()

            # Actualizar reporte en la interfaz
            if self.reporte_cb.isChecked():
                self.reporte_text.setPlainText("\n".join(reporte_lines))

            # Guardar capa de salida
            if self.temp_rb.isChecked():
                self.output_layer = temp_layer
                QgsProject.instance().addMapLayer(self.output_layer)

            if self.file_rb.isChecked() and self.file_widget.filePath():
                output_path = self.file_widget.filePath()
                options = QgsVectorFileWriter.SaveVectorOptions()
                options.driverName = "ESRI Shapefile" if output_path.endswith('.shp') else "GPKG"
                options.fileEncoding = 'UTF-8'
                
                error = QgsVectorFileWriter.writeAsVectorFormatV2(
                    temp_layer,
                    output_path,
                    QgsCoordinateTransformContext(),
                    options
                )
                
                if error[0] != QgsVectorFileWriter.NoError:
                    QMessageBox.warning(self, "Error", f"No se pudo guardar el archivo: {error[1]}")
                else:
                    self.output_layer = QgsVectorLayer(output_path, os.path.basename(output_path), "ogr")
                    QgsProject.instance().addMapLayer(self.output_layer)

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
                self.exportar_a_csv(self.csv_file_widget.filePath(), temp_layer)
            
            if self.export_excel_rb.isChecked() and self.excel_file_widget.filePath():
                self.exportar_a_excel(self.excel_file_widget.filePath(), temp_layer)
            
            if self.export_pdf_rb.isChecked() and self.pdf_file_widget.filePath():
                self.exportar_a_pdf(self.pdf_file_widget.filePath())

            self.log_label.setText("Proceso completado con éxito")
            self.tabs.setCurrentIndex(1)  # Mostrar pestaña de registro/reporte

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error:\n{str(e)}")
            self.log_label.setText(f"Error: {str(e)}")
        finally:
            self.button_box.setEnabled(True)
            QApplication.processEvents()

    def extraer_vertices(self, geometry):
        """Extrae los vértices de una geometría de polígono"""
        vertices = []
        if geometry.isMultipart():
            for part in geometry.asMultiPolygon():
                for ring in part:
                    vertices.extend([QgsPointXY(point) for point in ring])
        else:
            ring = geometry.asPolygon()[0]
            vertices = [QgsPointXY(point) for point in ring]
        return vertices

    def calcular_angulo(self, p0, p1, p2):
        """Calcula los ángulos interno y externo entre tres puntos"""
        # Vectores
        v1 = QgsPointXY(p0.x() - p1.x(), p0.y() - p1.y())
        v2 = QgsPointXY(p2.x() - p1.x(), p2.y() - p1.y())
        
        # Ángulo entre vectores
        dot = v1.x() * v2.x() + v1.y() * v2.y()
        det = v1.x() * v2.y() - v1.y() * v2.x()
        angulo_rad = math.atan2(det, dot)
        
        # Ajustar ángulo a rango [0, 2π]
        if angulo_rad < 0:
            angulo_rad += 2 * math.pi
            
        interno_rad = angulo_rad
        externo_rad = 2 * math.pi - angulo_rad
        
        return interno_rad, externo_rad

    def calcular_azimut(self, p1, p2):
        """Calcula el azimut entre dos puntos"""
        dx = p2.x() - p1.x()
        dy = p2.y() - p1.y()
        
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

    def formatear_angulo(self, angulo_rad):
        """Formatea el ángulo según las unidades seleccionadas"""
        decimales = self.decimal_spin.value()
        angulo_grados = math.degrees(angulo_rad) % 360
        
        if self.format_combo.currentIndex() == 0:  # Grados decimales
            return f"{angulo_grados:.{decimales}f}°", angulo_grados
        elif self.format_combo.currentIndex() == 1:  # Grados/Minutos/Segundos
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
            
            return f"{grados}° {minutos}' {segundos}\"", angulo_grados
        else:  # Radianes
            return f"{angulo_rad:.{decimales}f} rad", angulo_rad

    def convertir_area(self, area_m2):
        """Convierte el área a las unidades seleccionadas"""
        if self.area_unit_combo.currentIndex() == 0:  # Metros cuadrados
            return area_m2
        elif self.area_unit_combo.currentIndex() == 1:  # Hectáreas
            return area_m2 / 10000
        else:  # Kilómetros cuadrados
            return area_m2 / 1000000

    def exportar_a_csv(self, file_path, layer):
        """Exporta los resultados a CSV"""
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                
                # Escribir encabezados
                headers = [field.name() for field in layer.fields()]
                writer.writerow(headers)
                
                # Escribir datos
                for feature in layer.getFeatures():
                    writer.writerow([feature[field] for field in headers])
                
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
            ws.title = "Cálculos de Polígonos"
            
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
    return CalculosPoligonosDialog(iface)