from PyQt5.QtWidgets import (
    QMainWindow,
    QPushButton,
    QFileDialog,
    QVBoxLayout,
    QWidget,
    QLabel,
    QComboBox,
    QSpinBox,
    QHBoxLayout,
)
import sys
from PyQt5.QtWidgets import QApplication
from logic.excel_analisis import procesador_excel
from datetime import datetime


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Macro Nómina")
        self.setGeometry(100, 100, 600, 400)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel("Seleccione el archivo a procesar:")
        layout.addWidget(self.label)

        # === Selección de Quincena ===
        quincena_layout = QHBoxLayout()
        quincena_layout.addWidget(QLabel("Quincena:"))
        self.quincena_combo = QComboBox()
        self.quincena_combo.addItems(["Primera", "Segunda"])
        quincena_layout.addWidget(self.quincena_combo)
        layout.addLayout(quincena_layout)

        # === Selección de Mes ===
        mes_layout = QHBoxLayout()
        mes_layout.addWidget(QLabel("Mes:"))
        self.mes_combo = QComboBox()
        self.mes_combo.addItems(
            [
                "Enero",
                "Febrero",
                "Marzo",
                "Abril",
                "Mayo",
                "Junio",
                "Julio",
                "Agosto",
                "Septiembre",
                "Octubre",
                "Noviembre",
                "Diciembre",
            ]
        )
        mes_layout.addWidget(self.mes_combo)
        layout.addLayout(mes_layout)

        # === Selección de anio ===
        anio_layout = QHBoxLayout()
        anio_layout.addWidget(QLabel("Año:"))
        self.anio_spinbox = QSpinBox()
        self.anio_spinbox.setRange(2020, 2100)
        self.anio_spinbox.setValue(datetime.now().year)
        anio_layout.addWidget(self.anio_spinbox)
        layout.addLayout(anio_layout)

        # === Botón Selección de archivo ===
        self.button = QPushButton("Seleccionar archivo")
        self.button.clicked.connect(self.select_file)
        layout.addWidget(self.button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def select_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo de Excel",
            "",
            "Archivos de Excel (*.xlsm);;Todos los archivos (*)",
            options=options,
        )
        if file_name:
            # Aquí solo se obtienen los valores, aún no se usan en el procesador
            quincena = self.quincena_combo.currentText()
            mes = self.mes_combo.currentText()
            anio = self.anio_spinbox.value()

            self.label.setText(
                f"Archivo seleccionado: {file_name}\n"
                f"Periodo: {quincena} quincena de {mes} {anio}"
            )

            # Procesamiento posterior (se actualizará luego con los nuevos parámetros)
            processor = procesador_excel()
            processor.calcular_horas_y_colores(file_name, quincena, mes, anio)

            self.label.setText(
                f"Archivo procesado y guardado como: {file_name[:-5]}_actualizado.xlsm"
            )
