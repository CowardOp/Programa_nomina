from PyQt5.QtWidgets import (
    QMainWindow,
    QPushButton,
    QFileDialog,
    QVBoxLayout,
    QWidget,
    QLabel,
)
import sys
from PyQt5.QtWidgets import QApplication
from logic.excel_analisis import procesador_excel


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Macro Nomina")
        self.setGeometry(100, 100, 600, 400)

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel("Seleccione el archivo a procesar:")
        layout.addWidget(self.label)

        self.button = QPushButton("Seleccione archivo")
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
            self.label.setText(f"Archivo seleccionado: {file_name}")
            processor = procesador_excel()
            processor.calcular_horas_y_colores(file_name)
            self.label.setText(
                f"Archivo procesado y guardado como: {file_name[:-5]}_actualizado.xlsm"
            )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
