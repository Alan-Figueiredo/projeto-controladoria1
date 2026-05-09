from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QLabel,
    QLineEdit,
    QPushButton,
    QVBoxLayout,
    QMessageBox,
)
from PySide6.QtCore import Qt

class Home_page(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Módulos")
        self.setFixedSize(700, 700)

        layout = QVBoxLayout()
        layout.setSpacing(15)

        #titulo
        title = QLabel("Selecione o módulo que deseja utilizar abaixo: ")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            font-size: 20px;
            font-weight: bold;
            font-family: Arial;
            color: #1a4973;
        """)

        self.setStyleSheet("""
                QWidget {
                    background-color: #f6f4f4;
                    color: black;
                    font-size: 14px;
                }

                QPushButton {
                    padding: 10px;
                    border: none;
                    border-radius: 10px;
                    background-color: #1a4973;
                    color: white;
                    font-size: 20px;
                    font-weight: bold;
                    font-family: Arial;
                }

                QPushButton:hover {
                    background-color: #143756;
                }
            """)
        
        #Botoes
        conference_balancete = QPushButton("Conferencia de balancete")
        conference_balancete.setCursor(Qt.PointingHandCursor)
        
        count_func = QPushButton("Contagem de funcionario")
        count_func.setCursor(Qt.PointingHandCursor)

        compensation_managers = QPushButton("Remuneração gerencial")
        compensation_managers.setCursor(Qt.PointingHandCursor)




        layout.addStretch()
        layout.addWidget(title)
        layout.addStretch()

        layout.addWidget(conference_balancete,alignment=Qt.AlignCenter)
        conference_balancete.setFixedWidth(400)
        conference_balancete.setFixedHeight(70)
       

        layout.addWidget(count_func,alignment=Qt.AlignCenter)
        count_func.setFixedWidth(400)
        count_func.setFixedHeight(70)
       


        layout.addWidget(compensation_managers,alignment=Qt.AlignCenter)
        compensation_managers.setFixedWidth(400)
        compensation_managers.setFixedHeight(70)
        layout.addStretch()

        self.setLayout(layout)