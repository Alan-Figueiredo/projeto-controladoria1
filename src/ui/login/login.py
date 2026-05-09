import sys
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
from PySide6.QtGui import QPixmap
from src.ui.home_page import home_page as hp
from src.ui.login.Validaded_login import Validaded_login as vl
from src.ui.login.Message_box import Message_box as Mb


class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Login")
        self.setFixedSize(400, 450)

        # Layout principal
        layout = QVBoxLayout()
        layout.setSpacing(15)

        #imagem
        image = QLabel(self)
        pixmap = QPixmap(r"src\ui\login\content\controladoriaImg.jpg")
        pixmap = pixmap.scaled(400,450)
        image.setPixmap(pixmap)

        #titulo
        title = QLabel("Controladoria")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("""
            font-size: 20px;
            font-weight: bold;
            font-family: Arial;
            color: #1a4973;
        """)
        
        # Campo usuário
        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("Usuário")

        # Campo senha
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Senha")
        self.password_input.setEchoMode(QLineEdit.Password)

        # Botão
        login_button = QPushButton("Entrar")
        login_button.setCursor(Qt.PointingHandCursor)
        login_button.clicked.connect(self.login)

        # Estilo
        self.setStyleSheet("""
            QWidget {
                background-color: #f6f4f4;
                color: black;
                font-size: 14px;
            }

            QLineEdit {
                padding: 8px;
                border: 1px solid #444;
                border-radius: 10px;
                background-color: #f6f4f4;
            }

            QPushButton {
                padding: 10px;
                border: none;
                border-radius: 10px;
                background-color: #1a4973;
                color: white;
                font-weight: bold;
            }

            QPushButton:hover {
                background-color: #143756;
            }
        """)

        # Adicionando widgets
        layout.addWidget(image)
        layout.addWidget(title)
        layout.addWidget(self.user_input)
        layout.addWidget(self.password_input)
        layout.addWidget(login_button)

        self.setLayout(layout)

    def login(self):
        self.window = hp.Home_page()
        
        user = self.user_input.text()
        password = self.password_input.text()


        auth_passwd = vl(user,password)
        auth_passwd = auth_passwd.readDB(user,password)

        if user.strip() == "" or password.strip()== "":

            mb = Mb(self)
            mb.message_box_waring("Erro", "Usuário ou senha inválidos")
        elif len(auth_passwd) == 0:
            mb = Mb(self)
            mb.message_box_waring("Erro", "Usuário ou senha inválidos")
        else :
            if password == auth_passwd:
                mb = Mb(self)
                mb.message_box_information("Sucesso", "Login realizado!")
                #QMessageBox.information(self, "Sucesso", "Login realizado!")
                self.window.show()
                self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = LoginWindow()
    window.show()

    sys.exit(app.exec())