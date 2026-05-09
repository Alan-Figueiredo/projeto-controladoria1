from PySide6.QtWidgets import QMessageBox
from PySide6.QtCore import Qt


class Message_box:
    def __init__(self, parent=None):
        self.parent = parent

    def message_box_waring(self, title: str, message: str):
        msg = QMessageBox(self.parent)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setStandardButtons(QMessageBox.Ok)

        botao = msg.button(QMessageBox.Ok)
        botao.setCursor(Qt.PointingHandCursor)

        msg.exec()
    
    def message_box_information(self, title: str, message: str):
        msg = QMessageBox(self.parent)
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setStandardButtons(QMessageBox.Ok)

        botao = msg.button(QMessageBox.Ok)
        botao.setCursor(Qt.PointingHandCursor)

        msg.exec()