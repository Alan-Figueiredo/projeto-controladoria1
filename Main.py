from src.core.starting_processing import starting
from src.ui.login import login
from src.ui.home_page import home_page as hp
from PySide6.QtWidgets import QApplication
import sys
   


def main():
    app = QApplication(sys.argv)
    window = login.LoginWindow()
    #window = hp.Home_page()
    window.show()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()