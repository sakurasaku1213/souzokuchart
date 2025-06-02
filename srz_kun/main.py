import sys
from PyQt5.QtWidgets import QApplication
from srz_kun.ui.main_window import MainWindowUI # Corrected import path

def main():
    app = QApplication(sys.argv)
    window = MainWindowUI() # Use the MainWindowUI class
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
