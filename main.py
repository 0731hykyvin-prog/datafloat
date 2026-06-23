import sys
from ui.qt_compat import QApplication
from ui.main_window import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1200, 680)
    window.show()
    sys.exit(app.exec_())
