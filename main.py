import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QLabel


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoStamp - 批量盖章工具")
        self.setMinimumSize(1000, 700)
        self.setCentralWidget(QLabel("AutoStamp 启动成功"))


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
