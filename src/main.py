import sys

from PyQt6.QtWidgets import (
    QApplication,
)

from service import DocxService
from widgets import GeradorCrachas


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Gerador de Crachás ECRI")
    docx_service = DocxService()
    w = GeradorCrachas(docx_service)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
