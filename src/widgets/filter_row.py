from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QComboBox, QFrame, QHBoxLayout, QLineEdit, QPushButton

from styles.constants import ACCENT, BORDER, DANGER, TEXT_PRIMARY


class FilterRow(QFrame):
    """
    Uma linha de filtro: [dropdown de coluna] [campo de valor] [botão remover]
    Expõe column e value como propriedades.
    """

    def __init__(self, columns: list[str], on_remove, parent=None):
        """
        columns   : colunas filtráveis
        on_remove : callable() chamado ao clicar em remover
        """
        super().__init__(parent)
        self.setStyleSheet("QFrame{background:transparent;border:none;}")

        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        self._combo = QComboBox()
        self._combo.setFixedHeight(32)
        self._combo.setMinimumWidth(140)
        self._combo.setFont(QFont("Segoe UI", 9))
        self._combo.addItems(columns)
        self._combo.setStyleSheet(
            f"QComboBox{{border:1px solid {BORDER};border-radius:6px;"
            f"padding:0 8px;background:white;color:{TEXT_PRIMARY};}}"
            f"QComboBox:focus{{border-color:{ACCENT};}}"
            f"QComboBox::drop-down{{border:none;width:20px;}}"
        )
        lay.addWidget(self._combo)

        self._input = QLineEdit()
        self._input.setFixedHeight(32)
        self._input.setPlaceholderText("Valor...")
        self._input.setFont(QFont("Segoe UI", 9))
        self._input.setStyleSheet(
            f"QLineEdit{{border:1px solid {BORDER};border-radius:6px;"
            f"padding:0 8px;background:white;color:{TEXT_PRIMARY};}}"
            f"QLineEdit:focus{{border-color:{ACCENT};}}"
        )
        lay.addWidget(self._input, 1)

        btn_remove = QPushButton("✕")
        btn_remove.setFixedSize(32, 32)
        btn_remove.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_remove.setToolTip("Remover filtro")
        btn_remove.setStyleSheet(
            f"QPushButton{{background:#FDE8E8;color:{DANGER};"
            f"border:1px solid {DANGER};border-radius:6px;font-weight:bold;}}"
            f"QPushButton:hover{{background:{DANGER};color:white;}}"
        )
        btn_remove.clicked.connect(on_remove)
        lay.addWidget(btn_remove)

    @property
    def column(self) -> str:
        return self._combo.currentText()

    @property
    def value(self) -> str:
        return self._input.text().strip()

    def as_dict(self) -> dict[str, str]:
        return {"column": self.column, "value": self.value}
