from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton

from styles.constants import (
    ACCENT,
    ACCENT_DARK,
    ACCENT_LIGHT,
    DANGER,
    TEXT_MUTED,
    TEXT_PRIMARY,
)


class NameCard(QFrame):
    def __init__(self, name: str, position: int, on_edit, on_delete, parent=None):
        super().__init__(parent)
        self.setObjectName("NameCard")
        self.setStyleSheet(f"QFrame#NameCard:hover{{border-color:{ACCENT_DARK};}}")
        self.setFixedHeight(48)

        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 0, 8, 0)
        lay.setSpacing(6)

        num = QLabel(f"{position}.")
        num.setFixedWidth(28)
        num.setFont(QFont("Segoe UI", 9))
        num.setStyleSheet(f"color:{TEXT_MUTED};")
        lay.addWidget(num)

        lbl = QLabel(name)
        lbl.setFont(QFont("Segoe UI", 10, QFont.Weight.Medium))
        lbl.setStyleSheet(f"color:{TEXT_PRIMARY};")
        lay.addWidget(lbl, 1)

        for txt, bg, fg, cb in [
            ("Editar", ACCENT_LIGHT, ACCENT, on_edit),
            ("Apagar", "#FDE8E8", DANGER, on_delete),
        ]:
            b = QPushButton(txt)
            b.setFixedSize(80, 30)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setStyleSheet(
                f"QPushButton{{background:{bg};color:{fg};border:1px solid {fg};"
                f"border-radius:6px;font-size:11px;font-weight:600;}}"
                f"QPushButton:hover{{background:{fg};color:white;}}"
            )
            b.clicked.connect(cb)
            lay.addWidget(b)
