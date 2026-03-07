from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QFrame, QHBoxLayout, QLabel, QPushButton

from styles import ACCENT_DARK, BG_CARD, BORDER
from styles.constants import ACCENT, ACCENT_LIGHT, DANGER, TEXT_MUTED, TEXT_PRIMARY


class NameCard(QFrame):
    def __init__(self, name, index, on_edit, on_delete, parent=None):
        super().__init__(parent)
        self.index = index
        self.setObjectName("NameCard")
        self.setStyleSheet(
            f"QFrame#NameCard{{background:{BG_CARD};border:1px solid {BORDER};border-radius:8px;}}QFrame#NameCard:hover{{border-color:{ACCENT_DARK};}}"
        )
        self.setFixedHeight(48)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 0, 8, 0)
        lay.setSpacing(6)
        num = QLabel(f"{index+1}.")
        num.setFixedWidth(28)
        num.setFont(QFont("Segoe UI", 9))
        num.setStyleSheet(f"color:{TEXT_MUTED};")
        lay.addWidget(num)
        lbl = QLabel(name)
        lbl.setFont(QFont("Segoe UI", 10, QFont.Weight.Medium))
        lbl.setStyleSheet(f"color:{TEXT_PRIMARY};")
        lay.addWidget(lbl, 1)
        for txt, bg, fg, cb in [
            ("Editar", ACCENT_LIGHT, ACCENT, lambda: on_edit(self.index)),
            ("Apagar", "#FDE8E8", DANGER, lambda: on_delete(self.index)),
        ]:
            b = QPushButton(txt)
            b.setFixedSize(80, 30)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setStyleSheet(
                f"QPushButton{{background:{bg};color:{fg};border:1px solid {fg};border-radius:6px;font-size:11px;font-weight:600;}}QPushButton:hover{{background:{fg};color:white;}}"
            )
            b.clicked.connect(cb)
            lay.addWidget(b)
