from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget

from styles import BORDER
from styles.constants import ACCENT, BG_PAGE
from widgets.name_card import NameCard


class PageGroup(QWidget):
    def __init__(self, page_num, names, start_idx, on_edit, on_delete, parent=None):
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 14)
        lay.setSpacing(4)
        hdr = QFrame()
        hdr.setObjectName("PH")
        hdr.setFixedHeight(32)
        hdr.setStyleSheet(f"QFrame#PH{{background:{ACCENT};border-radius:6px;}}")
        hl = QHBoxLayout(hdr)
        hl.setContentsMargins(12, 0, 12, 0)
        l1 = QLabel(f"Página {page_num}")
        l1.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        l1.setStyleSheet("color:white;")
        hl.addWidget(l1)
        l2 = QLabel(f"{len(names)} crachá{'s' if len(names)!=1 else ''}")
        l2.setFont(QFont("Segoe UI", 9))
        l2.setStyleSheet("color:rgba(255,255,255,0.8);")
        l2.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        hl.addWidget(l2)
        lay.addWidget(hdr)
        box = QFrame()
        box.setStyleSheet(
            f"background:{BG_PAGE};border-radius:8px;border:1px solid {BORDER};"
        )
        bl = QVBoxLayout(box)
        bl.setContentsMargins(8, 8, 8, 8)
        bl.setSpacing(4)
        for i, name in enumerate(names):
            bl.addWidget(NameCard(name, start_idx + i, on_edit, on_delete))
        lay.addWidget(box)
