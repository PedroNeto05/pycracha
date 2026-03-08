from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget

from styles.constants import ACCENT, BG_PAGE, BORDER
from widgets.name_card import NameCard


class PageGroup(QWidget):
    def __init__(
        self,
        page_num: int,
        names: list[str],
        global_start: int,
        on_edit,
        on_delete,
        parent=None,
    ):
        """
        page_num     : número da página (1-based), usado só para exibição
        names        : nomes desta página (já fatiados pelo serviço)
        global_start : índice global do primeiro nome desta página
        on_edit      : callable(global_idx) acionado pelo NameCard ao editar
        on_delete    : callable(global_idx) acionado pelo NameCard ao apagar
        """
        super().__init__(parent)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 14)
        lay.setSpacing(4)

        lay.addWidget(self._build_header(page_num, len(names)))
        lay.addWidget(self._build_cards(names, global_start, on_edit, on_delete))

    def _build_header(self, page_num: int, count: int) -> QFrame:
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

        l2 = QLabel(f"{count} crachá{'s' if count != 1 else ''}")
        l2.setFont(QFont("Segoe UI", 9))
        l2.setStyleSheet("color:rgba(255,255,255,0.8);")
        l2.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        hl.addWidget(l2)

        return hdr

    def _build_cards(self, names, global_start, on_edit, on_delete) -> QFrame:
        box = QFrame()
        box.setStyleSheet(
            f"background:{BG_PAGE};border-radius:8px;border:1px solid {BORDER};"
        )
        bl = QVBoxLayout(box)
        bl.setContentsMargins(8, 8, 8, 8)
        bl.setSpacing(4)

        for i, name in enumerate(names):
            global_idx = global_start + i
            bl.addWidget(
                NameCard(
                    name=name,
                    position=global_idx + 1,  # exibição 1-based
                    on_edit=lambda _, idx=global_idx: on_edit(idx),
                    on_delete=lambda _, idx=global_idx: on_delete(idx),
                )
            )

        return box
