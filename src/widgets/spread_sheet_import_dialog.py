import os

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from styles.constants import (
    ACCENT,
    ACCENT_LIGHT,
    BG_MAIN,
    BG_SIDEBAR,
    BORDER,
    DANGER,
    SUCCESS,
    TEXT_MUTED,
    TEXT_PRIMARY,
)


class SpreadsheetImportDialog(QDialog):

    def __init__(self, get_columns_fn, parent=None):
        super().__init__(parent)
        self._get_columns = get_columns_fn
        self._file_path = ""

        self._lbl_file: QLabel = QLabel()
        self._combo_name: QComboBox = QComboBox()
        self._combo_surname: QComboBox = QComboBox()
        self._chk_abbreviate: QCheckBox = QCheckBox()
        self._btn_ok: QPushButton = QPushButton()
        self._columns_widget: QWidget = QWidget()

        self.setWindowTitle("Importar Planilha")
        self.setMinimumWidth(460)
        self.setStyleSheet(f"background:{BG_SIDEBAR};color:{TEXT_PRIMARY};")
        self._build_ui()

    @property
    def file_path(self) -> str:
        return self._file_path

    @property
    def name_column(self) -> str:
        return self._combo_name.currentText()

    @property
    def surname_column(self) -> str | None:
        text = self._combo_surname.currentText()
        return text if text != "— nenhum —" else None

    @property
    def abbreviate(self) -> bool:
        return self._chk_abbreviate.isChecked()

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 24, 24, 20)
        lay.setSpacing(16)

        lay.addWidget(self._build_title())
        lay.addWidget(self._build_separator())
        lay.addWidget(self._build_upload_section())
        lay.addWidget(self._build_columns_section())
        lay.addWidget(self._build_options_section())
        lay.addWidget(self._build_separator())
        lay.addWidget(self._build_buttons())

        self._set_columns_enabled(False)

    def _build_title(self) -> QWidget:
        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(2)

        title = QLabel("Importar nomes de planilha")
        title.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        title.setStyleSheet(f"color:{TEXT_PRIMARY};")
        vl.addWidget(title)

        sub = QLabel("Selecione uma planilha .xlsx ou .csv e mapeie as colunas.")
        sub.setFont(QFont("Segoe UI", 9))
        sub.setStyleSheet(f"color:{TEXT_MUTED};")
        vl.addWidget(sub)

        return w

    def _build_separator(self) -> QFrame:
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"color:{BORDER};")
        return sep

    def _build_upload_section(self) -> QWidget:
        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(8)

        lbl = QLabel("Arquivo da planilha")
        lbl.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        lbl.setStyleSheet(f"color:{TEXT_PRIMARY};")
        vl.addWidget(lbl)

        row = QWidget()
        rl = QHBoxLayout(row)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(8)

        self._lbl_file = QLabel("Nenhum arquivo selecionado")
        self._lbl_file.setFont(QFont("Segoe UI", 9))
        self._lbl_file.setStyleSheet(
            f"color:{TEXT_MUTED};background:{BG_MAIN};"
            f"border:1px solid {BORDER};border-radius:6px;padding:6px 10px;"
        )
        self._lbl_file.setMinimumHeight(34)
        rl.addWidget(self._lbl_file, 1)

        btn_select = QPushButton("Selecionar")
        btn_select.setFixedSize(110, 34)
        btn_select.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_select.setFont(QFont("Segoe UI", 9, QFont.Weight.Bold))
        btn_select.setStyleSheet(
            f"QPushButton{{background:{ACCENT_LIGHT};color:{ACCENT};"
            f"border:1px solid {ACCENT};border-radius:6px;}}"
            f"QPushButton:hover{{background:{ACCENT};color:white;}}"
        )
        btn_select.clicked.connect(self._select_file)
        rl.addWidget(btn_select)

        vl.addWidget(row)
        return w

    def _build_columns_section(self) -> QWidget:
        self._columns_widget = QWidget()
        vl = QVBoxLayout(self._columns_widget)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(12)

        section_lbl = QLabel("Mapeamento de colunas")
        section_lbl.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        section_lbl.setStyleSheet(f"color:{TEXT_PRIMARY};")
        vl.addWidget(section_lbl)

        name_row, self._combo_name = self._build_combo_row(
            label="Coluna de nome  *",
            tooltip="Obrigatório",
        )
        surname_row, self._combo_surname = self._build_combo_row(
            label="Coluna de sobrenome",
            tooltip="Opcional — deixe em '— nenhum —' para ignorar",
        )

        vl.addWidget(name_row)
        vl.addWidget(surname_row)

        return self._columns_widget

    def _build_combo_row(self, label: str, tooltip: str) -> tuple[QWidget, QComboBox]:
        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(4)

        header_row = QWidget()
        hl = QHBoxLayout(header_row)
        hl.setContentsMargins(0, 0, 0, 0)

        lbl = QLabel(label)
        lbl.setFont(QFont("Segoe UI", 9))
        lbl.setStyleSheet(f"color:{TEXT_PRIMARY};")
        hl.addWidget(lbl, 1)

        hint = QLabel(tooltip)
        hint.setFont(QFont("Segoe UI", 8))
        hint.setStyleSheet(f"color:{TEXT_MUTED};")
        hl.addWidget(hint)

        vl.addWidget(header_row)

        combo = QComboBox()
        combo.setFixedHeight(34)
        combo.setFont(QFont("Segoe UI", 9))
        combo.setStyleSheet(
            f"QComboBox{{border:2px solid {BORDER};border-radius:6px;"
            f"padding:0 10px;background:white;color:{TEXT_PRIMARY};}}"
            f"QComboBox:focus{{border-color:{ACCENT};}}"
            f"QComboBox::drop-down{{border:none;width:24px;}}"
        )
        vl.addWidget(combo)

        return w, combo

    def _build_options_section(self) -> QWidget:
        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(8)

        lbl = QLabel("Opções de processamento")
        lbl.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        lbl.setStyleSheet(f"color:{TEXT_PRIMARY};")
        vl.addWidget(lbl)

        self._chk_abbreviate = QCheckBox("Abreviar nomes importados da planilha")
        self._chk_abbreviate.setFont(QFont("Segoe UI", 10))
        self._chk_abbreviate.setStyleSheet(f"color:{TEXT_PRIMARY};")
        self._chk_abbreviate.setToolTip(
            'Ex: "Carlos da Silva de Oliveira" → "Carlos Oliveira"'
        )
        vl.addWidget(self._chk_abbreviate)

        hint = QLabel("Mantém apenas o primeiro nome e o último sobrenome.")
        hint.setFont(QFont("Segoe UI", 8))
        hint.setStyleSheet(f"color:{TEXT_MUTED};margin-left:22px;")
        vl.addWidget(hint)

        return w

    def _build_buttons(self) -> QDialogButtonBox:
        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )

        self._btn_ok = btns.button(QDialogButtonBox.StandardButton.Ok)
        self._btn_ok.setText("Importar nomes")
        self._btn_ok.setFixedHeight(36)
        self._btn_ok.setEnabled(False)
        self._btn_ok.setStyleSheet(
            f"QPushButton{{background:{SUCCESS};color:white;"
            f"border-radius:7px;border:none;font-weight:bold;padding:0 16px;}}"
            f"QPushButton:hover{{background:#1B5E20;}}"
            f"QPushButton:disabled{{background:#B0BEC5;}}"
        )

        btn_cancel = btns.button(QDialogButtonBox.StandardButton.Cancel)
        btn_cancel.setText("Cancelar")
        btn_cancel.setFixedHeight(36)
        btn_cancel.setStyleSheet(
            f"QPushButton{{background:transparent;color:{DANGER};"
            f"border:1px solid {DANGER};border-radius:7px;padding:0 16px;}}"
            f"QPushButton:hover{{background:#FDE8E8;}}"
        )

        btns.accepted.connect(self._on_accept)
        btns.rejected.connect(self.reject)
        return btns

    def _select_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar planilha",
            "",
            "Planilhas (*.xlsx *.csv);;Excel (*.xlsx);;CSV (*.csv)",
        )
        if not path:
            return

        try:
            columns = self._get_columns(path)
        except Exception as e:
            QMessageBox.critical(self, "Erro ao ler planilha", str(e))
            return

        self._file_path = path
        self._lbl_file.setText(os.path.basename(path))
        self._lbl_file.setStyleSheet(
            f"color:{SUCCESS};background:{BG_MAIN};"
            f"border:1px solid {SUCCESS};border-radius:6px;padding:6px 10px;"
        )

        self._populate_combos(columns)
        self._set_columns_enabled(True)
        self._btn_ok.setEnabled(True)

    def _populate_combos(self, columns: list[str]):
        self._combo_name.clear()
        self._combo_name.addItems(columns)

        self._combo_surname.clear()
        self._combo_surname.addItem("— nenhum —")
        self._combo_surname.addItems(columns)

    def _set_columns_enabled(self, enabled: bool):
        self._columns_widget.setEnabled(enabled)

    def _on_accept(self):
        if (
            self.surname_column is not None
            and self._combo_name.currentText() == self._combo_surname.currentText()
        ):
            QMessageBox.warning(
                self,
                "Colunas iguais",
                "A coluna de nome e sobrenome não podem ser a mesma.",
            )
            return
        self.accept()
