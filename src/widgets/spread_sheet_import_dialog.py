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
from typing import cast

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
from widgets.filter_row import FilterRow


class SpreadsheetImportDialog(QDialog):
    def __init__(self, get_columns_fn, get_filterable_columns_fn, parent=None):
        """
        get_columns_fn            : callable(file_path) -> list[str]
        get_filterable_columns_fn : callable(file_path) -> list[str]
        """
        super().__init__(parent)
        self._get_columns = get_columns_fn
        self._get_filterable_columns = get_filterable_columns_fn
        self._file_path = ""
        self._filterable_columns: list[str] = []
        self._filter_rows: list[FilterRow] = []

        # Widgets
        self._lbl_file: QLabel = QLabel()
        self._all_columns: list[str] = []
        self._column_rows: list[QWidget] = []
        self._column_combos: list[QComboBox] = []
        self._columns_layout: QVBoxLayout = QVBoxLayout()
        self._chk_abbreviate: QCheckBox = QCheckBox()
        self._btn_ok: QPushButton = QPushButton()
        self._columns_widget: QWidget = QWidget()
        self._filters_widget: QWidget = QWidget()
        self._filters_layout: QVBoxLayout = QVBoxLayout()
        self._lbl_no_filters: QLabel = QLabel()

        self.setWindowTitle("Importar Planilha")
        self.setMinimumWidth(500)
        self.setStyleSheet(f"background:{BG_SIDEBAR};color:{TEXT_PRIMARY};")
        self._build_ui()

    # ── Propriedades expostas ────────────────────────────────────────────────
    @property
    def file_path(self) -> str:
        return self._file_path

    @property
    def name_columns(self) -> list[str]:
        return [
            combo.currentText() for combo in self._column_combos if combo.currentText()
        ]

    @property
    def abbreviate(self) -> bool:
        return self._chk_abbreviate.isChecked()

    @property
    def active_filters(self) -> list[dict[str, str]]:
        return [r.as_dict() for r in self._filter_rows if r.value]

    # ── Construção da UI ─────────────────────────────────────────────────────
    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 24, 24, 20)
        lay.setSpacing(16)

        lay.addWidget(self._build_title())
        lay.addWidget(self._build_separator())
        lay.addWidget(self._build_upload_section())
        lay.addWidget(self._build_columns_section())
        lay.addWidget(self._build_filters_section())
        lay.addWidget(self._build_options_section())
        lay.addWidget(self._build_separator())
        lay.addWidget(self._build_buttons())

        self._set_sections_enabled(False)

    def _build_title(self) -> QWidget:
        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(2)

        title = QLabel("📊 Importar nomes de planilha")
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

        btn_select = QPushButton("📂 Selecionar")
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
        vl.setSpacing(8)

        # Cabeçalho da seção
        header = QWidget()
        hl = QHBoxLayout(header)
        hl.setContentsMargins(0, 0, 0, 0)

        lbl = QLabel("Mapeamento de colunas para o nome")
        lbl.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        lbl.setStyleSheet(f"color:{TEXT_PRIMARY};")
        hl.addWidget(lbl, 1)

        btn_add = QPushButton("＋ Adicionar coluna")
        btn_add.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_add.setFixedHeight(26)
        btn_add.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
        btn_add.setStyleSheet(
            f"QPushButton{{background:{ACCENT_LIGHT};color:{ACCENT};"
            f"border:1px solid {ACCENT};border-radius:5px;padding:0 8px;}}"
            f"QPushButton:hover{{background:{ACCENT};color:white;}}"
        )
        btn_add.clicked.connect(self._add_column_row)
        hl.addWidget(btn_add)

        vl.addWidget(header)

        # Container das linhas dinâmicas de colunas
        cols_container = QWidget()
        self._columns_layout = QVBoxLayout(cols_container)
        self._columns_layout.setContentsMargins(0, 0, 0, 0)
        self._columns_layout.setSpacing(4)
        vl.addWidget(cols_container)

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

    def _build_filters_section(self) -> QWidget:
        self._filters_widget = QWidget()
        vl = QVBoxLayout(self._filters_widget)
        vl.setContentsMargins(0, 0, 0, 0)
        vl.setSpacing(8)

        # Cabeçalho da seção
        header = QWidget()
        hl = QHBoxLayout(header)
        hl.setContentsMargins(0, 0, 0, 0)

        lbl = QLabel("Filtros")
        lbl.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        lbl.setStyleSheet(f"color:{TEXT_PRIMARY};")
        hl.addWidget(lbl, 1)

        btn_add = QPushButton("＋ Adicionar filtro")
        btn_add.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_add.setFixedHeight(26)
        btn_add.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
        btn_add.setStyleSheet(
            f"QPushButton{{background:{ACCENT_LIGHT};color:{ACCENT};"
            f"border:1px solid {ACCENT};border-radius:5px;padding:0 8px;}}"
            f"QPushButton:hover{{background:{ACCENT};color:white;}}"
        )
        btn_add.clicked.connect(self._add_filter)
        hl.addWidget(btn_add)

        vl.addWidget(header)

        # Container das FilterRows
        filters_container = QWidget()
        self._filters_layout = QVBoxLayout(filters_container)
        self._filters_layout.setContentsMargins(0, 0, 0, 0)
        self._filters_layout.setSpacing(4)
        vl.addWidget(filters_container)

        # Hint quando não há filtros
        self._lbl_no_filters = QLabel("Nenhum filtro adicionado.")
        self._lbl_no_filters.setFont(QFont("Segoe UI", 8))
        self._lbl_no_filters.setStyleSheet(f"color:{TEXT_MUTED};")
        vl.addWidget(self._lbl_no_filters)

        return self._filters_widget

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
            'Ex: "Pedro Nascimento de Paiva Fernandes Neto" → "Pedro Neto"'
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

        self._btn_ok = cast(
            QPushButton, btns.button(QDialogButtonBox.StandardButton.Ok)
        )
        self._btn_ok.setText("✅ Importar nomes")
        self._btn_ok.setFixedHeight(36)
        self._btn_ok.setEnabled(False)
        self._btn_ok.setStyleSheet(
            f"QPushButton{{background:{SUCCESS};color:white;"
            f"border-radius:7px;border:none;font-weight:bold;padding:0 16px;}}"
            f"QPushButton:hover{{background:#1B5E20;}}"
            f"QPushButton:disabled{{background:#B0BEC5;}}"
        )

        btn_cancel = cast(
            QPushButton, btns.button(QDialogButtonBox.StandardButton.Cancel)
        )
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

    # ── Interações ───────────────────────────────────────────────────────────
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
            all_columns = self._get_columns(path)
            self._filterable_columns = self._get_filterable_columns(path)
        except Exception as e:
            QMessageBox.critical(self, "Erro ao ler planilha", str(e))
            return

        self._file_path = path
        self._lbl_file.setText(os.path.basename(path))
        self._lbl_file.setStyleSheet(
            f"color:{SUCCESS};background:{BG_MAIN};"
            f"border:1px solid {SUCCESS};border-radius:6px;padding:6px 10px;"
        )

        self._populate_combos(all_columns)
        self._clear_filters()
        self._set_sections_enabled(True)
        self._btn_ok.setEnabled(True)

    def _populate_combos(self, columns: list[str]):
        self._all_columns = columns
        
        # Limpa as seleções anteriores caso o usuário carregue outro arquivo
        for row, combo in zip(list(self._column_rows), list(self._column_combos)):
            self._remove_column_row(row, combo)
            
        # Adiciona a primeira linha obrigatória
        self._add_column_row()

    def _add_filter(self):
        if not self._filterable_columns:
            return
        row = FilterRow(
            columns=self._filterable_columns,
            on_remove=lambda: self._remove_filter(row),
        )
        self._filter_rows.append(row)
        self._filters_layout.addWidget(row)
        self._update_no_filters_hint()

    def _remove_filter(self, row: FilterRow):
        self._filter_rows.remove(row)
        self._filters_layout.removeWidget(row)
        row.deleteLater()
        self._update_no_filters_hint()

    def _clear_filters(self):
        for row in self._filter_rows:
            self._filters_layout.removeWidget(row)
            row.deleteLater()
        self._filter_rows.clear()
        self._update_no_filters_hint()

    def _update_no_filters_hint(self):
        self._lbl_no_filters.setVisible(len(self._filter_rows) == 0)

    def _set_sections_enabled(self, enabled: bool):
        self._columns_widget.setEnabled(enabled)
        self._filters_widget.setEnabled(enabled)

    def _on_accept(self):
        if not self.name_columns:
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos uma coluna para formar o nome.")
            return
        self.accept()

    def _add_column_row(self):
        if not self._all_columns:
            return

        row_widget = QWidget()
        hl = QHBoxLayout(row_widget)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.setSpacing(6)

        combo = QComboBox()
        combo.setFixedHeight(32)
        combo.setFont(QFont("Segoe UI", 9))
        combo.setStyleSheet(
            f"QComboBox{{border:1px solid {BORDER};border-radius:6px;"
            f"padding:0 8px;background:white;color:{TEXT_PRIMARY};}}"
            f"QComboBox:focus{{border-color:{ACCENT};}}"
            f"QComboBox::drop-down{{border:none;width:20px;}}"
        )
        combo.addItems(self._all_columns)
        hl.addWidget(combo, 1)

        btn_remove = QPushButton("✕")
        btn_remove.setFixedSize(32, 32)
        btn_remove.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_remove.setToolTip("Remover coluna")
        btn_remove.setStyleSheet(
            f"QPushButton{{background:#FDE8E8;color:{DANGER};"
            f"border:1px solid {DANGER};border-radius:6px;font-weight:bold;}}"
            f"QPushButton:hover{{background:{DANGER};color:white;}}"
        )
        btn_remove.clicked.connect(lambda: self._remove_column_row(row_widget, combo))

        # Oculta o botão de remover se for a primeira/única coluna, garantindo que haja pelo menos uma.
        if len(self._column_rows) == 0:
            btn_remove.setVisible(False)

        hl.addWidget(btn_remove)

        self._columns_layout.addWidget(row_widget)
        self._column_rows.append(row_widget)
        self._column_combos.append(combo)

    def _remove_column_row(self, row_widget: QWidget, combo: QComboBox):
        self._columns_layout.removeWidget(row_widget)
        self._column_rows.remove(row_widget)
        self._column_combos.remove(combo)
        row_widget.deleteLater()
