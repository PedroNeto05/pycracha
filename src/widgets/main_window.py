import os

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QButtonGroup,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from service import DocxService
from styles.constants import (
    ACCENT,
    ACCENT_DARK,
    ACCENT_LIGHT,
    BG_MAIN,
    BG_SIDEBAR,
    BORDER,
    COLOR_DOTS,
    COLOR_OPTIONS,
    DANGER,
    SUCCESS,
    TEXT_MUTED,
    TEXT_PRIMARY,
)
from widgets.page_group import PageGroup


class GeradorCrachas(QMainWindow):
    def __init__(self, docx_service: DocxService):
        super().__init__()
        self.docx_service = docx_service
        self.templates: dict[str, str] = {}
        self.setWindowTitle("Gerador de Crachás ECRI")
        self.setMinimumSize(980, 660)
        self.resize(1120, 740)
        self._build_ui()
        self.setStyleSheet(f"QMainWindow{{background:{BG_MAIN};}}")

    # ── Construção da UI ─────────────────────────────────────────────────────
    def _build_ui(self):
        c = QWidget()
        self.setCentralWidget(c)
        root = QHBoxLayout(c)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root.addWidget(self._left_panel(), 0)
        root.addWidget(self._right_panel(), 1)

    def _sep(self):
        f = QFrame()
        f.setFrameShape(QFrame.Shape.HLine)
        f.setStyleSheet(f"color:{BORDER};")
        return f

    def _grpbox(self, title: str) -> QGroupBox:
        gb = QGroupBox(title)
        gb.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        gb.setStyleSheet(
            f"QGroupBox{{color:{TEXT_PRIMARY};border:1px solid {BORDER};"
            f"border-radius:8px;margin-top:10px;padding-top:8px;}}"
            f"QGroupBox::title{{subcontrol-origin:margin;left:12px;padding:0 4px;}}"
        )
        return gb

    def _btn(self, text: str, height: int, style: str, cb) -> QPushButton:
        b = QPushButton(text)
        b.setFixedHeight(height)
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        b.setStyleSheet(style)
        b.clicked.connect(cb)
        return b

    def _left_panel(self) -> QFrame:
        panel = QFrame()
        panel.setFixedWidth(340)
        panel.setObjectName("LP")
        panel.setStyleSheet(
            f"QFrame#LP{{background:{BG_SIDEBAR};border-right:1px solid {BORDER};}}"
        )
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(20, 24, 20, 24)
        lay.setSpacing(14)

        t = QLabel("Gerador de Crachás")
        t.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        t.setStyleSheet(f"color:{TEXT_PRIMARY};")
        lay.addWidget(t)

        s = QLabel("ECRI – Encontro de Crianças com Cristo")
        s.setFont(QFont("Segoe UI", 9))
        s.setStyleSheet(f"color:{TEXT_MUTED};")
        lay.addWidget(s)

        lay.addWidget(self._sep())
        lay.addWidget(self._build_templates_group())
        lay.addWidget(self._build_add_name_group())

        self.count_lbl = QLabel("0 nomes adicionados")
        self.count_lbl.setFont(QFont("Segoe UI", 9))
        self.count_lbl.setStyleSheet(f"color:{TEXT_MUTED};")
        self.count_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self.count_lbl)

        lay.addStretch()
        lay.addWidget(self._sep())
        lay.addWidget(
            self._btn(
                "Gerar Arquivo .docx",
                46,
                f"QPushButton{{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,"
                f"stop:0 {SUCCESS},stop:1 #1B5E20);color:white;border-radius:10px;"
                f"border:none;font-size:11pt;font-weight:bold;}}"
                f"QPushButton:hover{{background:#2E7D32;}}"
                f"QPushButton:disabled{{background:#B0BEC5;}}",
                self._generate,
            )
        )
        lay.addWidget(
            self._btn(
                "Limpar Tudo",
                34,
                f"QPushButton{{background:transparent;color:{DANGER};"
                f"border:1px solid {DANGER};border-radius:8px;font-size:9pt;}}"
                f"QPushButton:hover{{background:#FDE8E8;}}",
                self._clear_all,
            )
        )
        return panel

    def _build_templates_group(self) -> QGroupBox:
        grp = self._grpbox("Modelos de Crachá")
        gl = QVBoxLayout(grp)
        gl.setSpacing(6)

        self.radio_group = QButtonGroup(self)
        self.template_labels = {}

        for i, color in enumerate(COLOR_OPTIONS):
            gl.addWidget(self._build_color_row(color))
            gl.addWidget(self._build_color_subrow(color))

            if i < len(COLOR_OPTIONS) - 1:
                div = QFrame()
                div.setFrameShape(QFrame.Shape.HLine)
                div.setStyleSheet(f"color:{BORDER};margin:2px 4px;")
                gl.addWidget(div)

        self.radio_group.buttons()[1].setChecked(True)
        return grp

    def _build_color_row(self, color: str) -> QWidget:
        row = QWidget()
        rl = QHBoxLayout(row)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(6)

        rb = QRadioButton(f"  {color}")
        rb.setFont(QFont("Segoe UI", 10))
        rb.setStyleSheet(
            f"QRadioButton{{color:{TEXT_PRIMARY};padding:4px;border-radius:6px;}}"
            f"QRadioButton:hover{{background:{ACCENT_LIGHT};}}"
        )
        self.radio_group.addButton(rb)
        rl.addWidget(rb, 1)

        dot = QLabel("●")
        dot.setFont(QFont("Segoe UI", 16))
        dot.setStyleSheet(f"color:{COLOR_DOTS[color]};padding-right:4px;")
        dot.setFixedWidth(22)
        rl.addWidget(dot)

        return row

    def _build_color_subrow(self, color: str) -> QWidget:
        sub = QWidget()
        sl = QHBoxLayout(sub)
        sl.setContentsMargins(28, 0, 0, 0)
        sl.setSpacing(6)

        lbl_st = QLabel("Nenhum modelo carregado")
        lbl_st.setFont(QFont("Segoe UI", 8))
        lbl_st.setStyleSheet(f"color:{TEXT_MUTED};")
        lbl_st.setWordWrap(True)
        self.template_labels[color] = lbl_st
        sl.addWidget(lbl_st, 1)

        btn_load = QPushButton("Carregar")
        btn_load.setFixedSize(94, 26)
        btn_load.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_load.setFont(QFont("Segoe UI", 8, QFont.Weight.Bold))
        btn_load.setStyleSheet(
            f"QPushButton{{background:{ACCENT_LIGHT};color:{ACCENT};"
            f"border:1px solid {ACCENT};border-radius:5px;}}"
            f"QPushButton:hover{{background:{ACCENT};color:white;}}"
        )
        btn_load.clicked.connect(lambda _, c=color: self._load_template(c))
        sl.addWidget(btn_load)

        return sub

    def _build_add_name_group(self) -> QGroupBox:
        grp = self._grpbox("Adicionar Nome")
        gl = QVBoxLayout(grp)
        gl.setSpacing(8)

        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Digite o nome aqui...")
        self.name_input.setFont(QFont("Segoe UI", 10))
        self.name_input.setFixedHeight(38)
        self.name_input.setStyleSheet(
            f"QLineEdit{{border:2px solid {BORDER};border-radius:8px;"
            f"padding:0 10px;background:white;color:{TEXT_PRIMARY};}}"
            f"QLineEdit:focus{{border-color:{ACCENT};}}"
        )
        self.name_input.returnPressed.connect(self._add_name)
        gl.addWidget(self.name_input)
        gl.addWidget(
            self._btn(
                "＋  Adicionar Nome",
                38,
                f"QPushButton{{background:{ACCENT};color:white;border-radius:8px;"
                f"border:none;font-size:10pt;font-weight:bold;}}"
                f"QPushButton:hover{{background:{ACCENT_DARK};}}",
                self._add_name,
            )
        )
        return grp

    def _right_panel(self) -> QWidget:
        panel = QWidget()
        panel.setStyleSheet(f"background:{BG_MAIN};")
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        hdr = QFrame()
        hdr.setFixedHeight(56)
        hdr.setStyleSheet(f"background:{BG_SIDEBAR};border-bottom:1px solid {BORDER};")
        hl = QHBoxLayout(hdr)
        hl.setContentsMargins(20, 0, 20, 0)

        ht = QLabel("Visualização de Páginas")
        ht.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        ht.setStyleSheet(f"color:{TEXT_PRIMARY};")
        hl.addWidget(ht)

        self.page_count_lbl = QLabel("")
        self.page_count_lbl.setFont(QFont("Segoe UI", 9))
        self.page_count_lbl.setStyleSheet(f"color:{TEXT_MUTED};")
        self.page_count_lbl.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        hl.addWidget(self.page_count_lbl, 1)
        lay.addWidget(hdr)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll.setStyleSheet(
            f"QScrollArea{{background:{BG_MAIN};}}"
            f"QScrollBar:vertical{{width:8px;background:transparent;}}"
            f"QScrollBar::handle:vertical{{background:{BORDER};border-radius:4px;min-height:20px;}}"
        )
        self.scroll_content = QWidget()
        self.scroll_content.setStyleSheet(f"background:{BG_MAIN};")
        self.scroll_lay = QVBoxLayout(self.scroll_content)
        self.scroll_lay.setContentsMargins(20, 20, 20, 20)
        self.scroll_lay.setSpacing(0)
        self.scroll_lay.addStretch()
        self.scroll.setWidget(self.scroll_content)
        lay.addWidget(self.scroll)

        self._show_empty()
        return panel

    # ── Helpers de seleção ───────────────────────────────────────────────────
    def _selected_color(self) -> str:
        for btn in self.radio_group.buttons():
            if btn.isChecked():
                return btn.text().strip()
        return "Verde"

    # ── Ações — delegam toda lógica ao DocxService ───────────────────────────
    def _load_template(self, color: str):
        path, _ = QFileDialog.getOpenFileName(
            self, f"Carregar modelo {color}", "", "Word Document (*.docx)"
        )
        if not path:
            return
        try:
            self.docx_service.validate_template(path)
        except ValueError as e:
            QMessageBox.critical(self, "Modelo inválido", str(e))
            return

        self.templates[color] = path
        fname = os.path.basename(path)
        self.template_labels[color].setText(f"{fname}")
        self.template_labels[color].setStyleSheet(
            f"color:{SUCCESS};font-size:8pt;font-weight:bold;"
        )

    def _add_name(self):
        name = self.name_input.text().strip()
        if not name:
            return
        self.docx_service.add_name(name)
        self.name_input.clear()
        self.name_input.setFocus()
        self._refresh()

    def _edit_name(self, idx: int):
        dlg = QDialog(self)
        dlg.setWindowTitle("Editar Nome")
        dlg.setFixedSize(360, 140)
        dlg.setStyleSheet(f"background:white;color:{TEXT_PRIMARY};")
        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(12)
        lay.addWidget(QLabel("Novo nome:"))

        edit = QLineEdit(self.docx_service.names[idx])
        edit.setFixedHeight(36)
        edit.setFont(QFont("Segoe UI", 10))
        edit.setStyleSheet(
            f"border:2px solid {ACCENT};border-radius:8px;padding:0 10px;"
        )
        edit.selectAll()
        lay.addWidget(edit)

        btns = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btns.button(QDialogButtonBox.StandardButton.Ok).setText("Salvar")
        btns.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        edit.returnPressed.connect(dlg.accept)
        lay.addWidget(btns)

        if dlg.exec() == QDialog.DialogCode.Accepted:
            new_name = edit.text().strip()
            if new_name:
                self.docx_service.edit_name(idx, new_name)
                self._refresh()

    def _delete_name(self, idx: int):
        self.docx_service.delete_name(idx)
        self._refresh()

    def _clear_all(self):
        if not self.docx_service.total_names:
            return
        if (
            QMessageBox.question(
                self,
                "Confirmar",
                "Deseja remover todos os nomes?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            == QMessageBox.StandardButton.Yes
        ):
            self.docx_service.clear_names()
            self._refresh()

    def _generate(self):
        if not self.docx_service.total_names:
            QMessageBox.warning(
                self, "Lista vazia", "Adicione pelo menos um nome antes de gerar."
            )
            return

        color = self._selected_color()
        if color not in self.templates:
            QMessageBox.warning(
                self,
                "Modelo não carregado",
                f"Carregue primeiro o arquivo modelo para a cor <b>{color}</b>.\n\n"
                f'Clique no botão  Carregar  ao lado de "{color}".',
            )
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar arquivo gerado",
            f"crachas_{color.lower()}.docx",
            "Word Document (*.docx)",
        )
        if not save_path:
            return

        try:
            docx_bytes = self.docx_service.generate_document(self.templates[color])
            with open(save_path, "wb") as f:
                f.write(docx_bytes)
            n = self.docx_service.total_names
            pages = self.docx_service.total_pages
            QMessageBox.information(
                self,
                "Sucesso!",
                f"Arquivo gerado com sucesso!\n\n {save_path}\n\n"
                f"Total: {n} crachá{'s' if n!=1 else ''} em "
                f"{pages} página{'s' if pages!=1 else ''}.",
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Erro ao gerar arquivo", f"Ocorreu um erro:\n{str(e)}"
            )

    # ── Atualização da lista visual ──────────────────────────────────────────
    def _refresh(self):
        while self.scroll_lay.count() > 1:
            item = self.scroll_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self.docx_service.total_names:
            self._show_empty()
            self.count_lbl.setText("0 nomes adicionados")
            self.page_count_lbl.setText("")
            return

        n = self.docx_service.total_names
        pages = self.docx_service.total_pages

        self.count_lbl.setText(
            f"{n} nome{'s' if n!=1 else ''} adicionado{'s' if n!=1 else ''}"
        )
        self.page_count_lbl.setText(
            f"{pages} página{'s' if pages!=1 else ''} · {n} crachá{'s' if n!=1 else ''}"
        )

        for page_idx in range(pages):
            grp = PageGroup(
                page_num=page_idx + 1,
                names=self.docx_service.get_page_names(page_idx),
                global_start=self.docx_service.get_global_index(page_idx),
                on_edit=self._edit_name,
                on_delete=self._delete_name,
            )
            self.scroll_lay.insertWidget(self.scroll_lay.count() - 1, grp)

    def _show_empty(self):
        while self.scroll_lay.count() > 1:
            item = self.scroll_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setAlignment(Qt.AlignmentFlag.AlignCenter)

        msg = QLabel(
            "Nenhum nome adicionado ainda.\n"
            "Digite um nome à esquerda e pressione Enter ou clique em Adicionar."
        )
        msg.setFont(QFont("Segoe UI", 11))
        msg.setStyleSheet(f"color:{TEXT_MUTED};")
        msg.setAlignment(Qt.AlignmentFlag.AlignCenter)
        vl.addWidget(msg)

        self.scroll_lay.insertWidget(0, w)
