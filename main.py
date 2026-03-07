#!/usr/bin/env python3
"""
Gerador de Crachás ECRI
Aplicativo desktop para gerenciar e gerar crachás a partir de arquivos modelo .docx
"""

import copy
import io
import os
import sys
import zipfile

from lxml import etree
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication,
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

# ── Constantes ────────────────────────────────────────────────────────────────
CRACHAS_POR_PAGINA = 6
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
COLOR_OPTIONS = ["Amarelo", "Verde", "Vermelho"]
COLOR_DOTS = {"Amarelo": "#F5C518", "Verde": "#27AE60", "Vermelho": "#C0392B"}

# ── Paleta ────────────────────────────────────────────────────────────────────
BG_MAIN = "#F0F4F8"
BG_SIDEBAR = "#FFFFFF"
BG_PAGE = "#E8EEF4"
BG_CARD = "#FFFFFF"
ACCENT = "#2C6EAB"
ACCENT_DARK = "#1A4F80"
ACCENT_LIGHT = "#EBF3FB"
DANGER = "#D94040"
SUCCESS = "#2E7D32"
TEXT_PRIMARY = "#1A1A2E"
TEXT_MUTED = "#6B7A8D"
BORDER = "#D0D8E4"


# ── Geração do .docx ──────────────────────────────────────────────────────────
def gerar_docx_de_modelo(template_path, names):
    with zipfile.ZipFile(template_path, "r") as z:
        all_files = {n: z.read(n) for n in z.namelist()}
    doc_xml = all_files["word/document.xml"]
    orig_tree = etree.fromstring(doc_xml)
    orig_body = orig_tree.find(f"{W}body")
    orig_tbl = orig_body.find(f"{W}tbl")
    sectPr = orig_body.find(f"{W}sectPr")

    if orig_tbl is None:
        raise ValueError("Nenhuma tabela encontrada no arquivo modelo.")

    new_body = etree.Element(f"{W}body")
    pages = (len(names) + CRACHAS_POR_PAGINA - 1) // CRACHAS_POR_PAGINA

    for page_idx in range(pages):
        start = page_idx * CRACHAS_POR_PAGINA
        end = min(start + CRACHAS_POR_PAGINA, len(names))
        page_names = names[start:end]
        page_tbl = copy.deepcopy(orig_tbl)

        for badge_idx, badge_tbl in enumerate(page_tbl.findall(f".//{W}tbl")):
            name = page_names[badge_idx] if badge_idx < len(page_names) else ""
            for t_elem in badge_tbl.findall(f".//{W}t"):
                if t_elem.text and "BRASIL" in t_elem.text:
                    t_elem.text = name.upper() if name else ""
                    break

        new_body.append(page_tbl)

        if page_idx < pages - 1:
            pb_p = etree.SubElement(new_body, f"{W}p")
            pb_r = etree.SubElement(pb_p, f"{W}r")
            pb_br = etree.SubElement(pb_r, f"{W}br")
            pb_br.set(f"{W}type", "page")

    if sectPr is not None:
        new_body.append(copy.deepcopy(sectPr))

    orig_tree.remove(orig_body)
    orig_tree.append(new_body)
    new_xml = etree.tostring(
        orig_tree, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for fname, data in all_files.items():
            zout.writestr(fname, new_xml if fname == "word/document.xml" else data)
    return buf.getvalue()


# ── Widgets auxiliares ────────────────────────────────────────────────────────
class NameCard(QFrame):
    def __init__(self, name, index, on_edit, on_delete, parent=None):
        super().__init__(parent)
        self.index = index
        self.setObjectName("NameCard")
        self.setStyleSheet(
            f"QFrame#NameCard{{background:{BG_CARD};border:1px solid {BORDER};border-radius:8px;}}QFrame#NameCard:hover{{border-color:{ACCENT};}}"
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
        lbl.setFont(QFont("Segoe UI", 10, QFont.Medium))
        lbl.setStyleSheet(f"color:{TEXT_PRIMARY};")
        lay.addWidget(lbl, 1)
        for txt, bg, fg, cb in [
            ("✏ Editar", ACCENT_LIGHT, ACCENT, lambda: on_edit(self.index)),
            ("🗑 Apagar", "#FDE8E8", DANGER, lambda: on_delete(self.index)),
        ]:
            b = QPushButton(txt)
            b.setFixedSize(80, 30)
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet(
                f"QPushButton{{background:{bg};color:{fg};border:1px solid {fg};border-radius:6px;font-size:11px;font-weight:600;}}QPushButton:hover{{background:{fg};color:white;}}"
            )
            b.clicked.connect(cb)
            lay.addWidget(b)


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
        l1 = QLabel(f"📄  Página {page_num}")
        l1.setFont(QFont("Segoe UI", 9, QFont.Bold))
        l1.setStyleSheet("color:white;")
        hl.addWidget(l1)
        l2 = QLabel(f"{len(names)} crachá{'s' if len(names)!=1 else ''}")
        l2.setFont(QFont("Segoe UI", 9))
        l2.setStyleSheet("color:rgba(255,255,255,0.8);")
        l2.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
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


# ── Janela principal ──────────────────────────────────────────────────────────
class GeradorCrachas(QMainWindow):
    def __init__(self):
        super().__init__()
        self.names = []
        self.templates = {}
        self.setWindowTitle("Gerador de Crachás ECRI")
        self.setMinimumSize(980, 660)
        self.resize(1120, 740)
        self._build_ui()
        self.setStyleSheet(f"QMainWindow{{background:{BG_MAIN};}}")

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
        f.setFrameShape(QFrame.HLine)
        f.setStyleSheet(f"color:{BORDER};")
        return f

    def _grpbox(self, title):
        gb = QGroupBox(title)
        gb.setFont(QFont("Segoe UI", 10, QFont.Bold))
        gb.setStyleSheet(
            f"QGroupBox{{color:{TEXT_PRIMARY};border:1px solid {BORDER};border-radius:8px;margin-top:10px;padding-top:8px;}}QGroupBox::title{{subcontrol-origin:margin;left:12px;padding:0 4px;}}"
        )
        return gb

    def _btn(self, text, height, style, cb):
        b = QPushButton(text)
        b.setFixedHeight(height)
        b.setCursor(Qt.PointingHandCursor)
        b.setStyleSheet(style)
        b.clicked.connect(cb)
        return b

    def _left_panel(self):
        panel = QFrame()
        panel.setFixedWidth(340)
        panel.setObjectName("LP")
        panel.setStyleSheet(
            f"QFrame#LP{{background:{BG_SIDEBAR};border-right:1px solid {BORDER};}}"
        )
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(20, 24, 20, 24)
        lay.setSpacing(14)

        t = QLabel("🎪 Gerador de Crachás")
        t.setFont(QFont("Segoe UI", 14, QFont.Bold))
        t.setStyleSheet(f"color:{TEXT_PRIMARY};")
        lay.addWidget(t)
        s = QLabel("ECRI – Encontro de Crianças com Cristo")
        s.setFont(QFont("Segoe UI", 9))
        s.setStyleSheet(f"color:{TEXT_MUTED};")
        lay.addWidget(s)
        lay.addWidget(self._sep())

        # ── Modelos ──
        grp = self._grpbox("Modelos de Crachá")
        gl = QVBoxLayout(grp)
        gl.setSpacing(6)
        self.radio_group = QButtonGroup(self)
        self.template_labels = {}

        for i, color in enumerate(COLOR_OPTIONS):
            dot_color = COLOR_DOTS[color]
            row = QWidget()
            rl = QHBoxLayout(row)
            rl.setContentsMargins(0, 0, 0, 0)
            rl.setSpacing(6)
            rb = QRadioButton(f"  {color}")
            rb.setFont(QFont("Segoe UI", 10))
            rb.setStyleSheet(
                f"QRadioButton{{color:{TEXT_PRIMARY};padding:4px;border-radius:6px;}}QRadioButton:hover{{background:{ACCENT_LIGHT};}}"
            )
            self.radio_group.addButton(rb)
            rl.addWidget(rb, 1)
            dot = QLabel("●")
            dot.setFont(QFont("Segoe UI", 16))
            dot.setStyleSheet(f"color:{dot_color};padding-right:4px;")
            dot.setFixedWidth(22)
            rl.addWidget(dot)
            gl.addWidget(row)

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
            btn_load = QPushButton("📂 Carregar")
            btn_load.setFixedSize(94, 26)
            btn_load.setCursor(Qt.PointingHandCursor)
            btn_load.setFont(QFont("Segoe UI", 8, QFont.Bold))
            btn_load.setStyleSheet(
                f"QPushButton{{background:{ACCENT_LIGHT};color:{ACCENT};border:1px solid {ACCENT};border-radius:5px;}}QPushButton:hover{{background:{ACCENT};color:white;}}"
            )
            btn_load.clicked.connect(lambda _, c=color: self._load_template(c))
            sl.addWidget(btn_load)
            gl.addWidget(sub)

            if i < len(COLOR_OPTIONS) - 1:
                div = QFrame()
                div.setFrameShape(QFrame.HLine)
                div.setStyleSheet(f"color:{BORDER};margin:2px 4px;")
                gl.addWidget(div)

        self.radio_group.buttons()[1].setChecked(True)
        lay.addWidget(grp)

        # ── Adicionar nome ──
        grp2 = self._grpbox("Adicionar Nome")
        gl2 = QVBoxLayout(grp2)
        gl2.setSpacing(8)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Digite o nome aqui...")
        self.name_input.setFont(QFont("Segoe UI", 10))
        self.name_input.setFixedHeight(38)
        self.name_input.setStyleSheet(
            f"QLineEdit{{border:2px solid {BORDER};border-radius:8px;padding:0 10px;background:white;color:{TEXT_PRIMARY};}}QLineEdit:focus{{border-color:{ACCENT};}}"
        )
        self.name_input.returnPressed.connect(self._add_name)
        gl2.addWidget(self.name_input)
        gl2.addWidget(
            self._btn(
                "＋  Adicionar Nome",
                38,
                f"QPushButton{{background:{ACCENT};color:white;border-radius:8px;border:none;font-size:10pt;font-weight:bold;}}QPushButton:hover{{background:{ACCENT_DARK};}}",
                self._add_name,
            )
        )
        lay.addWidget(grp2)

        self.count_lbl = QLabel("0 nomes adicionados")
        self.count_lbl.setFont(QFont("Segoe UI", 9))
        self.count_lbl.setStyleSheet(f"color:{TEXT_MUTED};")
        self.count_lbl.setAlignment(Qt.AlignCenter)
        lay.addWidget(self.count_lbl)

        lay.addStretch()
        lay.addWidget(self._sep())

        lay.addWidget(
            self._btn(
                "🗂  Gerar Arquivo .docx",
                46,
                f"QPushButton{{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 {SUCCESS},stop:1 #1B5E20);color:white;border-radius:10px;border:none;font-size:11pt;font-weight:bold;}}QPushButton:hover{{background:#2E7D32;}}QPushButton:disabled{{background:#B0BEC5;}}",
                self._generate,
            )
        )
        lay.addWidget(
            self._btn(
                "🗑  Limpar Tudo",
                34,
                f"QPushButton{{background:transparent;color:{DANGER};border:1px solid {DANGER};border-radius:8px;font-size:9pt;}}QPushButton:hover{{background:#FDE8E8;}}",
                self._clear_all,
            )
        )
        return panel

    def _right_panel(self):
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
        ht.setFont(QFont("Segoe UI", 12, QFont.Bold))
        ht.setStyleSheet(f"color:{TEXT_PRIMARY};")
        hl.addWidget(ht)
        self.page_count_lbl = QLabel("")
        self.page_count_lbl.setFont(QFont("Segoe UI", 9))
        self.page_count_lbl.setStyleSheet(f"color:{TEXT_MUTED};")
        self.page_count_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        hl.addWidget(self.page_count_lbl, 1)
        lay.addWidget(hdr)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setStyleSheet(
            f"QScrollArea{{background:{BG_MAIN};}}QScrollBar:vertical{{width:8px;background:transparent;}}QScrollBar::handle:vertical{{background:{BORDER};border-radius:4px;min-height:20px;}}"
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

    def _selected_color(self):
        for btn in self.radio_group.buttons():
            if btn.isChecked():
                return btn.text().strip()
        return "Verde"

    def _load_template(self, color):
        path, _ = QFileDialog.getOpenFileName(
            self, f"Carregar modelo {color}", "", "Word Document (*.docx)"
        )
        if not path:
            return
        try:
            with zipfile.ZipFile(path, "r") as z:
                if "word/document.xml" not in z.namelist():
                    raise ValueError("Arquivo .docx inválido.")
                doc_xml = z.read("word/document.xml")
            tree = etree.fromstring(doc_xml)
            body = tree.find(f"{W}body")
            outer = body.find(f"{W}tbl") if body is not None else None
            if outer is None:
                raise ValueError("Nenhuma tabela encontrada no documento.")
            inner = outer.findall(f".//{W}tbl")
            if not inner:
                raise ValueError(
                    "Estrutura de crachás não encontrada (tabelas internas ausentes)."
                )
            brasil = sum(
                1
                for tbl in inner
                for t in tbl.findall(f".//{W}t")
                if t.text and "BRASIL" in t.text
            )
            if brasil == 0:
                raise ValueError(
                    'O modelo não contém o marcador "BRASIL".\nVerifique se é o arquivo correto.'
                )
        except zipfile.BadZipFile:
            QMessageBox.critical(self, "Erro", "O arquivo não é um .docx válido.")
            return
        except ValueError as e:
            QMessageBox.critical(self, "Modelo inválido", str(e))
            return
        self.templates[color] = path
        fname = os.path.basename(path)
        self.template_labels[color].setText(f"✅ {fname}")
        self.template_labels[color].setStyleSheet(
            f"color:{SUCCESS};font-size:8pt;font-weight:bold;"
        )

    def _add_name(self):
        name = self.name_input.text().strip()
        if not name:
            return
        self.names.append(name)
        self.name_input.clear()
        self.name_input.setFocus()
        self._refresh()

    def _edit_name(self, idx):
        dlg = QDialog(self)
        dlg.setWindowTitle("Editar Nome")
        dlg.setFixedSize(360, 140)
        dlg.setStyleSheet(f"background:white;color:{TEXT_PRIMARY};")
        lay = QVBoxLayout(dlg)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(12)
        lay.addWidget(QLabel("Novo nome:"))
        edit = QLineEdit(self.names[idx])
        edit.setFixedHeight(36)
        edit.setFont(QFont("Segoe UI", 10))
        edit.setStyleSheet(
            f"border:2px solid {ACCENT};border-radius:8px;padding:0 10px;"
        )
        edit.selectAll()
        lay.addWidget(edit)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Salvar")
        btns.button(QDialogButtonBox.Cancel).setText("Cancelar")
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        lay.addWidget(btns)
        edit.returnPressed.connect(dlg.accept)
        if dlg.exec_() == QDialog.Accepted:
            new = edit.text().strip()
            if new:
                self.names[idx] = new
                self._refresh()

    def _delete_name(self, idx):
        del self.names[idx]
        self._refresh()

    def _clear_all(self):
        if not self.names:
            return
        if (
            QMessageBox.question(
                self,
                "Confirmar",
                "Deseja remover todos os nomes?",
                QMessageBox.Yes | QMessageBox.No,
            )
            == QMessageBox.Yes
        ):
            self.names.clear()
            self._refresh()

    def _refresh(self):
        while self.scroll_lay.count() > 1:
            item = self.scroll_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        if not self.names:
            self._show_empty()
            self.count_lbl.setText("0 nomes adicionados")
            self.page_count_lbl.setText("")
            return
        n = len(self.names)
        pages = (n + CRACHAS_POR_PAGINA - 1) // CRACHAS_POR_PAGINA
        self.count_lbl.setText(
            f"{n} nome{'s' if n!=1 else ''} adicionado{'s' if n!=1 else ''}"
        )
        self.page_count_lbl.setText(
            f"{pages} página{'s' if pages!=1 else ''} · {n} crachá{'s' if n!=1 else ''}"
        )
        for p in range(pages):
            start = p * CRACHAS_POR_PAGINA
            end = min(start + CRACHAS_POR_PAGINA, n)
            grp = PageGroup(
                p + 1, self.names[start:end], start, self._edit_name, self._delete_name
            )
            self.scroll_lay.insertWidget(self.scroll_lay.count() - 1, grp)

    def _show_empty(self):
        while self.scroll_lay.count() > 1:
            item = self.scroll_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        w = QWidget()
        vl = QVBoxLayout(w)
        vl.setAlignment(Qt.AlignCenter)
        icon = QLabel("📋")
        icon.setFont(QFont("Segoe UI", 40))
        icon.setAlignment(Qt.AlignCenter)
        vl.addWidget(icon)
        msg = QLabel(
            "Nenhum nome adicionado ainda.\nDigite um nome à esquerda e pressione Enter ou clique em Adicionar."
        )
        msg.setFont(QFont("Segoe UI", 11))
        msg.setStyleSheet(f"color:{TEXT_MUTED};")
        msg.setAlignment(Qt.AlignCenter)
        vl.addWidget(msg)
        self.scroll_lay.insertWidget(0, w)

    def _generate(self):
        if not self.names:
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
                f'Clique no botão  📂 Carregar  ao lado de "{color}".',
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
            docx_bytes = gerar_docx_de_modelo(self.templates[color], self.names)
            with open(save_path, "wb") as f:
                f.write(docx_bytes)
            n = len(self.names)
            pages = (n + CRACHAS_POR_PAGINA - 1) // CRACHAS_POR_PAGINA
            QMessageBox.information(
                self,
                "Sucesso! 🎉",
                f"Arquivo gerado com sucesso!\n\n📁 {save_path}\n\n"
                f"Total: {n} crachá{'s' if n!=1 else ''} em {pages} página{'s' if pages!=1 else ''}.",
            )
        except Exception as e:
            QMessageBox.critical(
                self, "Erro ao gerar arquivo", f"Ocorreu um erro:\n{str(e)}"
            )


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Gerador de Crachás ECRI")
    app.setStyle("Fusion")
    w = GeradorCrachas()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
