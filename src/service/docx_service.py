import copy
import io
import zipfile

from lxml import etree

from service.constants import CRACHAS_POR_PAGINA, W


class DocxService:
    def __init__(self):
        self._names: list[str] = []

    # ── Gerenciamento de nomes ───────────────────────────────────────────────
    def add_name(self, name: str):
        self._names.append(name)

    def edit_name(self, idx: int, new_name: str):
        self._names[idx] = new_name

    def delete_name(self, idx: int):
        del self._names[idx]

    def clear_names(self):
        self._names.clear()

    # ── Consultas ────────────────────────────────────────────────────────────
    @property
    def names(self) -> list[str]:
        return self._names

    @property
    def total_names(self) -> int:
        return len(self._names)

    @property
    def total_pages(self) -> int:
        if not self._names:
            return 0
        return (self.total_names + CRACHAS_POR_PAGINA - 1) // CRACHAS_POR_PAGINA

    def get_page_names(self, page_idx: int) -> list[str]:
        start = page_idx * CRACHAS_POR_PAGINA
        end = min(start + CRACHAS_POR_PAGINA, self.total_names)
        return self._names[start:end]

    def get_global_index(self, page_idx: int) -> int:
        return page_idx * CRACHAS_POR_PAGINA

    # ── Validação do template ────────────────────────────────────────────────
    def validate_template(self, path: str):
        """Valida o arquivo modelo. Lança ValueError com mensagem legível se inválido."""
        try:
            with zipfile.ZipFile(path, "r") as z:
                if "word/document.xml" not in z.namelist():
                    raise ValueError("Arquivo .docx inválido.")
                doc_xml = z.read("word/document.xml")
        except zipfile.BadZipFile:
            raise ValueError("O arquivo não é um .docx válido.")

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
                'O modelo não contém o marcador "BRASIL".\n'
                "Verifique se é o arquivo correto."
            )

    # ── Geração do documento ─────────────────────────────────────────────────
    def generate_document(self, template_path: str) -> bytes:
        all_files = self._read_zip(template_path)
        orig_tree = self._parse_xml(all_files)
        orig_body, orig_tbl, sectPr = self._extract_body_parts(orig_tree)
        new_body = self._build_body(orig_tbl, sectPr)
        self._replace_body(orig_tree, orig_body, new_body)
        return self._pack_zip(all_files, orig_tree)

    def _read_zip(self, template_path: str) -> dict:
        with zipfile.ZipFile(template_path, "r") as z:
            return {name: z.read(name) for name in z.namelist()}

    def _parse_xml(self, all_files: dict):
        return etree.fromstring(all_files["word/document.xml"])

    def _extract_body_parts(self, tree):
        body = tree.find(f"{W}body")
        tbl = body.find(f"{W}tbl")
        sect_pr = body.find(f"{W}sectPr")

        if tbl is None:
            raise ValueError("Nenhuma tabela encontrada no arquivo modelo.")

        return body, tbl, sect_pr

    def _build_body(self, orig_tbl, sectPr) -> etree._Element:
        new_body = etree.Element(f"{W}body")

        for page_idx in range(self.total_pages):
            page_names = self.get_page_names(page_idx)
            new_body.append(self._fill_page(orig_tbl, page_names))

            if page_idx < self.total_pages - 1:
                new_body.append(self._page_break())

        if sectPr is not None:
            new_body.append(copy.deepcopy(sectPr))

        return new_body

    def _fill_page(self, orig_tbl, page_names: list[str]):
        page_tbl = copy.deepcopy(orig_tbl)

        for badge_idx, badge_tbl in enumerate(page_tbl.findall(f".//{W}tbl")):
            name = page_names[badge_idx] if badge_idx < len(page_names) else ""
            self._fill_badge(badge_tbl, name)

        return page_tbl

    def _fill_badge(self, badge_tbl, name: str):
        for t_elem in badge_tbl.findall(f".//{W}t"):
            if t_elem.text and "BRASIL" in t_elem.text:
                t_elem.text = name.upper() if name else ""
                break

    def _page_break(self) -> etree._Element:
        pb_p = etree.Element(f"{W}p")
        pb_r = etree.SubElement(pb_p, f"{W}r")
        pb_br = etree.SubElement(pb_r, f"{W}br")
        pb_br.set(f"{W}type", "page")
        return pb_p

    def _replace_body(self, tree, old_body, new_body):
        tree.remove(old_body)
        tree.append(new_body)

    def _pack_zip(self, all_files: dict, tree) -> bytes:
        new_xml = etree.tostring(
            tree, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for fname, data in all_files.items():
                zout.writestr(fname, new_xml if fname == "word/document.xml" else data)
        return buf.getvalue()
