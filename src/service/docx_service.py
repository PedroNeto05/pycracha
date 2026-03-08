import copy
import zipfile
import io
from lxml import etree
from service.constants import CRACHAS_POR_PAGINA, W


class DocxService:
    def __init__(self):
        self.names: list[str] = []

    def add_name(self, name: str):
        self.names.append(name)

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
        pages = (len(self.names) + CRACHAS_POR_PAGINA - 1) // CRACHAS_POR_PAGINA

        for page_idx in range(pages):
            page_names = self._get_page_names(page_idx)
            page_tbl = self._fill_page(orig_tbl, page_names)
            new_body.append(page_tbl)

            if page_idx < pages - 1:
                new_body.append(self._page_break())

        if sectPr is not None:
            new_body.append(copy.deepcopy(sectPr))

        return new_body

    def _get_page_names(self, page_idx: int) -> list[str]:
        start = page_idx * CRACHAS_POR_PAGINA
        end = min(start + CRACHAS_POR_PAGINA, len(self.names))
        return self.names[start:end]

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
