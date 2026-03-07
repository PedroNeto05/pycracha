import copy
import zipfile
import io
from lxml import etree

from service.constants import CRACHAS_POR_PAGINA, W


class DocxService:

    def gerar_docx_de_modelo(self, template_path, names):
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
