import copy
import re
import unicodedata
import io
import zipfile

import pandas as pd
from lxml import etree

from service.constants import CRACHAS_POR_PAGINA, W

_EXCLUDED_COLUMN_KEYWORDS = {"nome", "telefone", "email", "foto", "data", "endereço"}


class DocxService:
    def __init__(self):
        self._names: list[str] = []

    def add_name(self, name: str):
        self._names.append(name)

    def edit_name(self, idx: int, new_name: str):
        self._names[idx] = new_name

    def delete_name(self, idx: int):
        del self._names[idx]

    def clear_names(self):
        self._names.clear()

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

    def validate_template(self, path: str):
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

    def import_from_spreadsheet(
        self,
        file_path: str,
        name_columns: list[str],
        abbreviate: bool,
        filters: list[dict[str, str]] | None = None,
    ):
        df = self._read_spreadsheet(file_path)
        self._validate_columns(df, name_columns)

        if filters:
            df = self.apply_filters(df, filters)

        for _, row in df.iterrows():
            full_name = self._build_full_name(row, name_columns)
            if not full_name:
                continue
            if abbreviate:
                full_name = self._abbreviate_name(full_name)
            self._names.append(full_name.title())

    def _read_spreadsheet(self, file_path: str) -> "pd.DataFrame":
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)

        return self._consolidate_columns(df)

    def _validate_columns(
        self,
        df: "pd.DataFrame",
        name_columns: list[str],
    ):
        if not name_columns:
            raise ValueError("Nenhuma coluna de nome foi selecionada.")

        for col in name_columns:
            if col not in df.columns:
                raise ValueError(
                    f'Coluna "{col}" não encontrada na planilha.\n'
                    f"Colunas disponíveis: {', '.join(df.columns)}"
                )

    def _build_full_name(
        self,
        row: "pd.Series",
        name_columns: list[str],
    ) -> str:
        parts = []
        for col in name_columns:
            val = str(row[col]).strip() if pd.notna(row[col]) else ""
            if val:
                parts.append(val)

        if not parts:
            return ""

        return " ".join(parts).title()

    def _abbreviate_name(self, name: str) -> str:
        parts = name.split()
        if len(parts) <= 2:
            return name
        return f"{parts[0]} {parts[-1]}"

    def get_columns_from_spreadsheet(self, file_path: str) -> list[str]:
        df = self._read_spreadsheet(file_path)
        return list(df.columns)

    def get_filterable_columns(self, file_path: str) -> list[str]:
        all_columns = self.get_columns_from_spreadsheet(file_path)
        return [
            col
            for col in all_columns
            if not any(kw in col.lower() for kw in _EXCLUDED_COLUMN_KEYWORDS)
        ]

    def apply_filters(
        self,
        df: "pd.DataFrame",
        filters: list[dict[str, str]],
    ) -> "pd.DataFrame":
        """
        Aplica filtros aditivos (AND) com normalização.
        Recebe o df já lido para evitar reler o arquivo.
        """
        for f in filters:
            column = f["column"]
            term = self._normalize(f["value"])

            if column not in df.columns or not term:
                continue

            df = df[df[column].apply(lambda cell: term in self._normalize(str(cell)))]

        return df

    def _normalize(self, text: str) -> str:
        """Lowercase + remove acentos + remove espaços."""
        text = text.lower().strip()
        text = unicodedata.normalize("NFD", text)
        text = "".join(c for c in text if unicodedata.category(c) != "Mn")
        text = text.replace(" ", "")
        return text

    def _consolidate_columns(self, df: "pd.DataFrame") -> "pd.DataFrame":
        """
        Agrupa colunas duplicadas geradas pelo pandas (ex: Equipe, Equipe.1, Equipe.2)
        e consolida os valores pegando o primeiro valor não nulo da linha.
        """
        base_names = {}
        for col in df.columns:
            # Remove o sufixo '.n' (ex: .1, .2, .12) do final do nome da coluna
            base_name = re.sub(r"\.\d+$", "", str(col))
            if base_name not in base_names:
                base_names[base_name] = []
            base_names[base_name].append(col)

        new_df = pd.DataFrame()
        for base_name, cols in base_names.items():
            if len(cols) == 1:
                # Se não tem duplicata, apenas copia a coluna original
                new_df[base_name] = df[cols[0]]
            else:
                # Se tem duplicata, usamos combine_first iterativamente para
                # preencher os vazios com o primeiro valor não nulo encontrado
                consolidated = df[cols[0]]
                for col in cols[1:]:
                    consolidated = consolidated.combine_first(df[col])
                new_df[base_name] = consolidated

        return new_df
