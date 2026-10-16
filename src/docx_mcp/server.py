import base64
import io
import os
from pathlib import Path
from mcp.server.mcpserver import MCPServer
from copy import deepcopy
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm, Mm, Length
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_UNDERLINE, WD_BREAK, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION_START, WD_ORIENTATION
import latex2mathml.converter
import mathml2omml

mcp = MCPServer("DocxEditor")


def validate_docx_path(file_path: str, must_exist: bool = False) -> Path:
    if not file_path or not file_path.strip():
        raise ValueError("Path file tidak boleh kosong.")

    resolved_path = Path(file_path).expanduser().resolve()

    if resolved_path.suffix.lower() != ".docx":
        raise ValueError(
            f"Operasi ditolak: '{resolved_path.name}' bukan file .docx yang valid."
        )

    allowed_dir = os.environ.get("DOCX_MCP_ALLOWED_DIR")
    if allowed_dir:
        allowed_base = Path(allowed_dir).expanduser().resolve()
        if not resolved_path.is_relative_to(allowed_base):
            raise PermissionError(
                f"Akses ditolak: File berada di luar direktori yang diizinkan ({allowed_base})."
            )

    if must_exist and not resolved_path.is_file():
        raise FileNotFoundError(f"File '{resolved_path}' tidak ditemukan.")

    return resolved_path


def validate_image_path(image_path: str) -> Path:
    if not image_path or not image_path.strip():
        raise ValueError("Path gambar tidak boleh kosong.")

    resolved_path = Path(image_path).expanduser().resolve()
    valid_exts = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp"}
    if resolved_path.suffix.lower() not in valid_exts:
        raise ValueError(
            f"Format gambar '{resolved_path.suffix}' tidak didukung. Format yang didukung: {', '.join(sorted(valid_exts))}."
        )

    allowed_dir = os.environ.get("DOCX_MCP_ALLOWED_DIR")
    if allowed_dir:
        allowed_base = Path(allowed_dir).expanduser().resolve()
        if not resolved_path.is_relative_to(allowed_base):
            raise PermissionError(
                f"Akses ditolak: File gambar berada di luar direktori yang diizinkan ({allowed_base})."
            )

    if not resolved_path.is_file():
        raise FileNotFoundError(f"File gambar '{resolved_path}' tidak ditemukan.")

    return resolved_path


@mcp.tool()
def read_docx(file_path: str) -> str:
    """Membaca seluruh isi teks dan tabel dari file .docx"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))
    except Exception as e:
        return f"Error: Gagal membaca file {file_path}: {e}"

    content = []

    for i, p in enumerate(doc.paragraphs):
        p_text = p.text.strip()
        pPr = p._p.find(qn("w:pPr"))
        bdr_val = None
        if pPr is not None:
            bdr_elm = pPr.find(qn("w:pBdr"))
            if bdr_elm is not None:
                bottom = bdr_elm.find(qn("w:bottom"))
                if bottom is not None:
                    bdr_val = bottom.get(qn("w:val"))

        math_elms = p._p.xpath(".//m:oMath")
        math_texts = []
        for m in math_elms:
            m_ts = m.xpath('.//*[local-name()="t"]')
            txt = "".join([t.text or "" for t in m_ts]).strip()
            if txt:
                math_texts.append(txt)
        math_desc = " ".join([f"[Rumus: {m}]" for m in math_texts])

        bdr_desc = f" [Garis Pembatas: {bdr_val}]" if (bdr_val and bdr_val != "none") else ""

        if p_text:
            combined = f"{p.text}{bdr_desc}"
            if math_desc:
                combined += f" {math_desc}"
            content.append(f"[P{i}] {combined}")
        elif math_desc:
            content.append(f"[P{i}] {math_desc}{bdr_desc}")
        else:
            has_page_br = bool(p._p.xpath(".//w:br[@w:type='page' or not(@w:type)]"))
            has_col_br = bool(p._p.xpath(".//w:br[@w:type='column']"))
            if has_page_br:
                content.append(f"[P{i}] [Page Break]")
            elif has_col_br:
                content.append(f"[P{i}] [Column Break]")
            elif bdr_desc:
                content.append(f"[P{i}] --- Garis Horizontal ({bdr_val}) ---")

    for t_idx, table in enumerate(doc.tables):
        content.append(f"\n--- Tabel {t_idx} ---")
        for row in table.rows:
            row_text = []
            for cell in row.cells:
                cell_parts = []
                for cp in cell.paragraphs:
                    txt = cp.text.strip()
                    math_elms = cp._p.xpath(".//m:oMath")
                    m_txts = ["".join([t.text or "" for t in m.xpath('.//*[local-name()="t"]')]).strip() for m in math_elms]
                    m_txts = [m for m in m_txts if m]
                    m_desc = " ".join([f"[Rumus: {m}]" for m in m_txts])
                    if txt and m_desc:
                        cell_parts.append(f"{txt} {m_desc}")
                    elif txt:
                        cell_parts.append(txt)
                    elif m_desc:
                        cell_parts.append(m_desc)
                row_text.append(" ".join(cell_parts).strip())
            content.append(" | ".join(row_text))

    if doc.inline_shapes:
        content.append(f"\n--- Total Gambar: {len(doc.inline_shapes)} ---")

    return "\n".join(content)


@mcp.tool()
def create_or_append_docx(file_path: str, paragraphs: list[str], heading: str = "") -> str:
    """Membuat file .docx baru atau menambahkan paragraf ke file yang sudah ada"""
    try:
        path = validate_docx_path(file_path, must_exist=False)
        doc = Document(str(path)) if path.is_file() else Document()

        if heading:
            doc.add_heading(heading, level=1)

        for text in paragraphs:
            doc.add_paragraph(text)

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        return f"Berhasil menyimpan perubahan ke {path}"
    except Exception as e:
        return f"Error: Gagal menyimpan file {file_path}: {e}"


@mcp.tool()
def replace_text_in_docx(file_path: str, old_text: str, new_text: str) -> str:
    """Mencari dan mengganti teks spesifik di dalam paragraf maupun tabel .docx"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))
        count = 0

        for p in doc.paragraphs:
            if old_text in p.text:
                count += p.text.count(old_text)
                p.text = p.text.replace(old_text, new_text)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if old_text in cell.text:
                        count += cell.text.count(old_text)
                        cell.text = cell.text.replace(old_text, new_text)

        doc.save(str(path))
        return f"Berhasil mengganti {count} kemunculan '{old_text}' menjadi '{new_text}' di {path}."
    except Exception as e:
        return f"Error: Gagal memproses file {file_path}: {e}"


@mcp.tool()
def create_table_docx(file_path: str, data: list[list[str]], heading: str = "") -> str:
    """Membuat tabel baru dari matriks data 2D ke dalam file .docx"""
    try:
        if not data or not any(data):
            return "Error: Data tabel tidak boleh kosong."

        path = validate_docx_path(file_path, must_exist=False)
        doc = Document(str(path)) if path.is_file() else Document()

        if heading:
            doc.add_heading(heading, level=1)

        num_rows = len(data)
        num_cols = max(len(row) for row in data)
        try:
            table = doc.add_table(rows=num_rows, cols=num_cols, style="Table Grid")
        except Exception:
            table = doc.add_table(rows=num_rows, cols=num_cols)

        for r_idx, row_values in enumerate(data):
            for c_idx, val in enumerate(row_values):
                table.rows[r_idx].cells[c_idx].text = str(val)

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        return f"Berhasil membuat tabel ({num_rows} baris, {num_cols} kolom) di {path}."
    except Exception as e:
        return f"Error: Gagal membuat tabel di {file_path}: {e}"


@mcp.tool()
def edit_table_docx(file_path: str, table_index: int, row: int, col: int, new_text: str) -> str:
    """Mengubah isi teks pada sel tabel tertentu (0-indexed) dalam file .docx"""
    # shortcut: single-cell edit only; add batch edit or row/column insertion if requested.
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        if row < 0 or row >= len(table.rows):
            return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."

        target_row = table.rows[row]
        if col < 0 or col >= len(target_row.cells):
            return f"Error: Indeks kolom {col} tidak valid. Total kolom pada baris {row}: {len(target_row.cells)}."

        target_row.cells[col].text = new_text
        doc.save(str(path))
        return f"Berhasil mengubah sel [{row}, {col}] pada tabel {table_index} di {path}."
    except Exception as e:
        return f"Error: Gagal mengedit tabel di {file_path}: {e}"


@mcp.tool()
def delete_table_docx(file_path: str, table_index: int) -> str:
    """Menghapus tabel dari file .docx berdasarkan indeks tabel (0-indexed)"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        tbl_elem = table._element
        tbl_elem.getparent().remove(tbl_elem)
        doc.save(str(path))
        return f"Berhasil menghapus tabel {table_index} dari {path}."
    except Exception as e:
        return f"Error: Gagal menghapus tabel di {file_path}: {e}"


@mcp.tool()
def clear_docx(file_path: str) -> str:
    """Menghapus seluruh isi teks dan tabel dari file .docx sehingga menjadi dokumen kosong"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document()
        doc.save(str(path))
        return f"Berhasil mengosongkan seluruh isi file {path}."
    except Exception as e:
        return f"Error: Gagal mengosongkan file {file_path}: {e}"


@mcp.tool()
def add_column_docx(
    file_path: str,
    table_index: int,
    values: list[str] | None = None,
    width_inches: float | None = None,
    insert_index: int | None = None,
) -> str:
    """Menambahkan atau menyisipkan kolom pada tabel docx dengan teks sel opsional, lebar opsional, dan indeks posisi opsional"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        width_val = Inches(width_inches) if width_inches and width_inches > 0 else Inches(1.0)
        table.add_column(width=width_val)

        if insert_index is not None:
            if insert_index < 0 or insert_index > len(table.columns) - 1:
                return f"Error: insert_index {insert_index} tidak valid. Rentang yang diizinkan: 0 sampai {len(table.columns)-1}."
            if insert_index < len(table.columns) - 1:
                for row in table.rows:
                    tcs = row._tr.xpath("w:tc")
                    new_tc = tcs[-1]
                    tcs[insert_index].addprevious(new_tc)
                tblGrid = table._element.xpath("w:tblGrid")
                if tblGrid:
                    cols = tblGrid[0].xpath("w:gridCol")
                    if len(cols) > 1:
                        cols[insert_index].addprevious(cols[-1])
            target_col_idx = insert_index
        else:
            target_col_idx = len(table.columns) - 1

        if width_val:
            table.columns[target_col_idx].width = width_val

        if values:
            for idx, val in enumerate(values[: len(table.rows)]):
                cell = table.rows[idx].cells[target_col_idx]
                cell.text = str(val)
                if width_val:
                    cell.width = width_val
        elif width_val:
            for row in table.rows:
                row.cells[target_col_idx].width = width_val

        doc.save(str(path))
        return f"Berhasil menambahkan kolom baru pada indeks {target_col_idx} di tabel {table_index} pada {path}."
    except Exception as e:
        return f"Error: Gagal menambahkan kolom di {file_path}: {e}"


@mcp.tool()
def resize_table_docx(
    file_path: str,
    table_index: int,
    col_widths: list[float] | None = None,
    col_index: int | None = None,
    width_inches: float | None = None,
) -> str:
    """Mengatur lebar kolom tabel (dalam inci). Bisa mengatur seluruh kolom via col_widths atau satu kolom via col_index & width_inches"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]

        if col_widths is not None:
            for c_idx, w in enumerate(col_widths):
                if c_idx < len(table.columns):
                    w_val = Inches(w)
                    table.columns[c_idx].width = w_val
                    for row in table.rows:
                        if c_idx < len(row.cells):
                            row.cells[c_idx].width = w_val
            doc.save(str(path))
            return f"Berhasil mengatur lebar {min(len(col_widths), len(table.columns))} kolom pada tabel {table_index} di {path}."
        elif col_index is not None and width_inches is not None:
            if col_index < 0 or col_index >= len(table.columns):
                return f"Error: Indeks kolom {col_index} tidak valid. Total kolom: {len(table.columns)}."
            w_val = Inches(width_inches)
            table.columns[col_index].width = w_val
            for row in table.rows:
                if col_index < len(row.cells):
                    row.cells[col_index].width = w_val
            doc.save(str(path))
            return f"Berhasil mengubah lebar kolom {col_index} menjadi {width_inches} inci pada tabel {table_index} di {path}."
        else:
            return "Error: Berikan parameter 'col_widths' (list lebar inci tiap kolom) atau pasangan 'col_index' dan 'width_inches'."
    except Exception as e:
        return f"Error: Gagal mengubah ukuran tabel di {file_path}: {e}"


@mcp.tool()
def set_table_visibility_docx(
    file_path: str,
    table_index: int,
    visible: bool = True,
) -> str:
    """Mengatur visibilitas border tabel (visible=False untuk transparan/tanpa garis, visible=True untuk border garis)"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        tblPr = table._element.xpath("w:tblPr")
        if not tblPr:
            return f"Error: Elemen properti tabel {table_index} tidak ditemukan."

        for existing in tblPr[0].xpath("w:tblBorders"):
            tblPr[0].remove(existing)

        val = "single" if visible else "none"
        sz = "4" if visible else "0"
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="auto"/>'
            f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="auto"/>'
            f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="auto"/>'
            f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="auto"/>'
            f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="auto"/>'
            f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="auto"/>'
            f"</w:tblBorders>"
        )
        tblPr[0].append(borders)

        doc.save(str(path))
        status_text = "tampak (bergaris)" if visible else "transparan (tanpa border)"
        return f"Berhasil mengubah visibilitas tabel {table_index} menjadi {status_text} di {path}."
    except Exception as e:
        return f"Error: Gagal mengatur visibilitas tabel di {file_path}: {e}"


@mcp.tool()
def delete_column_docx(
    file_path: str,
    table_index: int,
    col_index: int,
) -> str:
    """Menghapus satu kolom tertentu dari tabel docx berdasarkan col_index (0-indexed)"""
    # shortcut: assume regular unmerged grid; complex merged cells may shift column indices.
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]

        if len(table.columns) <= 1:
            return f"Error: Tabel hanya memiliki 1 kolom. Gunakan delete_table_docx jika ingin menghapus seluruh tabel."

        if col_index < 0 or col_index >= len(table.columns):
            return f"Error: Indeks kolom {col_index} tidak valid. Total kolom: {len(table.columns)}."

        for row in table.rows:
            if col_index < len(row.cells):
                tc = row.cells[col_index]._element
                tc.getparent().remove(tc)

        tblGrid = table._element.xpath("w:tblGrid")
        if tblGrid:
            gridCols = tblGrid[0].xpath("w:gridCol")
            if col_index < len(gridCols):
                tblGrid[0].remove(gridCols[col_index])

        doc.save(str(path))
        return f"Berhasil menghapus kolom {col_index} dari tabel {table_index} di {path}."
    except Exception as e:
        return f"Error: Gagal menghapus kolom di {file_path}: {e}"


@mcp.tool()
def add_row_docx(
    file_path: str,
    table_index: int,
    values: list[str] | None = None,
    height_inches: float | None = None,
    insert_index: int | None = None,
) -> str:
    """Menambahkan atau menyisipkan baris baru pada tabel docx dengan teks sel opsional, tinggi opsional, dan indeks posisi opsional"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        row = table.add_row()

        if height_inches and height_inches > 0:
            row.height = Inches(height_inches)

        if insert_index is not None:
            if insert_index < 0 or insert_index > len(table.rows) - 1:
                return f"Error: insert_index {insert_index} tidak valid. Rentang yang diizinkan: 0 sampai {len(table.rows)-1}."
            if insert_index < len(table.rows) - 1:
                table.rows[insert_index]._tr.addprevious(row._tr)
            target_row_idx = insert_index
        else:
            target_row_idx = len(table.rows) - 1

        if values:
            for idx, val in enumerate(values[: len(table.rows[target_row_idx].cells)]):
                table.rows[target_row_idx].cells[idx].text = str(val)

        doc.save(str(path))
        return f"Berhasil menambahkan baris baru pada indeks {target_row_idx} di tabel {table_index} pada {path}."
    except Exception as e:
        return f"Error: Gagal menambahkan baris di {file_path}: {e}"


@mcp.tool()
def delete_row_docx(
    file_path: str,
    table_index: int,
    row_index: int,
) -> str:
    """Menghapus satu baris tertentu dari tabel docx berdasarkan row_index (0-indexed)"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]

        if len(table.rows) <= 1:
            return "Error: Tabel hanya memiliki 1 baris. Gunakan delete_table_docx jika ingin menghapus seluruh tabel."

        if row_index < 0 or row_index >= len(table.rows):
            return f"Error: Indeks baris {row_index} tidak valid. Total baris: {len(table.rows)}."

        tr = table.rows[row_index]._element
        tr.getparent().remove(tr)

        doc.save(str(path))
        return f"Berhasil menghapus baris {row_index} dari tabel {table_index} di {path}."
    except Exception as e:
        return f"Error: Gagal menghapus baris di {file_path}: {e}"


@mcp.tool()
def merge_cells_docx(
    file_path: str,
    table_index: int,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> str:
    """Menggabungkan (merge) area sel dari [start_row, start_col] sampai [end_row, end_col] (0-indexed)"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]

        if not (0 <= start_row < len(table.rows) and 0 <= end_row < len(table.rows)):
            return f"Error: Indeks baris tidak valid (rentang baris tabel: 0 sampai {len(table.rows)-1})."

        if not (0 <= start_col < len(table.columns) and 0 <= end_col < len(table.columns)):
            return f"Error: Indeks kolom tidak valid (rentang kolom tabel: 0 sampai {len(table.columns)-1})."

        cell_a = table.cell(start_row, start_col)
        cell_b = table.cell(end_row, end_col)
        cell_a.merge(cell_b)

        doc.save(str(path))
        return f"Berhasil menggabungkan sel dari [{start_row}, {start_col}] ke [{end_row}, {end_col}] pada tabel {table_index} di {path}."
    except Exception as e:
        return f"Error: Gagal menggabungkan sel di {file_path}: {e}"


@mcp.tool()
def set_cell_shading_docx(
    file_path: str,
    table_index: int,
    color_hex: str,
    row: int | None = None,
    col: int | None = None,
) -> str:
    """Memberikan warna latar belakang (shading) pada sel/baris/kolom/tabel dengan format HEX (contoh: 'E0E0E0' atau '#E0E0E0')"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        clean_hex = color_hex.lstrip("#").upper()
        if len(clean_hex) != 6 or not all(c in "0123456789ABCDEF" for c in clean_hex):
            return f"Error: Format warna '{color_hex}' tidak valid. Gunakan format HEX 6 karakter, misal 'E0E0E0' atau '#FFFFFF'."

        def apply_shd(cell):
            tcPr = cell._element.get_or_add_tcPr()
            for s in tcPr.xpath("w:shd"):
                tcPr.remove(s)
            tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{clean_hex}"/>'))

        if row is not None and col is not None:
            if row < 0 or row >= len(table.rows):
                return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
            if col < 0 or col >= len(table.rows[row].cells):
                return f"Error: Indeks kolom {col} tidak valid. Total kolom pada baris {row}: {len(table.rows[row].cells)}."
            apply_shd(table.rows[row].cells[col])
            target_desc = f"sel [{row}, {col}]"
        elif row is not None:
            if row < 0 or row >= len(table.rows):
                return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
            for cell in table.rows[row].cells:
                apply_shd(cell)
            target_desc = f"seluruh baris {row}"
        elif col is not None:
            if col < 0 or col >= len(table.columns):
                return f"Error: Indeks kolom {col} tidak valid. Total kolom: {len(table.columns)}."
            for r in table.rows:
                if col < len(r.cells):
                    apply_shd(r.cells[col])
            target_desc = f"seluruh kolom {col}"
        else:
            for r in table.rows:
                for cell in r.cells:
                    apply_shd(cell)
            target_desc = "seluruh tabel"

        doc.save(str(path))
        return f"Berhasil menerapkan warna #{clean_hex} pada {target_desc} di tabel {table_index} pada {path}."
    except Exception as e:
        return f"Error: Gagal memberikan warna sel di {file_path}: {e}"


@mcp.tool()
def set_table_style_docx(
    file_path: str,
    table_index: int,
    style_name: str,
) -> str:
    """Mengubah template style bawaan Word untuk tabel (misalnya 'Table Grid', 'Light Shading', 'Normal Table')"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        table.style = style_name
        doc.save(str(path))
        return f"Berhasil mengubah style tabel {table_index} menjadi '{style_name}' di {path}."
    except Exception as e:
        return f"Error: Gagal mengubah style tabel di {file_path}: {e}"


@mcp.tool()
def set_table_alignment_docx(
    file_path: str,
    table_index: int,
    alignment: str,
) -> str:
    """Mengatur perataan posisi tabel di halaman dokumen ('left', 'center', atau 'right')"""
    try:
        align_map = {
            "left": WD_TABLE_ALIGNMENT.LEFT,
            "center": WD_TABLE_ALIGNMENT.CENTER,
            "right": WD_TABLE_ALIGNMENT.RIGHT,
        }
        align_key = alignment.strip().lower()
        if align_key not in align_map:
            return f"Error: Alignment '{alignment}' tidak valid. Pilihan: 'left', 'center', 'right'."

        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        table.alignment = align_map[align_key]
        doc.save(str(path))
        return f"Berhasil mengubah perataan tabel {table_index} menjadi '{align_key}' di {path}."
    except Exception as e:
        return f"Error: Gagal mengatur perataan tabel di {file_path}: {e}"


@mcp.tool()
def set_cell_alignment_docx(
    file_path: str,
    table_index: int,
    horizontal: str = "",
    vertical: str = "",
    row: int | None = None,
    col: int | None = None,
) -> str:
    """Mengatur perataan teks sel: horizontal ('left', 'center', 'right', 'justify') dan/atau vertikal ('top', 'center', 'bottom')"""
    try:
        h_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        v_map = {
            "top": WD_ALIGN_VERTICAL.TOP,
            "center": WD_ALIGN_VERTICAL.CENTER,
            "bottom": WD_ALIGN_VERTICAL.BOTTOM,
        }

        h_val = h_map.get(horizontal.strip().lower()) if horizontal else None
        v_val = v_map.get(vertical.strip().lower()) if vertical else None

        if horizontal and h_val is None:
            return f"Error: Perataan horizontal '{horizontal}' tidak valid. Pilihan: 'left', 'center', 'right', 'justify'."
        if vertical and v_val is None:
            return f"Error: Perataan vertikal '{vertical}' tidak valid. Pilihan: 'top', 'center', 'bottom'."
        if not horizontal and not vertical:
            return "Error: Berikan setidaknya salah satu parameter 'horizontal' atau 'vertical'."

        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]

        def apply_align(cell):
            if v_val is not None:
                cell.vertical_alignment = v_val
            if h_val is not None:
                for p in cell.paragraphs:
                    p.alignment = h_val

        if row is not None and col is not None:
            if row < 0 or row >= len(table.rows):
                return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
            if col < 0 or col >= len(table.rows[row].cells):
                return f"Error: Indeks kolom {col} tidak valid. Total kolom pada baris {row}: {len(table.rows[row].cells)}."
            apply_align(table.rows[row].cells[col])
            target_desc = f"sel [{row}, {col}]"
        elif row is not None:
            if row < 0 or row >= len(table.rows):
                return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
            for cell in table.rows[row].cells:
                apply_align(cell)
            target_desc = f"seluruh baris {row}"
        elif col is not None:
            if col < 0 or col >= len(table.columns):
                return f"Error: Indeks kolom {col} tidak valid. Total kolom: {len(table.columns)}."
            for r in table.rows:
                if col < len(r.cells):
                    apply_align(r.cells[col])
            target_desc = f"seluruh kolom {col}"
        else:
            for r in table.rows:
                for cell in r.cells:
                    apply_align(cell)
            target_desc = "seluruh tabel"

        doc.save(str(path))
        return f"Berhasil mengubah perataan pada {target_desc} di tabel {table_index} pada {path}."
    except Exception as e:
        return f"Error: Gagal mengatur perataan sel di {file_path}: {e}"


@mcp.tool()
def set_table_pagination_docx(
    file_path: str,
    table_index: int,
    repeat_header_rows: int = 1,
    cant_split: bool = True,
) -> str:
    """Mengatur opsi cetak/pagination tabel: mengulang header di setiap halaman (repeat_header_rows) dan mencegah baris terbelah (cant_split)"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]

        if repeat_header_rows > 0:
            for idx in range(min(repeat_header_rows, len(table.rows))):
                trPr = table.rows[idx]._tr.get_or_add_trPr()
                for h in trPr.xpath("w:tblHeader"):
                    trPr.remove(h)
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

        if cant_split:
            for row in table.rows:
                trPr = row._tr.get_or_add_trPr()
                for cs in trPr.xpath("w:cantSplit"):
                    trPr.remove(cs)
                trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

        doc.save(str(path))
        return f"Berhasil mengatur opsi pagination (repeat_header_rows={repeat_header_rows}, cant_split={cant_split}) pada tabel {table_index} di {path}."
    except Exception as e:
        return f"Error: Gagal mengatur pagination tabel di {file_path}: {e}"


@mcp.tool()
def set_cell_margins_docx(
    file_path: str,
    table_index: int,
    top_pt: float = 6.0,
    bottom_pt: float = 6.0,
    left_pt: float = 8.0,
    right_pt: float = 8.0,
    row: int | None = None,
    col: int | None = None,
) -> str:
    """Mengatur padding/margin internal sel (dalam satuan pt). Berlaku untuk sel spesifik ([row, col]) atau seluruh tabel jika row/col tidak diisi"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        top_dxa = int(top_pt * 20)
        bottom_dxa = int(bottom_pt * 20)
        left_dxa = int(left_pt * 20)
        right_dxa = int(right_pt * 20)

        if row is not None and col is not None:
            if row < 0 or row >= len(table.rows):
                return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
            if col < 0 or col >= len(table.rows[row].cells):
                return f"Error: Indeks kolom {col} tidak valid. Total kolom pada baris {row}: {len(table.rows[row].cells)}."
            cell = table.rows[row].cells[col]
            tcPr = cell._element.get_or_add_tcPr()
            for m in tcPr.xpath("w:tcMar"):
                tcPr.remove(m)
            tcMar = parse_xml(
                f'<w:tcMar {nsdecls("w")}>'
                f'<w:top w:w="{top_dxa}" w:type="dxa"/>'
                f'<w:bottom w:w="{bottom_dxa}" w:type="dxa"/>'
                f'<w:left w:w="{left_dxa}" w:type="dxa"/>'
                f'<w:right w:w="{right_dxa}" w:type="dxa"/>'
                f'</w:tcMar>'
            )
            tcPr.append(tcMar)
            target_desc = f"sel [{row}, {col}]"
        else:
            tblPr = table._element.xpath("w:tblPr")
            if not tblPr:
                return f"Error: Elemen properti tabel {table_index} tidak ditemukan."
            for m in tblPr[0].xpath("w:tblCellMar"):
                tblPr[0].remove(m)
            tblCellMar = parse_xml(
                f'<w:tblCellMar {nsdecls("w")}>'
                f'<w:top w:w="{top_dxa}" w:type="dxa"/>'
                f'<w:bottom w:w="{bottom_dxa}" w:type="dxa"/>'
                f'<w:left w:w="{left_dxa}" w:type="dxa"/>'
                f'<w:right w:w="{right_dxa}" w:type="dxa"/>'
                f'</w:tblCellMar>'
            )
            tblPr[0].append(tblCellMar)
            target_desc = "seluruh tabel"

        doc.save(str(path))
        return f"Berhasil mengatur margin padding ({top_pt}pt, {bottom_pt}pt, {left_pt}pt, {right_pt}pt) pada {target_desc} di tabel {table_index} pada {path}."
    except Exception as e:
        return f"Error: Gagal mengatur margin sel di {file_path}: {e}"


@mcp.tool()
def set_table_borders_docx(
    file_path: str,
    table_index: int,
    color_hex: str = "auto",
    size: int = 4,
    val: str = "single",
    top: bool = True,
    bottom: bool = True,
    left: bool = True,
    right: bool = True,
    inside_h: bool = True,
    inside_v: bool = True,
) -> str:
    """Mengatur border tabel secara kustom per sisi dengan opsi warna (HEX atau 'auto'), ketebalan (size), dan style border"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        tblPr = table._element.xpath("w:tblPr")
        if not tblPr:
            return f"Error: Elemen properti tabel {table_index} tidak ditemukan."

        for b in tblPr[0].xpath("w:tblBorders"):
            tblPr[0].remove(b)

        clean_color = color_hex.lstrip("#").upper() if color_hex.lower() != "auto" else "auto"
        val_top = val if top else "none"
        val_bottom = val if bottom else "none"
        val_left = val if left else "none"
        val_right = val if right else "none"
        val_inside_h = val if inside_h else "none"
        val_inside_v = val if inside_v else "none"

        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'<w:top w:val="{val_top}" w:sz="{size if top else 0}" w:space="0" w:color="{clean_color}"/>'
            f'<w:bottom w:val="{val_bottom}" w:sz="{size if bottom else 0}" w:space="0" w:color="{clean_color}"/>'
            f'<w:left w:val="{val_left}" w:sz="{size if left else 0}" w:space="0" w:color="{clean_color}"/>'
            f'<w:right w:val="{val_right}" w:sz="{size if right else 0}" w:space="0" w:color="{clean_color}"/>'
            f'<w:insideH w:val="{val_inside_h}" w:sz="{size if inside_h else 0}" w:space="0" w:color="{clean_color}"/>'
            f'<w:insideV w:val="{val_inside_v}" w:sz="{size if inside_v else 0}" w:space="0" w:color="{clean_color}"/>'
            f'</w:tblBorders>'
        )
        tblPr[0].append(borders)

        doc.save(str(path))
        return f"Berhasil mengatur border kustom (warna={clean_color}, size={size}) pada tabel {table_index} di {path}."
    except Exception as e:
        return f"Error: Gagal mengatur border tabel di {file_path}: {e}"


@mcp.tool()
def set_cell_text_direction_docx(
    file_path: str,
    table_index: int,
    direction: str = "bt_lr",
    row: int | None = None,
    col: int | None = None,
) -> str:
    """Mengatur arah orientasi teks sel ('horizontal', 'bt_lr' untuk rotasi 90 derajat ke atas, atau 'tb_rl' untuk rotasi 270 derajat ke bawah)"""
    try:
        dir_map = {
            "horizontal": "lrTb",
            "lr_tb": "lrTb",
            "bt_lr": "btLr",
            "vertical_up": "btLr",
            "tb_rl": "tbRl",
            "vertical_down": "tbRl",
        }
        clean_dir = direction.strip().lower()
        if clean_dir not in dir_map:
            return f"Error: Arah teks '{direction}' tidak valid. Pilihan: 'horizontal', 'bt_lr' (ke atas), 'tb_rl' (ke bawah)."

        xml_val = dir_map[clean_dir]

        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]

        def apply_dir(cell):
            tcPr = cell._element.get_or_add_tcPr()
            for td in tcPr.xpath("w:textDirection"):
                tcPr.remove(td)
            if xml_val != "lrTb":
                tcPr.append(parse_xml(f'<w:textDirection {nsdecls("w")} w:val="{xml_val}"/>'))

        if row is not None and col is not None:
            if row < 0 or row >= len(table.rows):
                return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
            if col < 0 or col >= len(table.rows[row].cells):
                return f"Error: Indeks kolom {col} tidak valid. Total kolom pada baris {row}: {len(table.rows[row].cells)}."
            apply_dir(table.rows[row].cells[col])
            target_desc = f"sel [{row}, {col}]"
        elif row is not None:
            if row < 0 or row >= len(table.rows):
                return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
            for cell in table.rows[row].cells:
                apply_dir(cell)
            target_desc = f"seluruh baris {row}"
        elif col is not None:
            if col < 0 or col >= len(table.columns):
                return f"Error: Indeks kolom {col} tidak valid. Total kolom: {len(table.columns)}."
            for r in table.rows:
                if col < len(r.cells):
                    apply_dir(r.cells[col])
            target_desc = f"seluruh kolom {col}"
        else:
            for r in table.rows:
                for cell in r.cells:
                    apply_dir(cell)
            target_desc = "seluruh tabel"

        doc.save(str(path))
        return f"Berhasil mengubah arah teks menjadi '{clean_dir}' pada {target_desc} di tabel {table_index} pada {path}."
    except Exception as e:
        return f"Error: Gagal mengubah arah teks sel di {file_path}: {e}"


@mcp.tool()
def split_table_docx(
    file_path: str,
    table_index: int,
    split_row_index: int,
) -> str:
    """Membagi (split) satu tabel menjadi dua tabel terpisah pada baris split_row_index (0-indexed)"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]

        if split_row_index <= 0 or split_row_index >= len(table.rows):
            return f"Error: split_row_index {split_row_index} tidak valid. Harus berada di antara 1 dan {len(table.rows)-1}."

        orig_tbl = table._element
        new_tbl = OxmlElement("w:tbl")

        tblPr = orig_tbl.xpath("w:tblPr")
        if tblPr:
            new_tbl.append(deepcopy(tblPr[0]))

        tblGrid = orig_tbl.xpath("w:tblGrid")
        if tblGrid:
            new_tbl.append(deepcopy(tblGrid[0]))

        trs_to_move = orig_tbl.xpath("w:tr")[split_row_index:]
        for tr in trs_to_move:
            orig_tbl.remove(tr)
            new_tbl.append(tr)

        sep_p = OxmlElement("w:p")
        orig_tbl.addnext(new_tbl)
        orig_tbl.addnext(sep_p)

        doc.save(str(path))
        return f"Berhasil membagi tabel {table_index} pada baris {split_row_index}. Dokumen kini memiliki {len(doc.tables)} tabel di {path}."
    except Exception as e:
        return f"Error: Gagal membagi tabel di {file_path}: {e}"


HIGHLIGHT_MAP = {
    "yellow": WD_COLOR_INDEX.YELLOW,
    "green": WD_COLOR_INDEX.GREEN,
    "bright_green": WD_COLOR_INDEX.BRIGHT_GREEN,
    "cyan": WD_COLOR_INDEX.TURQUOISE,
    "turquoise": WD_COLOR_INDEX.TURQUOISE,
    "pink": WD_COLOR_INDEX.PINK,
    "blue": WD_COLOR_INDEX.BLUE,
    "red": WD_COLOR_INDEX.RED,
    "dark_blue": WD_COLOR_INDEX.DARK_BLUE,
    "teal": WD_COLOR_INDEX.TEAL,
    "violet": WD_COLOR_INDEX.VIOLET,
    "dark_red": WD_COLOR_INDEX.DARK_RED,
    "dark_yellow": WD_COLOR_INDEX.DARK_YELLOW,
    "gray": WD_COLOR_INDEX.GRAY_50,
    "black": WD_COLOR_INDEX.BLACK,
}

UNDERLINE_MAP = {
    "single": WD_UNDERLINE.SINGLE,
    "double": WD_UNDERLINE.DOUBLE,
    "thick": WD_UNDERLINE.THICK,
    "dotted": WD_UNDERLINE.DOTTED,
    "dashed": WD_UNDERLINE.DASH,
    "dot_dash": WD_UNDERLINE.DOT_DASH,
    "dot_dot_dash": WD_UNDERLINE.DOT_DOT_DASH,
    "wave": WD_UNDERLINE.WAVY,
    "wavy": WD_UNDERLINE.WAVY,
    "none": False,
}


def parse_rgb_color(color_str: str) -> RGBColor:
    named_colors = {
        "black": (0, 0, 0),
        "white": (255, 255, 255),
        "red": (255, 0, 0),
        "green": (0, 128, 0),
        "blue": (0, 0, 255),
        "yellow": (255, 255, 0),
        "gray": (128, 128, 128),
        "grey": (128, 128, 128),
        "orange": (255, 165, 0),
        "purple": (128, 0, 128),
    }
    s = color_str.strip().lower()
    if s in named_colors:
        return RGBColor(*named_colors[s])
    if "," in s:
        parts = [int(x.strip()) for x in s.split(",")]
        if len(parts) == 3 and all(0 <= p <= 255 for p in parts):
            return RGBColor(*parts)
    clean_hex = s.lstrip("#")
    if len(clean_hex) == 6:
        return RGBColor(int(clean_hex[0:2], 16), int(clean_hex[2:4], 16), int(clean_hex[4:6], 16))
    raise ValueError(f"Warna '{color_str}' tidak valid. Gunakan HEX (misal '#FF0000'), RGB ('255, 0, 0'), atau nama warna.")


def apply_font_to_run(
    r,
    font_name: str | None = None,
    font_size_pt: float | None = None,
    color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strike: bool | None = None,
    highlight_color: str | None = None,
    subscript: bool | None = None,
    superscript: bool | None = None,
    small_caps: bool | None = None,
    all_caps: bool | None = None,
    underline_style: str | None = None,
):
    if font_name:
        r.font.name = font_name
        r._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:hAnsi"), font_name)
    if font_size_pt is not None and font_size_pt > 0:
        r.font.size = Pt(font_size_pt)
    if color:
        r.font.color.rgb = parse_rgb_color(color)
    if bold is not None:
        r.font.bold = bold
    if italic is not None:
        r.font.italic = italic
    if underline is not None:
        r.font.underline = underline
    if strike is not None:
        r.font.strike = strike
    if subscript is not None:
        r.font.subscript = subscript
    if superscript is not None:
        r.font.superscript = superscript
    if small_caps is not None:
        r.font.small_caps = small_caps
    if all_caps is not None:
        r.font.all_caps = all_caps
    if underline_style:
        clean_u = underline_style.strip().lower()
        if clean_u in UNDERLINE_MAP:
            r.font.underline = UNDERLINE_MAP[clean_u]
    if highlight_color:
        clean_hl = highlight_color.strip().lower()
        if clean_hl in HIGHLIGHT_MAP:
            r.font.highlight_color = HIGHLIGHT_MAP[clean_hl]


def apply_style_to_paragraph(
    p,
    target_text: str | None = None,
    font_name: str | None = None,
    font_size_pt: float | None = None,
    color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strike: bool | None = None,
    highlight_color: str | None = None,
    subscript: bool | None = None,
    superscript: bool | None = None,
    small_caps: bool | None = None,
    all_caps: bool | None = None,
    underline_style: str | None = None,
) -> int:
    applied = 0
    font_kwargs = {
        "font_name": font_name,
        "font_size_pt": font_size_pt,
        "color": color,
        "bold": bold,
        "italic": italic,
        "underline": underline,
        "strike": strike,
        "highlight_color": highlight_color,
        "subscript": subscript,
        "superscript": superscript,
        "small_caps": small_caps,
        "all_caps": all_caps,
        "underline_style": underline_style,
    }
    if target_text and target_text in p.text:
        orig_text = p.text
        parts = orig_text.split(target_text)
        p.text = ""
        for i, part in enumerate(parts):
            if part:
                p.add_run(part)
            if i < len(parts) - 1:
                r = p.add_run(target_text)
                apply_font_to_run(r, **font_kwargs)
                applied += 1
    elif not target_text:
        if not p.runs and p.text:
            text = p.text
            p.text = ""
            r = p.add_run(text)
            apply_font_to_run(r, **font_kwargs)
            applied += 1
        else:
            for r in p.runs:
                apply_font_to_run(r, **font_kwargs)
                applied += 1
    return applied


@mcp.tool()
def format_text_docx(
    file_path: str,
    target_text: str | None = None,
    paragraph_index: int | None = None,
    table_index: int | None = None,
    row: int | None = None,
    col: int | None = None,
    font_name: str | None = None,
    font_size_pt: float | None = None,
    color: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strike: bool | None = None,
    highlight_color: str | None = None,
    subscript: bool | None = None,
    superscript: bool | None = None,
    small_caps: bool | None = None,
    all_caps: bool | None = None,
    underline_style: str | None = None,
) -> str:
    """Mengatur jenis font, ukuran (Pt), warna teks, bold, italic, underline, strikethrough, subscript, superscript, small_caps, all_caps, dan highlight"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))
        applied_total = 0

        target_paragraphs = []
        scope_desc = ""

        if table_index is not None:
            if table_index < 0 or table_index >= len(doc.tables):
                return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."
            table = doc.tables[table_index]
            if row is not None and col is not None:
                if row < 0 or row >= len(table.rows):
                    return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
                if col < 0 or col >= len(table.rows[row].cells):
                    return f"Error: Indeks kolom {col} tidak valid. Total kolom pada baris {row}: {len(table.rows[row].cells)}."
                target_paragraphs.extend(table.rows[row].cells[col].paragraphs)
                scope_desc = f"tabel {table_index} sel [{row}, {col}]"
            elif row is not None:
                if row < 0 or row >= len(table.rows):
                    return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
                for cell in table.rows[row].cells:
                    target_paragraphs.extend(cell.paragraphs)
                scope_desc = f"tabel {table_index} baris {row}"
            elif col is not None:
                if col < 0 or col >= len(table.columns):
                    return f"Error: Indeks kolom {col} tidak valid. Total kolom: {len(table.columns)}."
                for r in table.rows:
                    if col < len(r.cells):
                        target_paragraphs.extend(r.cells[col].paragraphs)
                scope_desc = f"tabel {table_index} kolom {col}"
            else:
                for r in table.rows:
                    for cell in r.cells:
                        target_paragraphs.extend(cell.paragraphs)
                scope_desc = f"seluruh tabel {table_index}"
        elif paragraph_index is not None:
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return f"Error: Indeks paragraf {paragraph_index} tidak ditemukan. Total paragraf: {len(doc.paragraphs)}."
            target_paragraphs.append(doc.paragraphs[paragraph_index])
            scope_desc = f"paragraf {paragraph_index}"
        else:
            target_paragraphs.extend(doc.paragraphs)
            for t in doc.tables:
                for r in t.rows:
                    for c in r.cells:
                        target_paragraphs.extend(c.paragraphs)
            scope_desc = "seluruh dokumen"

        for p in target_paragraphs:
            applied_total += apply_style_to_paragraph(
                p,
                target_text=target_text,
                font_name=font_name,
                font_size_pt=font_size_pt,
                color=color,
                bold=bold,
                italic=italic,
                underline=underline,
                strike=strike,
                highlight_color=highlight_color,
                subscript=subscript,
                superscript=superscript,
                small_caps=small_caps,
                all_caps=all_caps,
                underline_style=underline_style,
            )

        doc.save(str(path))
        styles_summary = []
        if font_name:
            styles_summary.append(f"font='{font_name}'")
        if font_size_pt:
            styles_summary.append(f"size={font_size_pt}pt")
        if color:
            styles_summary.append(f"color='{color}'")
        if bold is not None:
            styles_summary.append(f"bold={bold}")
        if italic is not None:
            styles_summary.append(f"italic={italic}")
        if underline is not None:
            styles_summary.append(f"underline={underline}")
        if underline_style:
            styles_summary.append(f"underline_style='{underline_style}'")
        if strike is not None:
            styles_summary.append(f"strike={strike}")
        if subscript is not None:
            styles_summary.append(f"subscript={subscript}")
        if superscript is not None:
            styles_summary.append(f"superscript={superscript}")
        if small_caps is not None:
            styles_summary.append(f"small_caps={small_caps}")
        if all_caps is not None:
            styles_summary.append(f"all_caps={all_caps}")
        if highlight_color:
            styles_summary.append(f"highlight='{highlight_color}'")

        style_str = ", ".join(styles_summary) if styles_summary else "default style"
        match_str = f" untuk teks '{target_text}'" if target_text else ""
        return f"Berhasil memformat teks pada {scope_desc}{match_str} ({style_str}) di {path}."
    except Exception as e:
        return f"Error: Gagal memformat teks pada {file_path}: {e}"


@mcp.tool()
def format_paragraph_docx(
    file_path: str,
    paragraph_index: int | None = None,
    table_index: int | None = None,
    row: int | None = None,
    col: int | None = None,
    line_spacing: float | None = None,
    line_spacing_pt: float | None = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    first_line_indent_inches: float | None = None,
    left_indent_inches: float | None = None,
    right_indent_inches: float | None = None,
    hanging_indent_inches: float | None = None,
    alignment: str = "",
    list_style: str = "",
    keep_with_next: bool | None = None,
    page_break_before: bool | None = None,
    widow_control: bool | None = None,
) -> str:
    """Mengatur format paragraf: spasi baris (misal 1.5, 2.0 atau Pt), spasi sebelum/sesudah (pt), indentasi/hanging indent (inci), perataan, list, dan kontrol halaman (keep_with_next, page_break_before)"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        target_paragraphs = []
        scope_desc = ""

        if table_index is not None:
            if table_index < 0 or table_index >= len(doc.tables):
                return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."
            table = doc.tables[table_index]
            if row is not None and col is not None:
                if row < 0 or row >= len(table.rows):
                    return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
                if col < 0 or col >= len(table.rows[row].cells):
                    return f"Error: Indeks kolom {col} tidak valid. Total kolom pada baris {row}: {len(table.rows[row].cells)}."
                target_paragraphs.extend(table.rows[row].cells[col].paragraphs)
                scope_desc = f"tabel {table_index} sel [{row}, {col}]"
            elif row is not None:
                if row < 0 or row >= len(table.rows):
                    return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
                for cell in table.rows[row].cells:
                    target_paragraphs.extend(cell.paragraphs)
                scope_desc = f"tabel {table_index} baris {row}"
            elif col is not None:
                if col < 0 or col >= len(table.columns):
                    return f"Error: Indeks kolom {col} tidak valid. Total kolom: {len(table.columns)}."
                for r_item in table.rows:
                    if col < len(r_item.cells):
                        target_paragraphs.extend(r_item.cells[col].paragraphs)
                scope_desc = f"tabel {table_index} kolom {col}"
            else:
                for r_item in table.rows:
                    for cell in r_item.cells:
                        target_paragraphs.extend(cell.paragraphs)
                scope_desc = f"seluruh tabel {table_index}"
        elif paragraph_index is not None:
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return f"Error: Indeks paragraf {paragraph_index} tidak ditemukan. Total paragraf: {len(doc.paragraphs)}."
            target_paragraphs.append(doc.paragraphs[paragraph_index])
            scope_desc = f"paragraf {paragraph_index}"
        else:
            target_paragraphs.extend(doc.paragraphs)
            scope_desc = "seluruh paragraf dalam dokumen"

        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        align_val = align_map.get(alignment.strip().lower()) if alignment else None

        list_map = {
            "bullet": "List Bullet",
            "list_bullet": "List Bullet",
            "number": "List Number",
            "list_number": "List Number",
        }
        style_val = list_map.get(list_style.strip().lower(), list_style) if list_style else None

        for p in target_paragraphs:
            if style_val:
                try:
                    p.style = style_val
                except Exception:
                    pass
            if align_val is not None:
                p.alignment = align_val
            if line_spacing_pt is not None:
                p.paragraph_format.line_spacing = Pt(line_spacing_pt)
            elif line_spacing is not None and line_spacing > 0:
                p.paragraph_format.line_spacing = line_spacing
            if space_before_pt is not None:
                p.paragraph_format.space_before = Pt(space_before_pt)
            if space_after_pt is not None:
                p.paragraph_format.space_after = Pt(space_after_pt)
            if hanging_indent_inches is not None:
                p.paragraph_format.left_indent = Inches(hanging_indent_inches)
                p.paragraph_format.first_line_indent = Inches(-hanging_indent_inches)
            if first_line_indent_inches is not None:
                p.paragraph_format.first_line_indent = Inches(first_line_indent_inches)
            if left_indent_inches is not None:
                p.paragraph_format.left_indent = Inches(left_indent_inches)
            if right_indent_inches is not None:
                p.paragraph_format.right_indent = Inches(right_indent_inches)
            if keep_with_next is not None:
                p.paragraph_format.keep_with_next = keep_with_next
            if page_break_before is not None:
                p.paragraph_format.page_break_before = page_break_before
            if widow_control is not None:
                p.paragraph_format.widow_control = widow_control

        doc.save(str(path))
        return f"Berhasil memformat tata letak dan spasi pada {scope_desc} di {path}."
    except Exception as e:
        return f"Error: Gagal memformat paragraf di {file_path}: {e}"


@mcp.tool()
def add_styled_paragraph_docx(
    file_path: str,
    text: str,
    heading_level: int = 0,
    font_name: str | None = None,
    font_size_pt: float | None = None,
    color: str | None = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    strike: bool = False,
    highlight_color: str | None = None,
    alignment: str = "",
    line_spacing: float | None = None,
    line_spacing_pt: float | None = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    first_line_indent_inches: float | None = None,
    left_indent_inches: float | None = None,
    right_indent_inches: float | None = None,
    hanging_indent_inches: float | None = None,
    keep_with_next: bool | None = None,
    page_break_before: bool | None = None,
    widow_control: bool | None = None,
    list_style: str = "",
    subscript: bool = False,
    superscript: bool = False,
    small_caps: bool = False,
    all_caps: bool = False,
    underline_style: str | None = None,
) -> str:
    """Menambahkan paragraf atau heading baru dengan opsi lengkap font, ukuran, warna, gaya teks, spasi, dan perataan ke file .docx"""
    try:
        path = validate_docx_path(file_path, must_exist=False)
        doc = Document(str(path)) if path.is_file() else Document()

        if heading_level > 0:
            p = doc.add_heading(level=min(heading_level, 9))
        else:
            p = doc.add_paragraph()

        if list_style:
            l_map = {"bullet": "List Bullet", "list_bullet": "List Bullet", "number": "List Number", "list_number": "List Number"}
            try:
                p.style = l_map.get(list_style.strip().lower(), list_style)
            except Exception:
                pass

        r = p.add_run(text)
        apply_font_to_run(
            r,
            font_name=font_name,
            font_size_pt=font_size_pt,
            color=color,
            bold=bold,
            italic=italic,
            underline=underline,
            strike=strike,
            highlight_color=highlight_color,
            subscript=subscript,
            superscript=superscript,
            small_caps=small_caps,
            all_caps=all_caps,
            underline_style=underline_style,
        )

        if alignment:
            align_map = {
                "left": WD_ALIGN_PARAGRAPH.LEFT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            }
            clean_align = alignment.strip().lower()
            if clean_align in align_map:
                p.alignment = align_map[clean_align]

        if line_spacing_pt is not None:
            p.paragraph_format.line_spacing = Pt(line_spacing_pt)
        elif line_spacing is not None and line_spacing > 0:
            p.paragraph_format.line_spacing = line_spacing
        if space_before_pt is not None:
            p.paragraph_format.space_before = Pt(space_before_pt)
        if space_after_pt is not None:
            p.paragraph_format.space_after = Pt(space_after_pt)
        if hanging_indent_inches is not None:
            p.paragraph_format.left_indent = Inches(hanging_indent_inches)
            p.paragraph_format.first_line_indent = Inches(-hanging_indent_inches)
        if first_line_indent_inches is not None:
            p.paragraph_format.first_line_indent = Inches(first_line_indent_inches)
        if left_indent_inches is not None:
            p.paragraph_format.left_indent = Inches(left_indent_inches)
        if right_indent_inches is not None:
            p.paragraph_format.right_indent = Inches(right_indent_inches)
        if keep_with_next is not None:
            p.paragraph_format.keep_with_next = keep_with_next
        if page_break_before is not None:
            p.paragraph_format.page_break_before = page_break_before
        if widow_control is not None:
            p.paragraph_format.widow_control = widow_control

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        return f"Berhasil menambahkan paragraf berformat ke {path}."
    except Exception as e:
        return f"Error: Gagal menambahkan paragraf berformat di {file_path}: {e}"


@mcp.tool()
def add_picture_docx(
    file_path: str,
    image_path: str = "",
    image_base64: str = "",
    width_inches: float | None = None,
    height_inches: float | None = None,
    width_cm: float | None = None,
    height_cm: float | None = None,
    alignment: str = "",
    caption: str = "",
    alt_text: str = "",
    title: str = "",
    paragraph_index: int | None = None,
    table_index: int | None = None,
    row: int | None = None,
    col: int | None = None,
) -> str:
    """Menyisipkan gambar (.png/.jpg/dll) ke dokumen Word melalui path file atau string base64 dengan opsi lebar/tinggi (Inches atau Cm), perataan, caption, alt text/title, atau ke dalam sel tabel"""
    try:
        path = validate_docx_path(file_path, must_exist=False)
        if not image_path and not image_base64:
            return "Error: Salah satu dari image_path atau image_base64 harus disertakan."

        if image_base64:
            raw_b64 = image_base64.strip()
            if "," in raw_b64 and "base64" in raw_b64:
                raw_b64 = raw_b64.split(",", 1)[1]
            try:
                img_bytes = base64.b64decode(raw_b64)
            except Exception as b64_err:
                return f"Error: Gagal mendecode image_base64: {b64_err}"
            image_source = io.BytesIO(img_bytes)
            img_desc_name = "gambar base64"
        else:
            img_path = validate_image_path(image_path)
            image_source = str(img_path)
            img_desc_name = f"gambar '{img_path.name}'"

        doc = Document(str(path)) if path.is_file() else Document()

        width = None
        if width_inches is not None and width_inches > 0:
            width = Inches(width_inches)
        elif width_cm is not None and width_cm > 0:
            width = Cm(width_cm)

        height = None
        if height_inches is not None and height_inches > 0:
            height = Inches(height_inches)
        elif height_cm is not None and height_cm > 0:
            height = Cm(height_cm)

        cell = None
        if table_index is not None:
            if table_index < 0 or table_index >= len(doc.tables):
                return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."
            table = doc.tables[table_index]
            if row is None or col is None:
                return "Error: Parameter row dan col wajib diisi saat table_index ditentukan."
            if row < 0 or row >= len(table.rows):
                return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
            if col < 0 or col >= len(table.rows[row].cells):
                return f"Error: Indeks kolom {col} tidak valid. Total kolom pada baris {row}: {len(table.rows[row].cells)}."
            cell = table.rows[row].cells[col]
            p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
            target_desc = f"tabel {table_index} sel [{row}, {col}]"
        elif paragraph_index is not None:
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return f"Error: Indeks paragraf {paragraph_index} tidak ditemukan. Total paragraf: {len(doc.paragraphs)}."
            p = doc.paragraphs[paragraph_index]
            target_desc = f"paragraf {paragraph_index}"
        else:
            p = doc.add_paragraph()
            target_desc = "dokumen"

        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        if alignment:
            clean_align = alignment.strip().lower()
            if clean_align in align_map:
                p.alignment = align_map[clean_align]

        r = p.add_run()
        kwargs = {}
        if width is not None:
            kwargs["width"] = width
        if height is not None:
            kwargs["height"] = height
        inline_shape = r.add_picture(image_source, **kwargs)

        if alt_text or title:
            try:
                docPr = inline_shape._inline.docPr
                if alt_text:
                    docPr.set("descr", alt_text)
                if title:
                    docPr.set("title", title)
            except Exception:
                pass

        if caption:
            cap_p = cell.add_paragraph() if cell is not None else doc.add_paragraph()
            if alignment:
                clean_align = alignment.strip().lower()
                if clean_align in align_map:
                    cap_p.alignment = align_map[clean_align]
            cap_r = cap_p.add_run(caption)
            cap_r.font.italic = True
            cap_r.font.size = Pt(9.5)
            cap_p.paragraph_format.space_before = Pt(3.0)
            cap_p.paragraph_format.space_after = Pt(6.0)

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))

        dim_desc = []
        if width_inches:
            dim_desc.append(f"lebar={width_inches}in")
        elif width_cm:
            dim_desc.append(f"lebar={width_cm}cm")
        if height_inches:
            dim_desc.append(f"tinggi={height_inches}in")
        elif height_cm:
            dim_desc.append(f"tinggi={height_cm}cm")
        dim_str = f" ({', '.join(dim_desc)})" if dim_desc else " (ukuran asli)"

        return f"Berhasil menyisipkan {img_desc_name} ke {target_desc}{dim_str} di {path}."
    except Exception as e:
        return f"Error: Gagal menyisipkan gambar ke {file_path}: {e}"


@mcp.tool()
def get_images_info_docx(file_path: str) -> str:
    """Mendapatkan daftar dan informasi dimensi seluruh gambar yang ada di dalam dokumen .docx"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))
        shapes = doc.inline_shapes
        if not shapes:
            return f"Tidak ada gambar ditemukan pada dokumen {path}."

        info_lines = [f"Total gambar pada {path.name}: {len(shapes)}"]
        for idx, shape in enumerate(shapes):
            w_in = shape.width / 914400.0
            h_in = shape.height / 914400.0
            w_cm = shape.width / 360000.0
            h_cm = shape.height / 360000.0
            info_lines.append(
                f"[Gambar {idx}] Lebar: {w_in:.2f} in ({w_cm:.2f} cm) | Tinggi: {h_in:.2f} in ({h_cm:.2f} cm)"
            )
        return "\n".join(info_lines)
    except Exception as e:
        return f"Error: Gagal membaca informasi gambar dari {file_path}: {e}"


@mcp.tool()
def resize_picture_docx(
    file_path: str,
    picture_index: int = 0,
    width_inches: float | None = None,
    height_inches: float | None = None,
    width_cm: float | None = None,
    height_cm: float | None = None,
) -> str:
    """Mengubah ukuran (resize) gambar yang sudah ada di dokumen berdasarkan indeks gambar (Inches atau Cm)"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))
        shapes = doc.inline_shapes
        if not shapes:
            return f"Error: Dokumen {path} tidak memiliki gambar."
        if picture_index < 0 or picture_index >= len(shapes):
            return f"Error: Indeks gambar {picture_index} tidak ditemukan. Total gambar: {len(shapes)}."

        shape = shapes[picture_index]

        if width_inches is not None and width_inches > 0:
            shape.width = Inches(width_inches)
        elif width_cm is not None and width_cm > 0:
            shape.width = Cm(width_cm)

        if height_inches is not None and height_inches > 0:
            shape.height = Inches(height_inches)
        elif height_cm is not None and height_cm > 0:
            shape.height = Cm(height_cm)

        doc.save(str(path))
        new_w_in = shape.width / 914400.0
        new_h_in = shape.height / 914400.0
        return f"Berhasil mengubah ukuran gambar {picture_index} di {path} menjadi lebar={new_w_in:.2f}in, tinggi={new_h_in:.2f}in."
    except Exception as e:
        return f"Error: Gagal mengubah ukuran gambar di {file_path}: {e}"


@mcp.tool()
def delete_picture_docx(file_path: str, picture_index: int = 0) -> str:
    """Menghapus gambar dari dokumen .docx berdasarkan indeks gambar"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))
        shapes = doc.inline_shapes
        if not shapes:
            return f"Error: Dokumen {path} tidak memiliki gambar."
        if picture_index < 0 or picture_index >= len(shapes):
            return f"Error: Indeks gambar {picture_index} tidak ditemukan. Total gambar: {len(shapes)}."

        shape = shapes[picture_index]
        inline_elem = shape._inline
        parent = inline_elem.getparent()
        if parent is not None:
            parent.remove(inline_elem)
        else:
            return f"Error: Gagal menemukan elemen induk untuk gambar {picture_index}."

        doc.save(str(path))
        return f"Berhasil menghapus gambar {picture_index} dari {path}."
    except Exception as e:
        return f"Error: Gagal menghapus gambar dari {file_path}: {e}"


@mcp.tool()
def extract_images_docx(file_path: str, output_dir: str) -> str:
    """Mengekstrak seluruh file gambar yang tertanam di dalam dokumen .docx dan menyimpannya ke folder output"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        out_path = Path(output_dir).expanduser().resolve()
        allowed_dir = os.environ.get("DOCX_MCP_ALLOWED_DIR")
        if allowed_dir:
            allowed_base = Path(allowed_dir).expanduser().resolve()
            if not out_path.is_relative_to(allowed_base):
                raise PermissionError(
                    f"Akses ditolak: Direktori output berada di luar direktori yang diizinkan ({allowed_base})."
                )

        out_path.mkdir(parents=True, exist_ok=True)
        extracted_files = []

        for rel in doc.part.rels.values():
            if "image" in rel.reltype:
                target_part = rel.target_part
                filename = Path(target_part.partname).name
                dest_file = out_path / filename
                dest_file.write_bytes(target_part.blob)
                extracted_files.append(filename)

        if not extracted_files:
            return f"Tidak ada gambar ditemukan di dalam dokumen {path}."

        return f"Berhasil mengekstrak {len(extracted_files)} gambar ke {out_path}: {', '.join(extracted_files)}"
    except Exception as e:
        return f"Error: Gagal mengekstrak gambar dari {file_path}: {e}"


PAPER_SIZES = {
    "a4": (21.0, 29.7),
    "a3": (29.7, 42.0),
    "a5": (14.8, 21.0),
    "letter": (21.59, 27.94),
    "legal": (21.59, 35.56),
    "folio": (21.5, 33.0),
    "f4": (21.5, 33.0),
    "executive": (18.415, 26.67),
}

MARGIN_PRESETS = {
    "normal": (2.54, 2.54, 2.54, 2.54),
    "narrow": (1.27, 1.27, 1.27, 1.27),
    "moderate": (2.54, 2.54, 1.905, 1.905),
    "wide": (2.54, 2.54, 5.08, 5.08),
    "skripsi": (4.0, 3.0, 4.0, 3.0),
}


def apply_section_formatting(
    sec,
    paper_size: str = "",
    page_width_inches: float | None = None,
    page_height_inches: float | None = None,
    page_width_cm: float | None = None,
    page_height_cm: float | None = None,
    orientation: str = "",
    margin_preset: str = "",
    top_margin_inches: float | None = None,
    bottom_margin_inches: float | None = None,
    left_margin_inches: float | None = None,
    right_margin_inches: float | None = None,
    top_margin_cm: float | None = None,
    bottom_margin_cm: float | None = None,
    left_margin_cm: float | None = None,
    right_margin_cm: float | None = None,
    gutter_inches: float | None = None,
    gutter_cm: float | None = None,
    header_distance_inches: float | None = None,
    header_distance_cm: float | None = None,
    footer_distance_inches: float | None = None,
    footer_distance_cm: float | None = None,
    different_first_page_header_footer: bool | None = None,
    columns_count: int | None = None,
    column_spacing_inches: float | None = None,
    column_spacing_cm: float | None = None,
    page_number_start: int | None = None,
    page_number_format: str = "",
):
    target_orient = orientation.strip().lower() if orientation else None

    if paper_size:
        clean_ps = paper_size.strip().lower()
        if clean_ps in PAPER_SIZES:
            pw, ph = PAPER_SIZES[clean_ps]
            is_land = (target_orient == "landscape") or (target_orient is None and sec.orientation == WD_ORIENTATION.LANDSCAPE)
            if is_land:
                sec.page_width = Cm(max(pw, ph))
                sec.page_height = Cm(min(pw, ph))
                sec.orientation = WD_ORIENTATION.LANDSCAPE
            else:
                sec.page_width = Cm(min(pw, ph))
                sec.page_height = Cm(max(pw, ph))
                sec.orientation = WD_ORIENTATION.PORTRAIT
    elif page_width_inches is not None and page_height_inches is not None:
        sec.page_width = Inches(page_width_inches)
        sec.page_height = Inches(page_height_inches)
    elif page_width_cm is not None and page_height_cm is not None:
        sec.page_width = Cm(page_width_cm)
        sec.page_height = Cm(page_height_cm)

    if target_orient and not paper_size:
        if target_orient == "landscape":
            sec.orientation = WD_ORIENTATION.LANDSCAPE
            if sec.page_width < sec.page_height:
                sec.page_width, sec.page_height = sec.page_height, sec.page_width
        elif target_orient == "portrait":
            sec.orientation = WD_ORIENTATION.PORTRAIT
            if sec.page_width > sec.page_height:
                sec.page_width, sec.page_height = sec.page_height, sec.page_width

    if margin_preset:
        clean_mp = margin_preset.strip().lower()
        if clean_mp in MARGIN_PRESETS:
            top_c, bot_c, left_c, right_c = MARGIN_PRESETS[clean_mp]
            sec.top_margin = Cm(top_c)
            sec.bottom_margin = Cm(bot_c)
            sec.left_margin = Cm(left_c)
            sec.right_margin = Cm(right_c)

    if top_margin_inches is not None:
        sec.top_margin = Inches(top_margin_inches)
    elif top_margin_cm is not None:
        sec.top_margin = Cm(top_margin_cm)

    if bottom_margin_inches is not None:
        sec.bottom_margin = Inches(bottom_margin_inches)
    elif bottom_margin_cm is not None:
        sec.bottom_margin = Cm(bottom_margin_cm)

    if left_margin_inches is not None:
        sec.left_margin = Inches(left_margin_inches)
    elif left_margin_cm is not None:
        sec.left_margin = Cm(left_margin_cm)

    if right_margin_inches is not None:
        sec.right_margin = Inches(right_margin_inches)
    elif right_margin_cm is not None:
        sec.right_margin = Cm(right_margin_cm)

    if gutter_inches is not None:
        sec.gutter = Inches(gutter_inches)
    elif gutter_cm is not None:
        sec.gutter = Cm(gutter_cm)

    if header_distance_inches is not None:
        sec.header_distance = Inches(header_distance_inches)
    elif header_distance_cm is not None:
        sec.header_distance = Cm(header_distance_cm)

    if footer_distance_inches is not None:
        sec.footer_distance = Inches(footer_distance_inches)
    elif footer_distance_cm is not None:
        sec.footer_distance = Cm(footer_distance_cm)

    if different_first_page_header_footer is not None:
        sec.different_first_page_header_footer = different_first_page_header_footer

    if columns_count is not None and columns_count >= 1:
        sectPr = sec._sectPr
        if column_spacing_inches is not None:
            space_twips = int(column_spacing_inches * 1440)
        elif column_spacing_cm is not None:
            space_twips = int(column_spacing_cm * 567)
        else:
            space_twips = 720

        cols_elm = sectPr.find(qn("w:cols"))
        if cols_elm is None:
            cols_elm = parse_xml(f'<w:cols {nsdecls("w")} w:space="{space_twips}"/>')
            sectPr.append(cols_elm)
        cols_elm.set(qn("w:num"), str(columns_count))
        cols_elm.set(qn("w:space"), str(space_twips))

    if page_number_start is not None or page_number_format:
        sectPr = sec._sectPr
        pg_elm = sectPr.find(qn("w:pgNumType"))
        if pg_elm is None:
            pg_elm = OxmlElement("w:pgNumType")
            sectPr.append(pg_elm)
        if page_number_start is not None:
            pg_elm.set(qn("w:start"), str(page_number_start))
        if page_number_format:
            pg_elm.set(qn("w:fmt"), page_number_format.strip())


@mcp.tool()
def get_page_setup_docx(file_path: str) -> str:
    """Membaca informasi ukuran kertas, orientasi, margin (cm/inci), gutter, header/footer distance, kolom, dan penomoran halaman untuk seluruh section dokumen"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))
        lines = [f"Pengaturan Halaman Dokumen {path.name} (Total {len(doc.sections)} section):"]

        mirror_elm = doc.settings.element.find(qn("w:mirrorMargins"))
        mirror_status = "Aktif (Facing Pages)" if mirror_elm is not None else "Tidak Aktif"
        lines.append(f"- Status Margin Cermin (Mirror Margins): {mirror_status}")

        for idx, sec in enumerate(doc.sections):
            w_cm = sec.page_width / 360000.0
            h_cm = sec.page_height / 360000.0
            w_in = sec.page_width / 914400.0
            h_in = sec.page_height / 914400.0
            orient = "Landscape" if sec.orientation == WD_ORIENTATION.LANDSCAPE else "Portrait"

            detected_size = "Kustom"
            for name, (cw, ch) in PAPER_SIZES.items():
                min_dim = min(cw, ch)
                max_dim = max(cw, ch)
                cur_min = min(w_cm, h_cm)
                cur_max = max(w_cm, h_cm)
                if abs(min_dim - cur_min) < 0.1 and abs(max_dim - cur_max) < 0.1:
                    detected_size = name.upper()
                    break

            top_cm = sec.top_margin / 360000.0
            bot_cm = sec.bottom_margin / 360000.0
            left_cm = sec.left_margin / 360000.0
            right_cm = sec.right_margin / 360000.0

            lines.append(f"\n[Section {idx}]")
            lines.append(f"- Ukuran Kertas: {detected_size} ({w_cm:.2f} x {h_cm:.2f} cm / {w_in:.2f} x {h_in:.2f} in)")
            lines.append(f"- Orientasi: {orient}")
            lines.append(f"- Margin: Atas={top_cm:.2f}cm, Bawah={bot_cm:.2f}cm, Kiri={left_cm:.2f}cm, Kanan={right_cm:.2f}cm")
            if sec.gutter:
                lines.append(f"- Gutter: {sec.gutter / 360000.0:.2f} cm")
            lines.append(f"- Beda Halaman Pertama (Cover): {'Ya' if sec.different_first_page_header_footer else 'Tidak'}")

            cols_elm = sec._sectPr.find(qn("w:cols"))
            if cols_elm is not None and cols_elm.get(qn("w:num")):
                num_c = cols_elm.get(qn("w:num"))
                sp_twips = int(cols_elm.get(qn("w:space"), "720"))
                lines.append(f"- Jumlah Kolom: {num_c} (Jeda={sp_twips / 567.0:.2f} cm)")

            pg_elm = sec._sectPr.find(qn("w:pgNumType"))
            if pg_elm is not None:
                pg_start = pg_elm.get(qn("w:start"), "otomatis")
                pg_fmt = pg_elm.get(qn("w:fmt"), "default")
                lines.append(f"- Penomoran Halaman: Mulai={pg_start}, Format={pg_fmt}")

        return "\n".join(lines)
    except Exception as e:
        return f"Error: Gagal membaca pengaturan halaman di {file_path}: {e}"


@mcp.tool()
def set_page_setup_docx(
    file_path: str,
    section_index: int | None = None,
    paper_size: str = "",
    page_width_inches: float | None = None,
    page_height_inches: float | None = None,
    page_width_cm: float | None = None,
    page_height_cm: float | None = None,
    orientation: str = "",
    margin_preset: str = "",
    top_margin_inches: float | None = None,
    bottom_margin_inches: float | None = None,
    left_margin_inches: float | None = None,
    right_margin_inches: float | None = None,
    top_margin_cm: float | None = None,
    bottom_margin_cm: float | None = None,
    left_margin_cm: float | None = None,
    right_margin_cm: float | None = None,
    gutter_inches: float | None = None,
    gutter_cm: float | None = None,
    header_distance_inches: float | None = None,
    header_distance_cm: float | None = None,
    footer_distance_inches: float | None = None,
    footer_distance_cm: float | None = None,
    different_first_page_header_footer: bool | None = None,
    columns_count: int | None = None,
    column_spacing_inches: float | None = None,
    column_spacing_cm: float | None = None,
    page_number_start: int | None = None,
    page_number_format: str = "",
    mirror_margins: bool | None = None,
) -> str:
    """Mengatur ukuran kertas (A4, Letter, F4/Folio, Legal, dll), orientasi, margin (cm/inci/preset: normal, narrow, skripsi), gutter, multi-kolom koran, restart nomor halaman, dan mirror margins"""
    try:
        path = validate_docx_path(file_path, must_exist=False)
        doc = Document(str(path)) if path.is_file() else Document()

        if mirror_margins is not None:
            settings_elm = doc.settings.element
            existing = settings_elm.find(qn("w:mirrorMargins"))
            if mirror_margins:
                if existing is None:
                    settings_elm.append(OxmlElement("w:mirrorMargins"))
            else:
                if existing is not None:
                    settings_elm.remove(existing)

        target_sections = []
        if section_index is not None:
            if section_index < 0 or section_index >= len(doc.sections):
                return f"Error: Indeks section {section_index} tidak ditemukan. Total section: {len(doc.sections)}."
            target_sections.append(doc.sections[section_index])
            target_desc = f"section {section_index}"
        else:
            target_sections.extend(doc.sections)
            target_desc = f"seluruh section ({len(doc.sections)} section)"

        for sec in target_sections:
            apply_section_formatting(
                sec=sec,
                paper_size=paper_size,
                page_width_inches=page_width_inches,
                page_height_inches=page_height_inches,
                page_width_cm=page_width_cm,
                page_height_cm=page_height_cm,
                orientation=orientation,
                margin_preset=margin_preset,
                top_margin_inches=top_margin_inches,
                bottom_margin_inches=bottom_margin_inches,
                left_margin_inches=left_margin_inches,
                right_margin_inches=right_margin_inches,
                top_margin_cm=top_margin_cm,
                bottom_margin_cm=bottom_margin_cm,
                left_margin_cm=left_margin_cm,
                right_margin_cm=right_margin_cm,
                gutter_inches=gutter_inches,
                gutter_cm=gutter_cm,
                header_distance_inches=header_distance_inches,
                header_distance_cm=header_distance_cm,
                footer_distance_inches=footer_distance_inches,
                footer_distance_cm=footer_distance_cm,
                different_first_page_header_footer=different_first_page_header_footer,
                columns_count=columns_count,
                column_spacing_inches=column_spacing_inches,
                column_spacing_cm=column_spacing_cm,
                page_number_start=page_number_start,
                page_number_format=page_number_format,
            )

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        return f"Berhasil memperbarui pengaturan halaman pada {target_desc} di {path}."
    except Exception as e:
        return f"Error: Gagal mengatur halaman di {file_path}: {e}"


@mcp.tool()
def add_section_docx(
    file_path: str,
    start_type: str = "new_page",
    paper_size: str = "",
    page_width_inches: float | None = None,
    page_height_inches: float | None = None,
    page_width_cm: float | None = None,
    page_height_cm: float | None = None,
    orientation: str = "",
    margin_preset: str = "",
    top_margin_inches: float | None = None,
    bottom_margin_inches: float | None = None,
    left_margin_inches: float | None = None,
    right_margin_inches: float | None = None,
    top_margin_cm: float | None = None,
    bottom_margin_cm: float | None = None,
    left_margin_cm: float | None = None,
    right_margin_cm: float | None = None,
    gutter_inches: float | None = None,
    gutter_cm: float | None = None,
    different_first_page_header_footer: bool | None = None,
    columns_count: int | None = None,
    column_spacing_inches: float | None = None,
    column_spacing_cm: float | None = None,
    page_number_start: int | None = None,
    page_number_format: str = "",
) -> str:
    """Menambahkan section baru ke dokumen Word (misal halaman baru dengan orientasi Landscape, format 2 kolom, atau restart nomor halaman)"""
    try:
        path = validate_docx_path(file_path, must_exist=False)
        doc = Document(str(path)) if path.is_file() else Document()

        start_map = {
            "new_page": WD_SECTION_START.NEW_PAGE,
            "continuous": WD_SECTION_START.CONTINUOUS,
            "even_page": WD_SECTION_START.EVEN_PAGE,
            "odd_page": WD_SECTION_START.ODD_PAGE,
        }
        st_val = start_map.get(start_type.strip().lower(), WD_SECTION_START.NEW_PAGE)

        new_sec = doc.add_section(st_val)
        apply_section_formatting(
            sec=new_sec,
            paper_size=paper_size,
            page_width_inches=page_width_inches,
            page_height_inches=page_height_inches,
            page_width_cm=page_width_cm,
            page_height_cm=page_height_cm,
            orientation=orientation,
            margin_preset=margin_preset,
            top_margin_inches=top_margin_inches,
            bottom_margin_inches=bottom_margin_inches,
            left_margin_inches=left_margin_inches,
            right_margin_inches=right_margin_inches,
            top_margin_cm=top_margin_cm,
            bottom_margin_cm=bottom_margin_cm,
            left_margin_cm=left_margin_cm,
            right_margin_cm=right_margin_cm,
            gutter_inches=gutter_inches,
            gutter_cm=gutter_cm,
            different_first_page_header_footer=different_first_page_header_footer,
            columns_count=columns_count,
            column_spacing_inches=column_spacing_inches,
            column_spacing_cm=column_spacing_cm,
            page_number_start=page_number_start,
            page_number_format=page_number_format,
        )

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        sec_idx = len(doc.sections) - 1
        return f"Berhasil menambahkan section baru (section {sec_idx}, tipe={start_type}) di {path}."
    except Exception as e:
        return f"Error: Gagal menambahkan section ke {file_path}: {e}"


@mcp.tool()
def add_page_break_docx(
    file_path: str,
    break_type: str = "page",
    paragraph_index: int | None = None,
    table_index: int | None = None,
    row: int | None = None,
    col: int | None = None,
) -> str:
    """Menyisipkan page break (pindah halaman), column break, atau line break ke dokumen .docx pada akhir dokumen, sebelum paragraf tertentu, atau di dalam sel tabel."""
    try:
        path = validate_docx_path(file_path, must_exist=False)
        doc = Document(str(path)) if path.is_file() else Document()

        type_map = {
            "page": WD_BREAK.PAGE,
            "column": WD_BREAK.COLUMN,
            "line": WD_BREAK.LINE,
            "text_wrapping": WD_BREAK.LINE,
        }
        b_type = type_map.get(break_type.lower().strip(), WD_BREAK.PAGE)
        type_label = break_type.lower().strip()

        if table_index is not None:
            if table_index < 0 or table_index >= len(doc.tables):
                return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."
            table = doc.tables[table_index]
            if row is None or col is None:
                return "Error: row dan col harus ditentukan jika table_index diisi."
            if row < 0 or row >= len(table.rows):
                return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
            if col < 0 or col >= len(table.rows[row].cells):
                return f"Error: Indeks kolom {col} tidak valid. Total kolom pada baris {row}: {len(table.rows[row].cells)}."
            cell = table.rows[row].cells[col]
            p = cell.add_paragraph()
            p.add_run().add_break(b_type)
            path.parent.mkdir(parents=True, exist_ok=True)
            doc.save(str(path))
            return f"Berhasil menyisipkan {type_label} break di dalam tabel {table_index} sel [{row}, {col}] pada {path}."

        if paragraph_index is not None:
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return f"Error: Indeks paragraf {paragraph_index} tidak ditemukan. Total paragraf: {len(doc.paragraphs)}."
            target_p = doc.paragraphs[paragraph_index]
            new_p = target_p.insert_paragraph_before()
            new_p.add_run().add_break(b_type)
            path.parent.mkdir(parents=True, exist_ok=True)
            doc.save(str(path))
            return f"Berhasil menyisipkan {type_label} break sebelum paragraf {paragraph_index} pada {path}."

        p = doc.add_paragraph()
        p.add_run().add_break(b_type)
        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        return f"Berhasil menyisipkan {type_label} break di akhir dokumen {path}."
    except Exception as e:
        return f"Error: Gagal menyisipkan break ke {file_path}: {e}"


@mcp.tool()
def add_horizontal_rule_docx(
    file_path: str,
    style: str = "single",
    size_pt: float = 1.5,
    color: str = "000000",
    space_before_pt: float = 0.0,
    space_after_pt: float = 6.0,
    border_position: str = "bottom",
    paragraph_index: int | None = None,
    apply_to_existing: bool = False,
    indent_left_inches: float | None = None,
    indent_right_inches: float | None = None,
    alignment: str = "",
) -> str:
    """Menyisipkan garis horizontal (rule / pembatas bab) atau menambahkan garis pembatas kop surat/paragraf dengan gaya (single, double, triple, thick, thick_thin, thin_thick, thin_thick_thin, wave, dotted, dashed), ketebalan (pt), indentasi, posisi border, dan warna kustom."""
    try:
        path = validate_docx_path(file_path, must_exist=False)
        doc = Document(str(path)) if path.is_file() else Document()

        style_map = {
            "single": "single",
            "double": "double",
            "triple": "triple",
            "thick": "thick",
            "dotted": "dotted",
            "dashed": "dashed",
            "dash": "dashed",
            "dot": "dotted",
            "wave": "wave",
            "double_wave": "doubleWave",
            "doublewave": "doubleWave",
            "thick_thin": "thickThinMediumGap",
            "thickthin": "thickThinMediumGap",
            "thick_thin_med": "thickThinMediumGap",
            "thin_thick": "thinThickMediumGap",
            "thinthick": "thinThickMediumGap",
            "thin_thick_med": "thinThickMediumGap",
            "thin_thick_thin": "thinThickThinMediumGap",
            "thinthickthin": "thinThickThinMediumGap",
            "thick_between_thin": "thinThickThinMediumGap",
            "none": "none",
            "nil": "none",
        }
        val = style_map.get(style.lower().strip(), style.strip())
        sz_val = str(max(1, int(round(size_pt * 8))))

        clean_color = "000000"
        if color:
            try:
                rgb = parse_rgb_color(color)
                clean_color = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
            except Exception:
                clean_color = "000000"

        if apply_to_existing:
            if paragraph_index is None:
                return "Error: paragraph_index wajib diisi jika apply_to_existing bernilai True."
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return f"Error: Indeks paragraf {paragraph_index} tidak ditemukan. Total paragraf: {len(doc.paragraphs)}."
            p = doc.paragraphs[paragraph_index]
            location_desc = f"pada paragraf {paragraph_index}"
        else:
            if paragraph_index is not None:
                if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                    return f"Error: Indeks paragraf {paragraph_index} tidak ditemukan. Total paragraf: {len(doc.paragraphs)}."
                p = doc.paragraphs[paragraph_index].insert_paragraph_before()
                location_desc = f"sebelum paragraf {paragraph_index}"
            else:
                p = doc.add_paragraph()
                location_desc = "di akhir dokumen"
            p.paragraph_format.space_before = Pt(max(0.0, space_before_pt))
            p.paragraph_format.space_after = Pt(max(0.0, space_after_pt))
            p.paragraph_format.line_spacing = Pt(1)

        if indent_left_inches is not None:
            p.paragraph_format.left_indent = Inches(max(0.0, indent_left_inches))
        if indent_right_inches is not None:
            p.paragraph_format.right_indent = Inches(max(0.0, indent_right_inches))
        if alignment:
            align_map = {
                "left": WD_ALIGN_PARAGRAPH.LEFT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            }
            if alignment.lower().strip() in align_map:
                p.alignment = align_map[alignment.lower().strip()]

        pPr = p._p.get_or_add_pPr()
        pBdr = pPr.find(qn("w:pBdr"))
        if pBdr is None:
            pBdr = OxmlElement("w:pBdr")
            pPr.insert_element_before(pBdr, "w:shd", "w:tabs", "w:spacing", "w:ind", "w:jc", "w:rPr", "w:sectPr")

        positions = []
        pos_lower = border_position.lower().strip()
        if pos_lower in ("top", "above"):
            positions = ["top"]
        elif pos_lower in ("top_bottom", "both"):
            positions = ["top", "bottom"]
        elif pos_lower in ("box", "all"):
            positions = ["top", "left", "bottom", "right"]
        else:
            positions = ["bottom"]

        space_attr = "1" if not apply_to_existing else str(max(1, int(round(space_after_pt))))

        for pos in positions:
            tag_name = qn(f"w:{pos}")
            border_elm = pBdr.find(tag_name)
            if border_elm is None:
                border_elm = OxmlElement(f"w:{pos}")
                pBdr.append(border_elm)
            border_elm.set(qn("w:val"), val)
            border_elm.set(qn("w:sz"), sz_val)
            border_elm.set(qn("w:space"), space_attr)
            border_elm.set(qn("w:color"), clean_color)

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        return f"Berhasil menambahkan garis pembatas ({style}, {size_pt}pt, #{clean_color}, posisi={border_position}) {location_desc} pada {path}."
    except Exception as e:
        return f"Error: Gagal menambahkan garis pembatas ke {file_path}: {e}"


def convert_to_omml_element(math_content: str, is_block: bool = False):
    s = math_content.strip()
    if s.startswith("<m:oMath") or s.startswith("<m:oMathPara"):
        omml_xml = s
        if "xmlns:m=" not in omml_xml:
            first_space = omml_xml.find(" ")
            first_close = omml_xml.find(">")
            idx = first_space if (first_space != -1 and first_space < first_close) else first_close
            omml_xml = omml_xml[:idx] + f" {nsdecls('m')}" + omml_xml[idx:]
    elif "<math" in s:
        omml_xml = mathml2omml.convert(s)
        if "xmlns:m=" not in omml_xml:
            omml_xml = omml_xml.replace("<m:oMath", f"<m:oMath {nsdecls('m')}", 1)
    else:
        mathml_str = latex2mathml.converter.convert(s)
        omml_xml = mathml2omml.convert(mathml_str)
        if "xmlns:m=" not in omml_xml:
            omml_xml = omml_xml.replace("<m:oMath", f"<m:oMath {nsdecls('m')}", 1)

    if is_block and not omml_xml.startswith("<m:oMathPara"):
        omml_xml = f"<m:oMathPara {nsdecls('m')}>{omml_xml}</m:oMathPara>"

    return parse_xml(omml_xml)


@mcp.tool()
def add_equation_docx(
    file_path: str,
    equation: str,
    display_mode: str = "block",
    text_before: str = "",
    text_after: str = "",
    equation_number: str = "",
    placeholder: str = "",
    font_size_pt: float | None = None,
    alignment: str = "",
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    paragraph_index: int | None = None,
    insert_before: bool = False,
    table_index: int | None = None,
    row: int | None = None,
    col: int | None = None,
) -> str:
    """Menyisipkan rumus matematika (Equation) ke file .docx dari format LaTeX, MathML, atau OMML. Mendukung mode block (persamaan mandiri di tengah) dan inline (sejajar teks), penomoran persamaan rata kanan ala jurnal/skripsi (equation_number), penggantian placeholder teks (placeholder), ukuran font (font_size_pt), perataan, serta penempatan di paragraf atau sel tabel."""
    try:
        path = validate_docx_path(file_path, must_exist=False)
        doc = Document(str(path)) if path.is_file() else Document()

        if not equation or not equation.strip():
            return "Error: Parameter equation tidak boleh kosong."

        is_block = (display_mode.lower().strip() in ("block", "display")) and (not text_before and not text_after and not equation_number and not placeholder)

        try:
            omml_elm = convert_to_omml_element(equation, is_block=is_block)
        except Exception as conv_err:
            return f"Error: Gagal mengonversi rumus matematika '{equation}': {conv_err}"

        def apply_font_size(paragraph, size_pt):
            if size_pt is not None:
                sz_val = str(int(round(size_pt * 2)))
                pPr = paragraph._p.get_or_add_pPr()
                rPr = pPr.find(qn("w:rPr"))
                if rPr is None:
                    rPr = OxmlElement("w:rPr")
                    pPr.append(rPr)
                sz = rPr.find(qn("w:sz"))
                if sz is None:
                    sz = OxmlElement("w:sz")
                    rPr.append(sz)
                sz.set(qn("w:val"), sz_val)
                for r in paragraph.runs:
                    r.font.size = Pt(size_pt)

        # Case A: Ganti Placeholder
        if placeholder:
            replaced_count = 0
            target_paragraphs = []
            if table_index is not None:
                if table_index < 0 or table_index >= len(doc.tables):
                    return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."
                t = doc.tables[table_index]
                if row is not None and col is not None:
                    if row < 0 or row >= len(t.rows) or col < 0 or col >= len(t.rows[row].cells):
                        return f"Error: Sel [{row}, {col}] tidak valid."
                    target_paragraphs.extend(t.rows[row].cells[col].paragraphs)
                else:
                    for r_item in t.rows:
                        for c_item in r_item.cells:
                            target_paragraphs.extend(c_item.paragraphs)
            elif paragraph_index is not None:
                if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                    return f"Error: Indeks paragraf {paragraph_index} tidak ditemukan."
                target_paragraphs.append(doc.paragraphs[paragraph_index])
            else:
                target_paragraphs.extend(doc.paragraphs)
                for t in doc.tables:
                    for r_item in t.rows:
                        for c_item in r_item.cells:
                            target_paragraphs.extend(c_item.paragraphs)

            for p in target_paragraphs:
                if placeholder in p.text:
                    full_text = p.text
                    before, after = full_text.split(placeholder, 1)
                    p.text = ""
                    if before:
                        p.add_run(before)
                    p._p.append(deepcopy(omml_elm))
                    if after:
                        p.add_run(after)
                    apply_font_size(p, font_size_pt)
                    replaced_count += 1
                    break

            if replaced_count == 0:
                return f"Error: Placeholder '{placeholder}' tidak ditemukan dalam target dokumen."
            path.parent.mkdir(parents=True, exist_ok=True)
            doc.save(str(path))
            return f"Berhasil mengganti placeholder '{placeholder}' dengan rumus matematika pada {path}."

        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }

        # Case B: Di dalam tabel
        if table_index is not None:
            if table_index < 0 or table_index >= len(doc.tables):
                return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."
            table = doc.tables[table_index]
            if row is None or col is None:
                return "Error: row dan col wajib diisi jika table_index ditentukan."
            if row < 0 or row >= len(table.rows):
                return f"Error: Indeks baris {row} tidak valid. Total baris: {len(table.rows)}."
            if col < 0 or col >= len(table.rows[row].cells):
                return f"Error: Indeks kolom {col} tidak valid. Total kolom pada baris {row}: {len(table.rows[row].cells)}."
            cell = table.rows[row].cells[col]
            if len(cell.paragraphs) == 1 and not cell.paragraphs[0].text.strip() and not cell.paragraphs[0]._p.xpath(".//m:oMath"):
                p = cell.paragraphs[0]
            else:
                p = cell.add_paragraph()
            target_desc = f"di dalam tabel {table_index} sel [{row}, {col}]"

        # Case C: Pada paragraf tertentu
        elif paragraph_index is not None:
            if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
                return f"Error: Indeks paragraf {paragraph_index} tidak ditemukan. Total paragraf: {len(doc.paragraphs)}."
            if insert_before:
                p = doc.paragraphs[paragraph_index].insert_paragraph_before()
                target_desc = f"sebelum paragraf {paragraph_index}"
            else:
                p = doc.paragraphs[paragraph_index]
                target_desc = f"pada paragraf {paragraph_index}"

        # Case D: Di akhir dokumen
        else:
            p = doc.add_paragraph()
            target_desc = "di akhir dokumen"

        # Academic / journal numbered equation with tab stops
        if equation_number:
            sec = doc.sections[0] if doc.sections else None
            if sec and sec.page_width and sec.left_margin and sec.right_margin:
                usable_width = sec.page_width - sec.left_margin - sec.right_margin
            else:
                usable_width = Inches(6.5)
            mid_pos = Length(int(usable_width / 2))

            p.paragraph_format.tab_stops.add_tab_stop(mid_pos, WD_TAB_ALIGNMENT.CENTER)
            p.paragraph_format.tab_stops.add_tab_stop(usable_width, WD_TAB_ALIGNMENT.RIGHT)

            r_lead = p.add_run()
            r_lead.add_tab()
            if text_before:
                r_lead.add_text(text_before)

            p._p.append(omml_elm)

            r_num = p.add_run()
            r_num.add_tab()
            r_num.add_text(equation_number)
            if text_after:
                r_num.add_text(f" {text_after}")
            target_desc += f" dengan nomor '{equation_number}'"
        else:
            if text_before:
                p.add_run(text_before)

            p._p.append(omml_elm)

            if text_after:
                p.add_run(text_after)

            if alignment and alignment.lower().strip() in align_map:
                p.alignment = align_map[alignment.lower().strip()]
            elif is_block and not alignment:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        apply_font_size(p, font_size_pt)

        if space_before_pt is not None:
            p.paragraph_format.space_before = Pt(max(0.0, space_before_pt))
        if space_after_pt is not None:
            p.paragraph_format.space_after = Pt(max(0.0, space_after_pt))

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        return f"Berhasil menyisipkan rumus matematika ({display_mode}) {target_desc} pada {path}."
    except Exception as e:
        return f"Error: Gagal menyisipkan rumus matematika ke {file_path}: {e}"


def populate_header_footer_container(
    container,
    text: str = "",
    include_page_number: bool = False,
    include_total_pages: bool = False,
    page_number_format_text: str = "",
    alignment: str = "",
    font_name: str | None = None,
    font_size_pt: float | None = None,
    color: str | None = None,
    bold: bool = False,
    italic: bool = False,
):
    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }

    p = container.paragraphs[0]
    p.text = ""
    for child in list(p._p):
        if child.tag.split("}")[-1] != "pPr":
            p._p.remove(child)
    for extra_p in list(container.paragraphs[1:]):
        extra_p._p.getparent().remove(extra_p._p)

    if alignment and alignment.lower().strip() in align_map:
        p.alignment = align_map[alignment.lower().strip()]

    def create_field_element(instr: str):
        fld = OxmlElement("w:fldSimple")
        fld.set(qn("w:instr"), instr)
        r_elm = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        if font_name:
            rf = OxmlElement("w:rFonts")
            rf.set(qn("w:ascii"), font_name)
            rf.set(qn("w:hAnsi"), font_name)
            rPr.append(rf)
        if font_size_pt and font_size_pt > 0:
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), str(int(round(font_size_pt * 2))))
            rPr.append(sz)
        if bold:
            rPr.append(OxmlElement("w:b"))
        if italic:
            rPr.append(OxmlElement("w:i"))
        if color:
            try:
                rgb = parse_rgb_color(color)
                c_elm = OxmlElement("w:color")
                c_elm.set(qn("w:val"), f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}")
                rPr.append(c_elm)
            except Exception:
                pass
        r_elm.append(rPr)
        t_elm = OxmlElement("w:t")
        t_elm.text = "1"
        r_elm.append(t_elm)
        fld.append(r_elm)
        return fld

    def add_styled_text_run(run_text: str):
        r = p.add_run(run_text)
        apply_font_to_run(r, font_name=font_name, font_size_pt=font_size_pt, color=color, bold=bold, italic=italic)
        return r

    if page_number_format_text:
        import re
        tokens = re.split(r"(\{page\}|\{pages\})", page_number_format_text)
        for token in tokens:
            if token == "{page}":
                p._p.append(create_field_element("PAGE"))
            elif token == "{pages}":
                p._p.append(create_field_element("NUMPAGES"))
            elif token:
                add_styled_text_run(token)
    else:
        if text:
            add_styled_text_run(text)
            if include_page_number:
                add_styled_text_run(" ")

        if include_page_number:
            p._p.append(create_field_element("PAGE"))

        if include_total_pages:
            add_styled_text_run(" / ")
            p._p.append(create_field_element("NUMPAGES"))


@mcp.tool()
def set_header_footer_docx(
    file_path: str,
    target: str = "footer",
    section_index: int | None = None,
    header_footer_type: str = "default",
    preset: str = "",
    text: str = "",
    include_page_number: bool = False,
    include_total_pages: bool = False,
    page_number_format_text: str = "",
    alignment: str = "",
    font_name: str | None = None,
    font_size_pt: float | None = None,
    color: str | None = None,
    bold: bool = False,
    italic: bool = False,
    is_linked_to_previous: bool | None = None,
) -> str:
    """Mengatur isi Header atau Footer dan menyisipkan field nomor halaman fisik (PAGE / NUMPAGES) pada dokumen .docx. Mendukung halaman pertama berbeda (different_first_page), unlink section, font styling, perataan, serta preset standar skripsi (skripsi_chapter, skripsi_preliminary)."""
    try:
        path = validate_docx_path(file_path, must_exist=False)
        doc = Document(str(path)) if path.is_file() else Document()

        if section_index is not None:
            if section_index < 0 or section_index >= len(doc.sections):
                return f"Error: Indeks section {section_index} tidak ditemukan. Total section: {len(doc.sections)}."
            target_sections = [(section_index, doc.sections[section_index])]
        else:
            target_sections = list(enumerate(doc.sections))

        preset_clean = preset.lower().strip()

        for s_idx, sec in target_sections:
            if is_linked_to_previous is not None:
                sec.header.is_linked_to_previous = is_linked_to_previous
                sec.footer.is_linked_to_previous = is_linked_to_previous
                sec.first_page_header.is_linked_to_previous = is_linked_to_previous
                sec.first_page_footer.is_linked_to_previous = is_linked_to_previous

            if preset_clean == "skripsi_chapter":
                sec.different_first_page_header_footer = True
                if s_idx > 0 and is_linked_to_previous is None:
                    sec.header.is_linked_to_previous = False
                    sec.footer.is_linked_to_previous = False

                f_name = font_name or "Times New Roman"
                f_size = font_size_pt or 12.0

                populate_header_footer_container(
                    sec.first_page_footer,
                    include_page_number=True,
                    alignment="center",
                    font_name=f_name,
                    font_size_pt=f_size,
                    color=color,
                    bold=bold,
                    italic=italic,
                )
                populate_header_footer_container(sec.first_page_header)
                populate_header_footer_container(
                    sec.header,
                    text=text,
                    include_page_number=True,
                    alignment="right",
                    font_name=f_name,
                    font_size_pt=f_size,
                    color=color,
                    bold=bold,
                    italic=italic,
                )
                populate_header_footer_container(sec.footer)

            elif preset_clean == "skripsi_preliminary":
                sec.different_first_page_header_footer = False
                if s_idx > 0 and is_linked_to_previous is None:
                    sec.header.is_linked_to_previous = False
                    sec.footer.is_linked_to_previous = False

                f_name = font_name or "Times New Roman"
                f_size = font_size_pt or 12.0

                populate_header_footer_container(sec.header)
                populate_header_footer_container(
                    sec.footer,
                    include_page_number=True,
                    alignment="center",
                    font_name=f_name,
                    font_size_pt=f_size,
                    color=color,
                    bold=bold,
                    italic=italic,
                )

            elif preset_clean == "page_number_bottom_center":
                populate_header_footer_container(
                    sec.footer,
                    text=text,
                    include_page_number=True,
                    include_total_pages=include_total_pages,
                    alignment="center",
                    font_name=font_name,
                    font_size_pt=font_size_pt,
                    color=color,
                    bold=bold,
                    italic=italic,
                )
            elif preset_clean == "page_number_top_right":
                populate_header_footer_container(
                    sec.header,
                    text=text,
                    include_page_number=True,
                    include_total_pages=include_total_pages,
                    alignment="right",
                    font_name=font_name,
                    font_size_pt=font_size_pt,
                    color=color,
                    bold=bold,
                    italic=italic,
                )
            elif preset_clean == "page_number_bottom_right":
                populate_header_footer_container(
                    sec.footer,
                    text=text,
                    include_page_number=True,
                    include_total_pages=include_total_pages,
                    alignment="right",
                    font_name=font_name,
                    font_size_pt=font_size_pt,
                    color=color,
                    bold=bold,
                    italic=italic,
                )
            elif preset_clean in ("clear", "clear_all"):
                populate_header_footer_container(sec.header)
                populate_header_footer_container(sec.footer)
                populate_header_footer_container(sec.first_page_header)
                populate_header_footer_container(sec.first_page_footer)
            else:
                tgt_clean = target.lower().strip()
                hf_type = header_footer_type.lower().strip()

                if tgt_clean == "header":
                    if hf_type == "first_page":
                        sec.different_first_page_header_footer = True
                        container = sec.first_page_header
                    elif hf_type == "even_page":
                        container = sec.even_page_header
                    else:
                        container = sec.header
                else:
                    if hf_type == "first_page":
                        sec.different_first_page_header_footer = True
                        container = sec.first_page_footer
                    elif hf_type == "even_page":
                        container = sec.even_page_footer
                    else:
                        container = sec.footer

                populate_header_footer_container(
                    container,
                    text=text,
                    include_page_number=include_page_number,
                    include_total_pages=include_total_pages,
                    page_number_format_text=page_number_format_text,
                    alignment=alignment,
                    font_name=font_name,
                    font_size_pt=font_size_pt,
                    color=color,
                    bold=bold,
                    italic=italic,
                )

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        sec_desc = f"section {section_index}" if section_index is not None else f"seluruh ({len(target_sections)}) section"
        mode_desc = f"preset '{preset}'" if preset else f"{target} ({header_footer_type})"
        return f"Berhasil mengatur Header/Footer ({mode_desc}) pada {sec_desc} di {path}."
    except Exception as e:
        return f"Error: Gagal mengatur Header/Footer pada {file_path}: {e}"


@mcp.tool()
def get_header_footer_docx(file_path: str, section_index: int | None = None) -> str:
    """Membaca isi teks, nomor halaman (field PAGE/NUMPAGES), status link, dan konfigurasi Header dan Footer untuk seluruh section dokumen .docx."""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if section_index is not None:
            if section_index < 0 or section_index >= len(doc.sections):
                return f"Error: Indeks section {section_index} tidak ditemukan. Total section: {len(doc.sections)}."
            target_sections = [(section_index, doc.sections[section_index])]
        else:
            target_sections = list(enumerate(doc.sections))

        def format_container_info(container, label: str):
            items = []
            for p in container.paragraphs:
                txt = p.text.strip()
                flds = [el.get(qn("w:instr")) for el in p._p.xpath(".//w:fldSimple")]
                fld_str = " ".join([f"[{fi}]" for fi in flds]) if flds else ""
                full = f"{txt} {fld_str}".strip()
                if full:
                    items.append(full)
            content_desc = " | ".join(items) if items else "(kosong)"
            align_desc = container.paragraphs[0].alignment if container.paragraphs else None
            align_str = f", Align={align_desc.name}" if align_desc else ""
            link_str = "Linked" if container.is_linked_to_previous else "Unlinked"
            return f"{label}: {content_desc} ({link_str}{align_str})"

        lines = [f"Informasi Header & Footer Dokumen {path.name}:"]
        for idx, sec in target_sections:
            lines.append(f"\n[Section {idx}]")
            lines.append(f"- Different First Page: {'Ya' if sec.different_first_page_header_footer else 'Tidak'}")
            lines.append(f"- {format_container_info(sec.header, 'Header Normal')}")
            if sec.different_first_page_header_footer:
                lines.append(f"- {format_container_info(sec.first_page_header, 'Header Hal. Pertama')}")
            lines.append(f"- {format_container_info(sec.footer, 'Footer Normal')}")
            if sec.different_first_page_header_footer:
                lines.append(f"- {format_container_info(sec.first_page_footer, 'Footer Hal. Pertama')}")

        return "\n".join(lines)
    except Exception as e:
        return f"Error: Gagal membaca Header/Footer di {file_path}: {e}"


if __name__ == "__main__":
    mcp.run()




