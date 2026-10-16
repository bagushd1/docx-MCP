import os
from pathlib import Path
from mcp.server.mcpserver import MCPServer
from copy import deepcopy
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_UNDERLINE

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
        if p.text.strip():
            content.append(f"[P{i}] {p.text}")

    for t_idx, table in enumerate(doc.tables):
        content.append(f"\n--- Tabel {t_idx} ---")
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells]
            content.append(" | ".join(row_text))

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
    line_spacing: float | None = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    first_line_indent_inches: float | None = None,
    left_indent_inches: float | None = None,
    right_indent_inches: float | None = None,
    alignment: str = "",
    list_style: str = "",
) -> str:
    """Mengatur format paragraf: spasi baris (misal 1.5, 2.0), spasi sebelum/sesudah (pt), indentasi (inci), perataan, dan gaya list (bullet/number)"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        target_paragraphs = []
        if paragraph_index is not None:
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
            if line_spacing is not None and line_spacing > 0:
                p.paragraph_format.line_spacing = line_spacing
            if space_before_pt is not None:
                p.paragraph_format.space_before = Pt(space_before_pt)
            if space_after_pt is not None:
                p.paragraph_format.space_after = Pt(space_after_pt)
            if first_line_indent_inches is not None:
                p.paragraph_format.first_line_indent = Inches(first_line_indent_inches)
            if left_indent_inches is not None:
                p.paragraph_format.left_indent = Inches(left_indent_inches)
            if right_indent_inches is not None:
                p.paragraph_format.right_indent = Inches(right_indent_inches)

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
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    first_line_indent_inches: float | None = None,
    left_indent_inches: float | None = None,
    right_indent_inches: float | None = None,
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

        if line_spacing is not None and line_spacing > 0:
            p.paragraph_format.line_spacing = line_spacing
        if space_before_pt is not None:
            p.paragraph_format.space_before = Pt(space_before_pt)
        if space_after_pt is not None:
            p.paragraph_format.space_after = Pt(space_after_pt)
        if first_line_indent_inches is not None:
            p.paragraph_format.first_line_indent = Inches(first_line_indent_inches)
        if left_indent_inches is not None:
            p.paragraph_format.left_indent = Inches(left_indent_inches)
        if right_indent_inches is not None:
            p.paragraph_format.right_indent = Inches(right_indent_inches)

        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
        return f"Berhasil menambahkan paragraf berformat ke {path}."
    except Exception as e:
        return f"Error: Gagal menambahkan paragraf berformat di {file_path}: {e}"


if __name__ == "__main__":
    mcp.run()
