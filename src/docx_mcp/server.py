import os
from pathlib import Path
from mcp.server.mcpserver import MCPServer
from docx import Document
from docx.shared import Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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
) -> str:
    """Menambahkan satu kolom baru pada tabel docx dengan teks sel opsional dan lebar opsional"""
    try:
        path = validate_docx_path(file_path, must_exist=True)
        doc = Document(str(path))

        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: Indeks tabel {table_index} tidak ditemukan. Total tabel: {len(doc.tables)}."

        table = doc.tables[table_index]
        width_val = Inches(width_inches) if width_inches and width_inches > 0 else None
        table.add_column(width=width_val)

        if width_val:
            table.columns[-1].width = width_val

        if values:
            for idx, val in enumerate(values[: len(table.rows)]):
                cell = table.rows[idx].cells[-1]
                cell.text = str(val)
                if width_val:
                    cell.width = width_val
        elif width_val:
            for row in table.rows:
                row.cells[-1].width = width_val

        doc.save(str(path))
        new_col_idx = len(table.columns) - 1
        return f"Berhasil menambahkan kolom baru pada indeks {new_col_idx} di tabel {table_index} pada {path}."
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


if __name__ == "__main__":
    mcp.run()
