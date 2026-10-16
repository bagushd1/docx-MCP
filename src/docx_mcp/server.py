import os
from pathlib import Path
from mcp.server.mcpserver import MCPServer
from docx import Document

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


if __name__ == "__main__":
    mcp.run()
