import tempfile
import unittest
from pathlib import Path
from docx_mcp.server import (
    clear_docx,
    create_table_docx,
    delete_table_docx,
    edit_table_docx,
    read_docx,
)


class TestTableTools(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.TemporaryDirectory()
        self.file_path = str(Path(self.tmp_dir.name) / "test.docx")

    def tearDown(self):
        self.tmp_dir.cleanup()

    def test_create_and_read_table(self):
        data = [
            ["Nama", "Role"],
            ["Alice", "Admin"],
            ["Bob", "User"],
        ]
        res = create_table_docx(self.file_path, data, heading="Daftar Pengguna")
        self.assertIn("Berhasil membuat tabel", res)

        content = read_docx(self.file_path)
        self.assertIn("Daftar Pengguna", content)
        self.assertIn("--- Tabel 0 ---", content)
        self.assertIn("Nama | Role", content)
        self.assertIn("Alice | Admin", content)

    def test_create_table_empty_data(self):
        res = create_table_docx(self.file_path, [])
        self.assertIn("Error: Data tabel tidak boleh kosong", res)

    def test_edit_table(self):
        data = [["Col1", "Col2"], ["Val1", "Val2"]]
        create_table_docx(self.file_path, data)

        res = edit_table_docx(self.file_path, table_index=0, row=1, col=1, new_text="UpdatedVal")
        self.assertIn("Berhasil mengubah sel [1, 1]", res)

        content = read_docx(self.file_path)
        self.assertIn("Val1 | UpdatedVal", content)

    def test_edit_table_invalid_bounds(self):
        data = [["Col1", "Col2"], ["Val1", "Val2"]]
        create_table_docx(self.file_path, data)

        self.assertIn("Error: Indeks tabel 99 tidak ditemukan", edit_table_docx(self.file_path, 99, 0, 0, "X"))
        self.assertIn("Error: Indeks baris 5 tidak valid", edit_table_docx(self.file_path, 0, 5, 0, "X"))
        self.assertIn("Error: Indeks kolom 5 tidak valid", edit_table_docx(self.file_path, 0, 0, 5, "X"))

    def test_delete_table(self):
        data1 = [["T1"]]
        data2 = [["T2"]]
        create_table_docx(self.file_path, data1)
        create_table_docx(self.file_path, data2)

        res = delete_table_docx(self.file_path, table_index=0)
        self.assertIn("Berhasil menghapus tabel 0", res)

        content = read_docx(self.file_path)
        self.assertNotIn("T1", content)
        self.assertIn("T2", content)

    def test_delete_table_invalid_index(self):
        create_table_docx(self.file_path, [["A"]])
        res = delete_table_docx(self.file_path, table_index=3)
        self.assertIn("Error: Indeks tabel 3 tidak ditemukan", res)

    def test_clear_docx(self):
        create_table_docx(self.file_path, [["A", "B"]], heading="Header")
        res = clear_docx(self.file_path)
        self.assertIn("Berhasil mengosongkan seluruh isi file", res)
        self.assertEqual(read_docx(self.file_path), "")


if __name__ == "__main__":
    unittest.main()
