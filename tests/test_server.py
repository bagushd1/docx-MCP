import tempfile
import unittest
from pathlib import Path
from docx_mcp.server import (
    add_column_docx,
    add_row_docx,
    clear_docx,
    create_table_docx,
    delete_column_docx,
    delete_row_docx,
    delete_table_docx,
    edit_table_docx,
    merge_cells_docx,
    read_docx,
    resize_table_docx,
    set_cell_alignment_docx,
    set_cell_margins_docx,
    set_cell_shading_docx,
    set_cell_text_direction_docx,
    set_table_alignment_docx,
    set_table_borders_docx,
    set_table_pagination_docx,
    set_table_style_docx,
    set_table_visibility_docx,
    split_table_docx,
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

    def test_add_column(self):
        create_table_docx(self.file_path, [["A", "B"], ["1", "2"]])
        res = add_column_docx(self.file_path, table_index=0, values=["C", "3"], width_inches=1.5)
        self.assertIn("Berhasil menambahkan kolom baru", res)
        content = read_docx(self.file_path)
        self.assertIn("A | B | C", content)
        self.assertIn("1 | 2 | 3", content)

    def test_resize_table(self):
        create_table_docx(self.file_path, [["A", "B"]])
        res1 = resize_table_docx(self.file_path, table_index=0, col_widths=[1.5, 2.5])
        self.assertIn("Berhasil mengatur lebar 2 kolom", res1)

        res2 = resize_table_docx(self.file_path, table_index=0, col_index=0, width_inches=3.0)
        self.assertIn("Berhasil mengubah lebar kolom 0 menjadi 3.0 inci", res2)

    def test_set_table_visibility(self):
        create_table_docx(self.file_path, [["A", "B"]])
        res_hide = set_table_visibility_docx(self.file_path, table_index=0, visible=False)
        self.assertIn("transparan (tanpa border)", res_hide)

        res_show = set_table_visibility_docx(self.file_path, table_index=0, visible=True)
        self.assertIn("tampak (bergaris)", res_show)

    def test_delete_column(self):
        create_table_docx(self.file_path, [["A", "B", "C"], ["1", "2", "3"]])
        res = delete_column_docx(self.file_path, table_index=0, col_index=1)
        self.assertIn("Berhasil menghapus kolom 1", res)

        content = read_docx(self.file_path)
        self.assertIn("A | C", content)
        self.assertIn("1 | 3", content)

    def test_delete_column_last_remaining(self):
        create_table_docx(self.file_path, [["A"]])
        res = delete_column_docx(self.file_path, table_index=0, col_index=0)
        self.assertIn("Error: Tabel hanya memiliki 1 kolom", res)

    def test_add_row(self):
        create_table_docx(self.file_path, [["A", "B"]])
        res = add_row_docx(self.file_path, table_index=0, values=["C", "D"], height_inches=0.5)
        self.assertIn("Berhasil menambahkan baris baru pada indeks 1", res)
        content = read_docx(self.file_path)
        self.assertIn("C | D", content)

    def test_delete_row(self):
        create_table_docx(self.file_path, [["R0C0"], ["R1C0"], ["R2C0"]])
        res = delete_row_docx(self.file_path, table_index=0, row_index=1)
        self.assertIn("Berhasil menghapus baris 1", res)
        content = read_docx(self.file_path)
        self.assertIn("R0C0", content)
        self.assertNotIn("R1C0", content)
        self.assertIn("R2C0", content)

    def test_delete_row_last_remaining(self):
        create_table_docx(self.file_path, [["A"]])
        res = delete_row_docx(self.file_path, table_index=0, row_index=0)
        self.assertIn("Error: Tabel hanya memiliki 1 baris", res)

    def test_merge_cells(self):
        create_table_docx(self.file_path, [["A", "B"], ["C", "D"]])
        res = merge_cells_docx(self.file_path, table_index=0, start_row=0, start_col=0, end_row=0, end_col=1)
        self.assertIn("Berhasil menggabungkan sel dari [0, 0] ke [0, 1]", res)

    def test_set_cell_shading(self):
        create_table_docx(self.file_path, [["A", "B"], ["C", "D"]])
        res_cell = set_cell_shading_docx(self.file_path, table_index=0, color_hex="#CCCCCC", row=0, col=0)
        self.assertIn("Berhasil menerapkan warna #CCCCCC pada sel [0, 0]", res_cell)

        res_row = set_cell_shading_docx(self.file_path, table_index=0, color_hex="E0E0E0", row=1)
        self.assertIn("Berhasil menerapkan warna #E0E0E0 pada seluruh baris 1", res_row)

        res_invalid = set_cell_shading_docx(self.file_path, table_index=0, color_hex="INVALID")
        self.assertIn("Error: Format warna 'INVALID' tidak valid", res_invalid)

    def test_set_table_style(self):
        create_table_docx(self.file_path, [["A", "B"]])
        res = set_table_style_docx(self.file_path, table_index=0, style_name="Table Grid")
        self.assertIn("Berhasil mengubah style tabel 0", res)

    def test_set_table_alignment(self):
        create_table_docx(self.file_path, [["A", "B"]])
        res = set_table_alignment_docx(self.file_path, table_index=0, alignment="center")
        self.assertIn("Berhasil mengubah perataan tabel 0 menjadi 'center'", res)

    def test_set_cell_alignment(self):
        create_table_docx(self.file_path, [["A", "B"], ["C", "D"]])
        res = set_cell_alignment_docx(self.file_path, table_index=0, horizontal="center", vertical="center", row=0, col=0)
        self.assertIn("Berhasil mengubah perataan pada sel [0, 0]", res)

    def test_set_table_pagination(self):
        create_table_docx(self.file_path, [["Header 1", "Header 2"], ["Data 1", "Data 2"]])
        res = set_table_pagination_docx(self.file_path, table_index=0, repeat_header_rows=1, cant_split=True)
        self.assertIn("Berhasil mengatur opsi pagination", res)

    def test_insert_row_and_column_at_index(self):
        create_table_docx(self.file_path, [["A", "C"], ["1", "3"]])
        res_col = add_column_docx(self.file_path, table_index=0, values=["B", "2"], insert_index=1)
        self.assertIn("Berhasil menambahkan kolom baru pada indeks 1", res_col)
        content1 = read_docx(self.file_path)
        self.assertIn("A | B | C", content1)

        res_row = add_row_docx(self.file_path, table_index=0, values=["X", "Y", "Z"], insert_index=1)
        self.assertIn("Berhasil menambahkan baris baru pada indeks 1", res_row)
        content2 = read_docx(self.file_path)
        self.assertIn("X | Y | Z", content2)

    def test_set_cell_margins(self):
        create_table_docx(self.file_path, [["A", "B"]])
        res_table = set_cell_margins_docx(self.file_path, table_index=0, top_pt=5.0, bottom_pt=5.0)
        self.assertIn("Berhasil mengatur margin padding", res_table)

        res_cell = set_cell_margins_docx(self.file_path, table_index=0, row=0, col=0, left_pt=10.0)
        self.assertIn("pada sel [0, 0]", res_cell)

    def test_set_table_borders(self):
        create_table_docx(self.file_path, [["A", "B"]])
        res = set_table_borders_docx(self.file_path, table_index=0, color_hex="FF0000", size=8, top=False, bottom=True)
        self.assertIn("Berhasil mengatur border kustom", res)

    def test_set_cell_text_direction(self):
        create_table_docx(self.file_path, [["A", "B"]])
        res = set_cell_text_direction_docx(self.file_path, table_index=0, direction="bt_lr", row=0, col=0)
        self.assertIn("Berhasil mengubah arah teks menjadi 'bt_lr' pada sel [0, 0]", res)

    def test_split_table(self):
        create_table_docx(self.file_path, [["R0"], ["R1"], ["R2"], ["R3"]])
        res = split_table_docx(self.file_path, table_index=0, split_row_index=2)
        self.assertIn("Berhasil membagi tabel 0 pada baris 2", res)
        self.assertIn("memiliki 2 tabel", res)


if __name__ == "__main__":
    unittest.main()
