import base64
import tempfile
import unittest
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE
from docx_mcp.server import (
    add_column_docx,
    add_picture_docx,
    add_row_docx,
    add_styled_paragraph_docx,
    clear_docx,
    create_table_docx,
    delete_column_docx,
    delete_picture_docx,
    delete_row_docx,
    delete_table_docx,
    edit_table_docx,
    extract_images_docx,
    format_paragraph_docx,
    format_text_docx,
    get_images_info_docx,
    merge_cells_docx,
    read_docx,
    resize_picture_docx,
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

    def test_add_styled_paragraph(self):
        res = add_styled_paragraph_docx(
            self.file_path,
            text="Heading Bergaya",
            heading_level=1,
            font_name="Times New Roman",
            font_size_pt=18.0,
            color="#003366",
            bold=True,
            italic=True,
            alignment="center",
        )
        self.assertIn("Berhasil menambahkan paragraf berformat", res)

    def test_format_text_paragraph(self):
        add_styled_paragraph_docx(self.file_path, text="Teks awal sebelum diformat.")
        res = format_text_docx(
            self.file_path,
            paragraph_index=0,
            font_name="Arial",
            font_size_pt=12.0,
            color="red",
            bold=True,
            underline=True,
            strike=True,
            highlight_color="yellow",
        )
        self.assertIn("Berhasil memformat teks pada paragraf 0", res)
        self.assertIn("font='Arial'", res)
        self.assertIn("size=12.0pt", res)

    def test_format_text_target_phrase(self):
        add_styled_paragraph_docx(self.file_path, text="Kalimat ini ada kata kunci penting di dalamnya.")
        res = format_text_docx(
            self.file_path,
            target_text="kata kunci",
            font_name="Courier New",
            font_size_pt=14.0,
            color="#FF5500",
            bold=True,
        )
        self.assertIn("Berhasil memformat teks pada seluruh dokumen untuk teks 'kata kunci'", res)

    def test_format_text_table_cell(self):
        create_table_docx(self.file_path, [["Nama", "Skor"], ["Budi", "100"]])
        res = format_text_docx(
            self.file_path,
            table_index=0,
            row=0,
            font_name="Arial",
            bold=True,
            color="#FFFFFF",
        )
        self.assertIn("tabel 0 baris 0", res)

    def test_format_text_typography(self):
        add_styled_paragraph_docx(self.file_path, text="H2O formula and special text")
        res = format_text_docx(
            self.file_path,
            paragraph_index=0,
            subscript=True,
            small_caps=True,
            underline_style="double",
        )
        self.assertIn("subscript=True", res)
        self.assertIn("small_caps=True", res)
        self.assertIn("underline_style='double'", res)

        doc = Document(self.file_path)
        run = doc.paragraphs[0].runs[0]
        self.assertTrue(run.font.subscript)
        self.assertTrue(run.font.small_caps)
        self.assertEqual(run.font.underline, WD_UNDERLINE.DOUBLE)

    def test_format_paragraph(self):
        add_styled_paragraph_docx(self.file_path, text="Paragraf uji pemformatan spasi.")
        res = format_paragraph_docx(
            self.file_path,
            paragraph_index=0,
            line_spacing=1.5,
            space_before_pt=6.0,
            space_after_pt=12.0,
            first_line_indent_inches=0.5,
            left_indent_inches=0.25,
            alignment="justify",
            list_style="bullet",
        )
        self.assertIn("Berhasil memformat tata letak dan spasi pada paragraf 0", res)

        doc = Document(self.file_path)
        p = doc.paragraphs[0]
        self.assertEqual(p.paragraph_format.line_spacing, 1.5)
        self.assertAlmostEqual(p.paragraph_format.space_before.pt, 6.0)
        self.assertAlmostEqual(p.paragraph_format.space_after.pt, 12.0)
        self.assertAlmostEqual(p.paragraph_format.first_line_indent.inches, 0.5)
        self.assertAlmostEqual(p.paragraph_format.left_indent.inches, 0.25)
        self.assertEqual(p.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)

    def test_add_styled_paragraph_advanced(self):
        res = add_styled_paragraph_docx(
            self.file_path,
            text="E=mc2 formula",
            line_spacing=2.0,
            space_before_pt=4.0,
            space_after_pt=10.0,
            first_line_indent_inches=0.5,
            left_indent_inches=0.3,
            right_indent_inches=0.2,
            superscript=True,
            all_caps=True,
            underline_style="dotted",
            alignment="center",
        )
        self.assertIn("Berhasil menambahkan paragraf berformat", res)

        doc = Document(self.file_path)
        p = doc.paragraphs[0]
        self.assertEqual(p.paragraph_format.line_spacing, 2.0)
        self.assertAlmostEqual(p.paragraph_format.space_before.pt, 4.0)
        self.assertAlmostEqual(p.paragraph_format.space_after.pt, 10.0)
        self.assertAlmostEqual(p.paragraph_format.first_line_indent.inches, 0.5)
        self.assertAlmostEqual(p.paragraph_format.left_indent.inches, 0.3)
        self.assertAlmostEqual(p.paragraph_format.right_indent.inches, 0.2)
        self.assertEqual(p.alignment, WD_ALIGN_PARAGRAPH.CENTER)

        run = p.runs[0]
        self.assertTrue(run.font.superscript)
        self.assertTrue(run.font.all_caps)
        self.assertEqual(run.font.underline, WD_UNDERLINE.DOTTED)

    def test_format_paragraph_table_and_flow_control(self):
        create_table_docx(self.file_path, [["H1", "H2"], ["A", "B"]])
        res = format_paragraph_docx(
            self.file_path,
            table_index=0,
            row=1,
            col=1,
            line_spacing_pt=14.0,
            alignment="center",
        )
        self.assertIn("tabel 0 sel [1, 1]", res)

        doc = Document(self.file_path)
        p = doc.tables[0].rows[1].cells[1].paragraphs[0]
        self.assertAlmostEqual(p.paragraph_format.line_spacing.pt, 14.0)
        self.assertEqual(p.alignment, WD_ALIGN_PARAGRAPH.CENTER)

    def test_format_paragraph_hanging_and_flow(self):
        add_styled_paragraph_docx(self.file_path, text="Daftar pustaka referensi...")
        res = format_paragraph_docx(
            self.file_path,
            paragraph_index=0,
            hanging_indent_inches=0.5,
            keep_with_next=True,
            page_break_before=True,
            widow_control=True,
        )
        self.assertIn("Berhasil memformat tata letak dan spasi pada paragraf 0", res)

        doc = Document(self.file_path)
        p = doc.paragraphs[0]
        self.assertAlmostEqual(p.paragraph_format.left_indent.inches, 0.5)
        self.assertAlmostEqual(p.paragraph_format.first_line_indent.inches, -0.5)
        self.assertTrue(p.paragraph_format.keep_with_next)
        self.assertTrue(p.paragraph_format.page_break_before)
        self.assertTrue(p.paragraph_format.widow_control)

    def test_add_picture_basic_and_read(self):
        img_path = Path(self.tmp_dir.name) / "test.png"
        img_path.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="))

        res = add_picture_docx(
            self.file_path,
            str(img_path),
            width_inches=2.0,
            height_inches=1.5,
            alignment="center",
            caption="Gambar 1: Sampel Visual",
        )
        self.assertIn("Berhasil menyisipkan gambar 'test.png'", res)

        content = read_docx(self.file_path)
        self.assertIn("Total Gambar: 1", content)
        self.assertIn("Gambar 1: Sampel Visual", content)

        doc = Document(self.file_path)
        self.assertEqual(len(doc.inline_shapes), 1)
        shape = doc.inline_shapes[0]
        self.assertAlmostEqual(shape.width / 914400.0, 2.0, places=1)
        self.assertAlmostEqual(shape.height / 914400.0, 1.5, places=1)

    def test_add_picture_table_cell_and_cm(self):
        img_path = Path(self.tmp_dir.name) / "test_table.png"
        img_path.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="))

        create_table_docx(self.file_path, [["Produk", "Foto"]])
        res = add_picture_docx(
            self.file_path,
            str(img_path),
            table_index=0,
            row=0,
            col=1,
            width_cm=4.0,
            alignment="center",
        )
        self.assertIn("tabel 0 sel [0, 1]", res)

        doc = Document(self.file_path)
        cell = doc.tables[0].rows[0].cells[1]
        self.assertEqual(cell.paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)

    def test_get_images_info_and_resize(self):
        img_path = Path(self.tmp_dir.name) / "test_resize.png"
        img_path.write_bytes(base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="))

        add_picture_docx(self.file_path, str(img_path), width_inches=1.0)
        info = get_images_info_docx(self.file_path)
        self.assertIn("Total gambar pada", info)
        self.assertIn("[Gambar 0]", info)

        resize_res = resize_picture_docx(self.file_path, picture_index=0, width_inches=3.0, height_inches=2.0)
        self.assertIn("Berhasil mengubah ukuran gambar 0", resize_res)

        doc = Document(self.file_path)
        shape = doc.inline_shapes[0]
        self.assertAlmostEqual(shape.width / 914400.0, 3.0, places=1)
        self.assertAlmostEqual(shape.height / 914400.0, 2.0, places=1)

    def test_add_picture_base64_and_alt_text(self):
        b64_data = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
        res = add_picture_docx(
            self.file_path,
            image_base64=b64_data,
            width_inches=1.5,
            alt_text="Logo Perusahaan",
            title="Logo Utama",
        )
        self.assertIn("Berhasil menyisipkan gambar base64", res)

        doc = Document(self.file_path)
        self.assertEqual(len(doc.inline_shapes), 1)
        shape = doc.inline_shapes[0]
        self.assertEqual(shape._inline.docPr.get("descr"), "Logo Perusahaan")
        self.assertEqual(shape._inline.docPr.get("title"), "Logo Utama")

    def test_delete_picture(self):
        b64_data = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
        add_picture_docx(self.file_path, image_base64=b64_data)
        doc = Document(self.file_path)
        self.assertEqual(len(doc.inline_shapes), 1)

        res = delete_picture_docx(self.file_path, picture_index=0)
        self.assertIn("Berhasil menghapus gambar 0", res)

        doc_after = Document(self.file_path)
        self.assertEqual(len(doc_after.inline_shapes), 0)

    def test_extract_images(self):
        b64_data = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
        add_picture_docx(self.file_path, image_base64=b64_data)

        out_dir = Path(self.tmp_dir.name) / "extracted_imgs"
        res = extract_images_docx(self.file_path, str(out_dir))
        self.assertIn("Berhasil mengekstrak 1 gambar", res)
        self.assertTrue(out_dir.is_dir())
        files = list(out_dir.glob("*"))
        self.assertEqual(len(files), 1)


if __name__ == "__main__":
    unittest.main()




