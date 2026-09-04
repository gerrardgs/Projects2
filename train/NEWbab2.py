#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generator BAB II TINJAUAN PUSTAKA (MODIFIKASI TERBARU)
Topik: Monitoring Stunting pada Balita di Kalimantan Tengah Menggunakan Logika Fuzzy
       (Studi Kasus Data SKI 2023)

Modifikasi terbaru:
- Pivot dari Klasifikasi -> Monitoring
- 6 variabel: Usia, Jenis Kelamin, TBL, MP-ASI, TB ukur, BB ukur
- Dual Z-Score: saat lahir + saat pengukuran
- Tambahan teori Algoritma Genetika (GA) untuk optimasi boundary
- 108 aturan fuzzy per JK x 2 JK = 216 efektif
- Pre-filtering Jenis Kelamin

Mengikuti format UPDATED_TEMPLATE_BUKU_TA_IF_TUS:
- Times New Roman 12pt, line spacing 1.5
- Margin: kiri 4cm, kanan 3cm, atas 3cm, bawah 3cm
- Format tabel/gambar/persamaan: Tabel II.x, Gambar II.x, (II.x)
"""

import os
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================================
# Style constants
# ============================================================================
FONT_NAME = "Times New Roman"
FONT_SIZE_BODY = 12
FONT_SIZE_HEADING = 14
FONT_SIZE_TABLE = 11
LINE_SPACING = 1.5


def set_cell_borders(cell, color="000000", sz="4"):
    """Set borders for a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ('top', 'left', 'bottom', 'right'):
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), sz)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)


def set_cell_shading(cell, color_hex):
    """Set background shading for a cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)


def add_run(paragraph, text, bold=False, italic=False, size=FONT_SIZE_BODY, font=FONT_NAME):
    """Add a styled text run to a paragraph."""
    run = paragraph.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    # Make sure font applies to East Asian text too
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font)
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    run.bold = bold
    run.italic = italic
    return run


def set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                        first_line_indent_cm=None, line_spacing=LINE_SPACING,
                        space_before=0, space_after=0):
    """Set paragraph format."""
    p.alignment = alignment
    pf = p.paragraph_format
    pf.line_spacing = line_spacing
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_line_indent_cm is not None:
        pf.first_line_indent = Cm(first_line_indent_cm)


def add_para(doc, text_segments=None, plain_text=None,
             alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
             first_line_indent=1.27, line_spacing=LINE_SPACING,
             space_before=0, space_after=0, size=FONT_SIZE_BODY):
    """Add a paragraph. text_segments is list of (text, bold, italic) tuples
       OR plain_text is single string."""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=alignment,
                          first_line_indent_cm=first_line_indent,
                          line_spacing=line_spacing,
                          space_before=space_before, space_after=space_after)
    if plain_text is not None:
        add_run(p, plain_text, size=size)
    elif text_segments is not None:
        for seg in text_segments:
            if len(seg) == 3:
                txt, bold, italic = seg
                add_run(p, txt, bold=bold, italic=italic, size=size)
            elif len(seg) == 4:
                txt, bold, italic, sz = seg
                add_run(p, txt, bold=bold, italic=italic, size=sz)
    return p


def add_heading_chapter(doc, text_line1, text_line2):
    """Add chapter heading 'BAB II' over 'TINJAUAN PUSTAKA' (centered, bold)."""
    p1 = doc.add_paragraph()
    set_paragraph_format(p1, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                          first_line_indent_cm=0, line_spacing=LINE_SPACING,
                          space_before=0, space_after=0)
    add_run(p1, text_line1, bold=True, size=FONT_SIZE_HEADING)

    p2 = doc.add_paragraph()
    set_paragraph_format(p2, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                          first_line_indent_cm=0, line_spacing=LINE_SPACING,
                          space_before=0, space_after=24)
    add_run(p2, text_line2, bold=True, size=FONT_SIZE_HEADING)
    return p1, p2


def add_heading_section(doc, num, text):
    """Add level-2 heading like '2.1 Literatur Terkait Teori'."""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                          first_line_indent_cm=0, line_spacing=LINE_SPACING,
                          space_before=12, space_after=6)
    add_run(p, f"{num}   {text}", bold=True, size=FONT_SIZE_BODY)
    return p


def add_heading_subsection(doc, num, text, italic_first=False):
    """Add level-3 heading like '2.1.1 Penelitian Terdahulu'."""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                          first_line_indent_cm=0, line_spacing=LINE_SPACING,
                          space_before=12, space_after=6)
    add_run(p, f"{num}   {text}", bold=True, size=FONT_SIZE_BODY)
    return p


def add_heading_subsubsection(doc, num, text):
    """Add level-4 heading like '2.6.1 Fuzzyfikasi'."""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                          first_line_indent_cm=0, line_spacing=LINE_SPACING,
                          space_before=10, space_after=4)
    add_run(p, f"{num}   {text}", bold=True, italic=False, size=FONT_SIZE_BODY)
    return p


def add_figure_caption(doc, num, caption):
    """Caption for figure: 'Gambar II.x Caption text' centered."""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                          first_line_indent_cm=0, line_spacing=LINE_SPACING,
                          space_before=6, space_after=12)
    add_run(p, f"Gambar II.{num} ", bold=True, size=FONT_SIZE_BODY)
    add_run(p, caption, size=FONT_SIZE_BODY)
    return p


def add_table_caption(doc, num, caption):
    """Caption for table: 'Tabel II.x Caption text' centered, ABOVE the table."""
    p = doc.add_paragraph()
    set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                          first_line_indent_cm=0, line_spacing=LINE_SPACING,
                          space_before=6, space_after=6)
    add_run(p, f"Tabel II.{num} ", bold=True, size=FONT_SIZE_BODY)
    add_run(p, caption, size=FONT_SIZE_BODY)
    return p


def add_equation(doc, equation_text, number):
    """Add equation in a 2-column borderless table: equation | (II.x)."""
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Configure widths
    table.columns[0].width = Cm(12)
    table.columns[1].width = Cm(2.5)
    for row in table.rows:
        row.cells[0].width = Cm(12)
        row.cells[1].width = Cm(2.5)

    # Eqn cell
    cell_eq = table.cell(0, 0)
    cell_eq.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_eq = cell_eq.paragraphs[0]
    set_paragraph_format(p_eq, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                         first_line_indent_cm=0, line_spacing=1.15)
    add_run(p_eq, equation_text, italic=True, size=FONT_SIZE_BODY)

    # Number cell
    cell_num = table.cell(0, 1)
    cell_num.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_num = cell_num.paragraphs[0]
    set_paragraph_format(p_num, alignment=WD_ALIGN_PARAGRAPH.RIGHT,
                         first_line_indent_cm=0, line_spacing=1.15)
    add_run(p_num, f"(II.{number})", bold=True, size=FONT_SIZE_BODY)

    # Spacing after
    p_spacer = doc.add_paragraph()
    set_paragraph_format(p_spacer, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first_line_indent_cm=0, line_spacing=LINE_SPACING,
                         space_before=0, space_after=0)
    return table


def add_table_from_data(doc, headers, rows, col_widths_cm=None,
                        header_color="D9E2F3", font_size=FONT_SIZE_TABLE,
                        first_col_italic=False):
    """Add a table from headers + data rows.
       headers: list of strings (or list of (text,bold,italic) tuples)
       rows:    list of list of strings
    """
    n_cols = len(headers)
    n_rows = 1 + len(rows)
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Column widths
    if col_widths_cm is None:
        total_w = 15.0
        col_widths_cm = [total_w / n_cols] * n_cols

    for ci, w in enumerate(col_widths_cm):
        table.columns[ci].width = Cm(w)

    # Header row
    for ci, h in enumerate(headers):
        cell = table.cell(0, ci)
        cell.width = Cm(col_widths_cm[ci])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_borders(cell)
        set_cell_shading(cell, header_color)
        # remove default empty para
        cell.paragraphs[0].text = ""
        p = cell.paragraphs[0]
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.CENTER,
                              first_line_indent_cm=0, line_spacing=1.15)
        if isinstance(h, str):
            add_run(p, h, bold=True, size=font_size)
        else:
            for seg in h:
                txt, b, i = seg
                add_run(p, txt, bold=b, italic=i, size=font_size)

    # Data rows
    for ri, row_data in enumerate(rows, start=1):
        for ci, cell_data in enumerate(row_data):
            cell = table.cell(ri, ci)
            cell.width = Cm(col_widths_cm[ci])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_borders(cell)
            cell.paragraphs[0].text = ""
            p = cell.paragraphs[0]
            set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                                  first_line_indent_cm=0, line_spacing=1.15)
            if isinstance(cell_data, str):
                add_run(p, cell_data, size=font_size,
                        italic=(first_col_italic and ci == 0))
            else:
                for seg in cell_data:
                    if len(seg) == 3:
                        txt, b, i = seg
                        add_run(p, txt, bold=b, italic=i, size=font_size)
                    elif len(seg) == 4:
                        txt, b, i, sz = seg
                        add_run(p, txt, bold=b, italic=i, size=sz)

    # Spacing after
    p_spacer = doc.add_paragraph()
    set_paragraph_format(p_spacer, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                         first_line_indent_cm=0, line_spacing=LINE_SPACING,
                         space_before=0, space_after=0)
    return table


def add_bullet_list(doc, items, indent_cm=1.0):
    """Add bulleted list items."""
    for item in items:
        p = doc.add_paragraph(style='List Bullet')
        set_paragraph_format(p, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
                              first_line_indent_cm=0, line_spacing=LINE_SPACING,
                              space_before=0, space_after=0)
        p.paragraph_format.left_indent = Cm(indent_cm)
        if isinstance(item, str):
            add_run(p, item, size=FONT_SIZE_BODY)
        else:
            for seg in item:
                if len(seg) == 3:
                    txt, b, i = seg
                    add_run(p, txt, bold=b, italic=i, size=FONT_SIZE_BODY)


# ============================================================================
# DOCUMENT SETUP
# ============================================================================
def setup_document():
    doc = Document()

    # Set default font via styles
    style = doc.styles['Normal']
    style.font.name = FONT_NAME
    style.font.size = Pt(FONT_SIZE_BODY)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), FONT_NAME)
    rFonts.set(qn('w:hAnsi'), FONT_NAME)
    rFonts.set(qn('w:eastAsia'), FONT_NAME)

    # A4 page with template margins
    for section in doc.sections:
        section.page_height = Cm(29.7)
        section.page_width  = Cm(21.0)
        section.top_margin    = Cm(3.0)
        section.bottom_margin = Cm(3.0)
        section.left_margin   = Cm(4.0)
        section.right_margin  = Cm(3.0)

    return doc


# ============================================================================
# CONTENT BUILDER
# ============================================================================
def build_bab2(doc):
    # === Chapter heading ===
    add_heading_chapter(doc, "BAB II", "TINJAUAN PUSTAKA")

    # Bridge paragraph
    add_para(doc,
        plain_text=(
            "Bab ini menguraikan kajian pustaka yang melandasi penelitian, mulai dari penelitian "
            "terdahulu yang relevan, konsep stunting beserta indikator antropometri, perhitungan "
            "Z-Score WHO, deskripsi dataset Survei Kesehatan Indonesia (SKI) 2023, variabel penelitian, "
            "logika fuzzy Sugeno, optimasi berbasis algoritma genetika, hingga teknik validasi dan "
            "metrik evaluasi yang digunakan. Pembahasan ditutup dengan alasan pemilihan metode "
            "dalam konteks penelitian monitoring stunting balita di Kalimantan Tengah."
        ),
        first_line_indent=1.27
    )

    # ========================================================================
    # 2.1 LITERATUR TERKAIT TEORI
    # ========================================================================
    add_heading_section(doc, "2.1", "Literatur Terkait Teori")

    # 2.1.1 Penelitian Terdahulu
    add_heading_subsection(doc, "2.1.1", "Penelitian Terdahulu")

    add_para(doc,
        plain_text=(
            "Dalam rangka memperkuat fondasi penelitian yang diusulkan, peneliti mengkaji berbagai "
            "literatur yang berkaitan dengan pemantauan dan klasifikasi risiko stunting pada balita "
            "melalui pendekatan logika fuzzy serta optimasi metaheuristik. Kajian ini meliputi "
            "metode, variabel, ukuran data, dan tingkat akurasi yang diperoleh, baik dari penelitian "
            "berbasis fuzzy Sugeno murni maupun yang dipadukan dengan algoritma genetika (GA) sebagai "
            "metode optimasi. Ringkasan kajian penelitian terdahulu ditunjukkan pada Tabel II.1."
        ),
        first_line_indent=1.27
    )

    # Tabel II.1 - Penelitian Terdahulu (5 entries)
    add_table_caption(doc, 1, "Hasil Kajian Penelitian Terdahulu")

    headers = [
        [("No", True, False)],
        [("Judul, Penulis, Tahun", True, False)],
        [("Metode & Variabel", True, False)],
        [("Hasil Utama", True, False)],
        [("Persamaan / Perbedaan dengan Penelitian Ini", True, False)],
    ]
    rows = [
        [
            "1",
            [
                ("Prediksi Status Gizi Balita Menggunakan Logika Fuzzy Metode Sugeno Berdasarkan Pengukuran Antropometri ", False, True),
                ("(Safal et al., 2024)", False, False),
            ],
            [
                ("Fuzzy Sugeno; 2 variabel (berat dan tinggi badan); fungsi keanggotaan segitiga; defuzzifikasi ", False, False),
                ("weighted average", False, True),
                (".", False, False),
            ],
            "Akurasi 81% pada 37 sampel balita di Puskesmas Kandai, Kendari.",
            [
                ("Persamaan: ", True, False),
                ("sama-sama memakai Fuzzy Sugeno. ", False, False),
                ("Perbedaan: ", True, False),
                ("hanya 2 variabel antropometri tanpa variabel determinan (BBL, PBL, MP-ASI), tanpa optimasi parameter, sampel kecil.", False, False),
            ],
        ],
        [
            "2",
            [
                ("Deteksi Gizi Buruk pada Balita Menggunakan Metode Fuzzy Logic (Studi Kasus Puskesmas Kecamatan Semidang Alas Kabupaten Seluma) ", False, True),
                ("(Reswan et al., 2024)", False, False),
            ],
            [("Fuzzy Sugeno DSS; variabel usia, berat, tinggi badan; antarmuka web PHP/MySQL.", False, False)],
            "Sistem mampu mendeteksi gizi buruk balita dengan keluaran nilai defuzzifikasi yang konsisten dibanding perhitungan manual.",
            [
                ("Persamaan: ", True, False),
                ("Fuzzy Sugeno untuk indikator antropometri. ", False, False),
                ("Perbedaan: ", True, False),
                ("tidak menggunakan variabel riwayat lahir maupun MP-ASI, lingkup hanya satu puskesmas, tidak ada optimasi parameter.", False, False),
            ],
        ],
        [
            "3",
            [
                ("Sistem Pakar Diagnosis Stunting pada Balita Menggunakan Metode Forward Chaining dan Logika Fuzzy Sugeno ", False, True),
                ("(Ananda & Sriani, 2024)", False, False),
            ],
            [
                (" Forward chaining + Fuzzy Sugeno; variabel Z-score BB/U, TB/U, BB/TB; 48 aturan; output tingkat risiko (rendah, tinggi, sangat tinggi).", False, False),
            ],
            "Dari 122 data balita Puskesmas Pematang Panjang, 106 terdiagnosis stunting (57% tinggi, 26% sangat tinggi, 17% rendah).",
            [
                ("Persamaan: ", True, False),
                ("menggunakan Z-score WHO dan Sugeno. ", False, False),
                ("Perbedaan: ", True, False),
                ("variabel input ringan, tanpa optimasi GA, data terbatas satu puskesmas.", False, False),
            ],
        ],
        [
            "4",
            [
                ("Mobile Decision Support System to Determine Toddler's Nutrition using Fuzzy Sugeno ", False, True),
                ("(Surhajito et al., 2017)", False, False),
            ],
            [
                ("Aplikasi Android; Fuzzy Sugeno; 5 variabel (usia, berat, tinggi, lingkar kepala, jenis kelamin); 145 aturan; defuzzifikasi ", False, False),
                ("weighted average", False, True),
                (".", False, False),
            ],
            "Hasil sistem konsisten dengan tabel antropometri WHO; menjadi alat bantu mobile monitoring status gizi.",
            [
                ("Persamaan: ", True, False),
                ("Fuzzy Sugeno multivariat dengan integrasi WHO. ", False, False),
                ("Perbedaan: ", True, False),
                ("tidak melibatkan riwayat lahir/MP-ASI, parameter fuzzy bersifat tetap, tidak ada metode optimasi otomatis.", False, False),
            ],
        ],
        [
            "5",
            [
                ("Genetic Algorithm Optimization of Fuzzy Membership Function for Multivariate Health Risk Classification ", False, True),
                ("(Suresh et al., 2023)", False, False),
            ],
            [
                ("Fuzzy Sugeno + Genetic Algorithm; optimasi titik batas fungsi keanggotaan (", False, False),
                ("boundary points", False, True),
                (") menggunakan GA dengan fitness berbasis F1-Score.", False, False),
            ],
            "GA terbukti meningkatkan F1-Score sebesar 7-15% dibanding parameter fuzzy konvensional; ruang pencarian kontinu lebih unggul dari grid search.",
            [
                ("Persamaan: ", True, False),
                ("pendekatan GA untuk tuning ", False, False),
                ("boundary points", False, True),
                (" fungsi keanggotaan tanpa mengubah arsitektur fuzzy. ", False, False),
                ("Perbedaan: ", True, False),
                ("domain berbeda (kesehatan umum, bukan stunting); tidak menggunakan struktur monitoring dual Z-Score.", False, False),
            ],
        ],
    ]
    add_table_from_data(doc, headers, rows,
                        col_widths_cm=[0.8, 3.5, 3.2, 3.0, 3.5],
                        font_size=10)

    # Narrative paragraph per study (5 paragraphs)
    add_para(doc, text_segments=[
        ("Penelitian pertama oleh ", False, False),
        ("Safal et al. (2024)", False, False),
        (" mengembangkan sistem prediksi status gizi balita menggunakan logika fuzzy Sugeno dengan dua variabel antropometri utama, yaitu tinggi badan dan berat badan, yang dibedakan menurut usia dan jenis kelamin sesuai Permenkes RI No. 2 Tahun 2022. Sistem dibangun melalui tahapan akuisisi pengetahuan, fuzzyfikasi dengan fungsi keanggotaan segitiga, implikasi minimum, agregasi maksimum, dan defuzzifikasi ", False, False),
        ("weighted average", False, True),
        (" dengan konsekuen ", False, False),
        ("singleton", False, True),
        (". Pengujian pada 37 balita Puskesmas Kandai Kendari menghasilkan akurasi 81%, membuktikan bahwa pendekatan Sugeno mampu menghasilkan klasifikasi yang lebih fleksibel dibandingkan ambang batas Z-Score konvensional. Penelitian ini menjadi referensi penting bagi pendekatan defuzzifikasi Sugeno yang sederhana namun efektif.", False, False),
    ], first_line_indent=1.27)

    add_para(doc, text_segments=[
        ("Penelitian kedua oleh ", False, False),
        ("Reswan et al. (2024)", False, False),
        (" mengembangkan aplikasi ", False, False),
        ("Decision Support System", False, True),
        (" berbasis Fuzzy Sugeno untuk mendeteksi gizi buruk pada balita usia 0-5 tahun di Puskesmas Kecamatan Semidang Alas, Kabupaten Seluma. Variabel input meliputi usia, berat badan, dan tinggi badan yang diolah melalui tahapan fuzzyfikasi, penentuan derajat keanggotaan minimum, dan defuzzifikasi. Aplikasi dibangun dengan PHP dan MySQL, dilengkapi antarmuka pengguna dan divalidasi melalui studi kasus serta perhitungan manual. Sistem mampu memberikan deteksi gizi buruk secara konsisten, sehingga layak dijadikan alat bantu awal pemantauan status gizi balita. Penelitian ini mengkonfirmasi nilai praktis Sugeno untuk deteksi gizi anak, namun belum mempertimbangkan variabel determinan riwayat kelahiran maupun MP-ASI.", False, False),
    ], first_line_indent=1.27)

    add_para(doc, text_segments=[
        ("Penelitian ketiga oleh ", False, False),
        ("Ananda dan Sriani (2024)", False, False),
        (" mengembangkan sistem pakar berbasis web untuk diagnosis stunting dengan mengkombinasikan metode ", False, False),
        ("forward chaining", False, True),
        (" sebagai penalaran berbasis fakta dan Fuzzy Sugeno untuk menentukan tingkat keparahan risiko. Diagnosis menggunakan indikator antropometri WHO (BB/U, TB/U, BB/TB) yang dikonversi ke Z-Score, kemudian difuzzifikasi ke himpunan linguistik dan diproses melalui 48 aturan dengan operator minimum dan defuzzifikasi ", False, False),
        ("weighted average", False, True),
        (" untuk menghasilkan level risiko rendah, tinggi, dan sangat tinggi. Pengujian pada 122 data balita Puskesmas Pematang Panjang menghasilkan distribusi 57% tinggi, 26% sangat tinggi, dan 17% rendah. Penelitian ini menegaskan validitas Sugeno yang dipadukan dengan Z-Score WHO sebagai fondasi penilaian risiko stunting, sebuah pendekatan yang juga diadopsi dalam penelitian ini.", False, False),
    ], first_line_indent=1.27)

    add_para(doc, text_segments=[
        ("Penelitian keempat oleh ", False, False),
        ("Surhajito et al. (2017)", False, False),
        (" mengusulkan ", False, False),
        ("mobile decision support system", False, True),
        (" berbasis Android untuk menentukan status gizi balita dengan Fuzzy Sugeno. Sistem memanfaatkan lima indikator antropometri (usia, berat badan, tinggi badan, lingkar kepala, jenis kelamin) yang diproses melalui 145 aturan fuzzy dan defuzzifikasi ", False, False),
        ("weighted average", False, True),
        (" menjadi kategori status gizi (buruk, kurang, sedang, baik, lebih). Data uji dari Klinik Silalas Medan menunjukkan keluaran Sugeno konsisten dengan perhitungan manual berbasis WHO. Penelitian ini menjadi rujukan penting bagi pendekatan multivariat fuzzy Sugeno, namun parameter fuzzy bersifat tetap dan belum melibatkan metode optimasi otomatis seperti algoritma genetika.", False, False),
    ], first_line_indent=1.27)

    add_para(doc, text_segments=[
        ("Penelitian kelima oleh ", False, False),
        ("Suresh et al. (2023)", False, False),
        (" menggabungkan logika fuzzy Sugeno dengan ", False, False),
        ("Genetic Algorithm", False, True),
        (" (GA) untuk mengklasifikasikan risiko kesehatan multivariat. GA digunakan secara khusus untuk mengoptimasi titik batas (", False, False),
        ("boundary points", False, True),
        (") fungsi keanggotaan tanpa mengubah arsitektur fuzzy, sehingga interpretabilitas pakar tetap terjaga. Fungsi ", False, False),
        ("fitness", False, True),
        (" yang digunakan berbasis F1-Score pada data uji. Hasil eksperimen menunjukkan peningkatan F1-Score sebesar 7-15% dibandingkan parameter fuzzy konvensional yang dirancang secara manual. Penelitian ini menjadi rujukan kunci bagi pendekatan optimasi yang diadopsi dalam penelitian ini, di mana GA dimanfaatkan untuk menyesuaikan ", False, False),
        ("boundary points", False, True),
        (" fungsi keanggotaan terhadap distribusi empiris balita Kalimantan Tengah tanpa mengubah jumlah aturan, jumlah label linguistik, maupun nilai konsekuen.", False, False),
    ], first_line_indent=1.27)

    # ========================================================================
    # 2.2 STUNTING
    # ========================================================================
    add_heading_section(doc, "2.2", "Stunting")

    add_para(doc, text_segments=[
        ("Stunting", False, True),
        (" merupakan kondisi gangguan pertumbuhan dan perkembangan balita yang diakibatkan oleh kekurangan gizi kronis, infeksi berulang, serta kurangnya stimulasi psikososial yang memadai ", False, False),
        ("(Kurniawan, 2025)", False, False),
        (". Penilaian status ", False, False),
        ("stunting", False, True),
        (" umumnya dilakukan berdasarkan pengukuran tinggi badan menurut usia (", False, False),
        ("height-for-age", False, True),
        (") yang dibandingkan dengan ", False, False),
        ("World Health Organization", False, True),
        (" (WHO) ", False, False),
        ("Child Growth Standards", False, True),
        (". Penentuan batas antar kategori secara ilmiah didasarkan pada nilai ", False, False),
        ("Height-for-Age Z-score", False, True),
        (" (HAZ) ", False, False),
        ("(Perumal et al., 2023)", False, False),
        (".", False, False),
    ], first_line_indent=1.27)

    add_para(doc, text_segments=[
        ("Dalam penelitian ini, penilaian status ", False, False),
        ("stunting", False, True),
        (" dilakukan dalam dua kerangka waktu yang saling melengkapi, yaitu kondisi pada saat balita lahir dan kondisi pada saat pengukuran survei berlangsung. Pendekatan ini memungkinkan sistem tidak hanya menilai status ", False, False),
        ("stunting", False, True),
        (" sebagai potret tunggal, tetapi juga sebagai trajektori pertumbuhan, sehingga lebih relevan untuk kebutuhan ", False, False),
        ("monitoring", False, True),
        (" berkelanjutan terhadap balita. Konsep dual penilaian inilah yang membedakan pendekatan ", False, False),
        ("monitoring", False, True),
        (" dengan klasifikasi statis pada penelitian terdahulu.", False, False),
    ], first_line_indent=1.27)

    # ========================================================================
    # 2.3 Z-SCORE
    # ========================================================================
    add_heading_section(doc, "2.3", "Z-Score")

    add_para(doc, text_segments=[
        ("Z-score", False, True),
        (" (", False, False),
        ("standard score", False, True),
        (") merupakan ukuran baku yang menunjukkan seberapa jauh suatu nilai menyimpang dari rerata populasi dalam satuan simpangan baku. Dalam penilaian ", False, False),
        ("stunting", False, True),
        (", ", False, False),
        ("Z-score", False, True),
        (" khususnya ", False, False),
        ("Height-for-Age Z-score", False, True),
        (" (HAZ) digunakan untuk menilai status gizi balita dengan membandingkan tinggi badan terhadap usia berdasarkan standar antropometri WHO ", False, False),
        ("(Watson et al., 2024)", False, False),
        (". Persamaan ", False, False),
        ("Z-score", False, True),
        (" dasar dapat dilihat pada Persamaan (II.1) berikut ini.", False, False),
    ], first_line_indent=1.27)

    add_equation(doc, "Z-Score = (x − μ) / σ", 1)

    add_para(doc, plain_text="dimana:", first_line_indent=0)
    add_bullet_list(doc, [
        [("x", False, True), (" = nilai individu pengukuran dari tinggi atau panjang badan.", False, False)],
        [("μ", False, True), (" = nilai rata-rata (", False, False), ("mean", False, True), (") tinggi badan populasi referensi untuk usia dan jenis kelamin yang sama.", False, False)],
        [("σ", False, True), (" = simpangan baku tinggi badan populasi referensi untuk usia dan jenis kelamin yang sama.", False, False)],
    ])

    add_para(doc, text_segments=[
        ("Mengingat kurva pertumbuhan balita tidak sepenuhnya normal terutama pada usia muda, WHO menggunakan transformasi LMS (", False, False),
        ("Lambda-Mu-Sigma", False, True),
        (") yang menggeser distribusi mendekati normal sebelum perhitungan Z-Score. Persamaan ", False, False),
        ("Z-score", False, True),
        (" versi LMS WHO dapat dilihat pada Persamaan (II.2) berikut.", False, False),
    ], first_line_indent=1.27)

    add_equation(doc, "Z-Score = [(x / M)^L − 1] / (L × S)", 2)

    add_para(doc, plain_text="dimana:", first_line_indent=0)
    add_bullet_list(doc, [
        [("x", False, True), (" = nilai pengukuran antropometri individu (panjang atau tinggi badan dalam cm).", False, False)],
        [("M", False, True), (" (", False, False), ("median", False, True), (") = nilai ", False, False), ("median", False, True), (" populasi referensi untuk kelompok usia dan jenis kelamin tertentu.", False, False)],
        [("L", False, True), (" (", False, False), ("lambda", False, True), (") = parameter ", False, False), ("power", False, True), (" untuk mentransformasi data menuju normalitas.", False, False)],
        [("S", False, True), (" (", False, False), ("sigma", False, True), (") = koefisien variasi (atau simpangan baku yang disesuaikan) populasi referensi.", False, False)],
    ])

    add_para(doc, text_segments=[
        ("Penelitian ini menggunakan rumus ", False, False),
        ("Z-Score", False, True),
        (" LMS WHO pada Persamaan (II.2) sebagai dasar pelabelan ", False, False),
        ("ground truth", False, True),
        (". Lebih jauh lagi, terdapat dua jenis HAZ yang dihitung yaitu ", False, False),
        ("Z-Score Saat Lahir", False, True),
        (" (HAZ awal, menggunakan panjang badan lahir terhadap median WHO usia 0 bulan menurut jenis kelamin) dan ", False, False),
        ("Z-Score Saat Pengukuran", False, True),
        (" (HAZ saat ukur, menggunakan tinggi badan saat survei terhadap median WHO usia pengukuran menurut jenis kelamin). Penggunaan dua HAZ ini menjadi landasan kerangka ", False, False),
        ("monitoring", False, True),
        (" yang membedakan penelitian ini dengan klasifikasi konvensional.", False, False),
    ], first_line_indent=1.27)

    add_para(doc, text_segments=[
        ("Berdasarkan klasifikasi WHO, status ", False, False),
        ("stunting", False, True),
        (" balita ditentukan dengan ambang batas HAZ yang ditunjukkan pada Tabel II.2.", False, False),
    ], first_line_indent=1.27)

    add_table_caption(doc, 2, "Daftar Klasifikasi Status Stunting Balita berdasarkan HAZ")
    headers2 = [
        [("Indeks Antropometri", True, False)],
        [("Status Gizi", True, False)],
        [("Ambang Batas / ", True, False), ("Threshold", True, True), (" (Z-Score)", True, False)],
    ]
    rows2 = [
        [
            "Panjang/Tinggi Badan menurut Usia (PB/U atau TB/U) balita usia 0-59 bulan (Soliman et al., 2021)",
            [("Pendek (", False, False), ("Stunted", False, True), ("/", False, False), ("Stunting", False, True), (")", False, False)],
            "x ≤ −2,00 SD",
        ],
        [
            "",
            "Normal",
            "x > −2,00 SD",
        ],
    ]
    add_table_from_data(doc, headers2, rows2,
                        col_widths_cm=[7.0, 3.5, 4.0])

    # ========================================================================
    # 2.4 SKI 2023
    # ========================================================================
    add_heading_section(doc, "2.4", "Survei Kesehatan Indonesia (SKI) 2023")

    add_para(doc, text_segments=[
        ("Survei Kesehatan Indonesia (SKI) 2023 merupakan survei nasional yang diselenggarakan oleh Kementerian Kesehatan untuk memetakan kondisi kesehatan masyarakat Indonesia melalui 86.364 observasi rumah tangga yang merepresentasikan faktor sosial, ekonomi, lingkungan, dan demografi yang mempengaruhi status gizi serta indikator kesehatan lainnya ", False, False),
        ("(Erda et al., 2024)", False, False),
        (". Survei ini disajikan dalam dua buku panduan resmi seperti ditunjukkan pada Gambar II.1.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 1,
        "Buku Panduan Survei Kesehatan Indonesia: (a) Buku Pedoman Fundamental SKI 2023 dan "
        "(b) Buku Pedoman Deskripsi SKI 2023 dalam Angka (Kemenkes, 2024).")

    add_para(doc, text_segments=[
        ("Survei SKI dilaksanakan dalam bentuk kuesioner. Kuesioner SKI 2023 terdiri dari tiga instrumen utama, yaitu rumah tangga, rumah tangga balita, dan individu, sebagaimana ditunjukkan pada Gambar II.2. Seluruh instrumen ini menjadi dasar penyusunan SKI 2023 dalam Angka dan Laporan Tematik SKI 2023 yang menyajikan ringkasan indikator kesehatan nasional dan analisis khusus, termasuk ", False, False),
        ("stunting", False, True),
        (" ", False, False),
        ("(Kemenkes, 2024)", False, False),
        (". Penelitian ini secara khusus memanfaatkan instrumen individu balita untuk Provinsi Kalimantan Tengah.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 2,
        "Bentuk Kuesioner Survei SKI 2023: (a) Contoh Survei Rumah Tangga, "
        "(b) Contoh Survei Rumah Tangga Balita, dan (c) Contoh Survei Individu (Kemenkes, 2024).")

    # ========================================================================
    # 2.5 VARIABEL PENELITIAN
    # ========================================================================
    add_heading_section(doc, "2.5", "Variabel Penelitian")

    add_para(doc, text_segments=[
        ("Variabel penelitian memegang peran utama dalam memahami determinan ", False, False),
        ("stunting", False, True),
        (". Variabel dependen pada penelitian ini adalah status ", False, False),
        ("stunting", False, True),
        (" yang direpresentasikan oleh ", False, False),
        ("priority flag", False, True),
        (" empat tingkat (", False, False),
        ("Normal, Stunting Low, Stunting Medium, Stunting High", False, True),
        ("), sedangkan variabel independennya terdiri atas enam variabel utama yaitu usia, jenis kelamin, panjang badan lahir, usia pengenalan MP-ASI, tinggi badan saat pengukuran, dan berat badan saat pengukuran. Penjelasan masing-masing variabel diuraikan pada sub-bab berikut.", False, False),
    ], first_line_indent=1.27)

    # 2.5.1 Usia Balita
    add_heading_subsection(doc, "2.5.1", "Usia Balita")
    add_para(doc, text_segments=[
        ("Usia balita dalam satuan bulan merupakan variabel penting karena merefleksikan akumulasi paparan risiko sejak lahir hingga 59 bulan. Secara keseluruhan, dalam rentang 0-59 bulan yang mencakup 1000 HPK, determinan biologis, perilaku, lingkungan, dan sosial-ekonomi terbukti konsisten memengaruhi prevalensi ", False, False),
        ("stunting", False, True),
        (" ", False, False),
        ("(Talukder et al., 2018)", False, False),
        (". Temuan pada Gambar II.3 menunjukkan bahwa prevalensi ", False, False),
        ("stunting", False, True),
        (" tertinggi terjadi pada kelompok usia 24-53 bulan, baik pada 2022 maupun 2023, meskipun angka tahun 2023 lebih rendah. Pada penelitian ini, usia digunakan sebagai variabel ", False, False),
        ("pre-filter", False, True),
        (" untuk perhitungan ", False, False),
        ("Z-Score", False, True),
        (" dan sebagai variabel ", False, False),
        ("input", False, True),
        (" fuzzy yang membentuk konteks evaluasi tinggi dan berat badan.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 3,
        "Prevalensi Stunting Berdasarkan Kelompok Umur Tahun 2022 dan 2023 (Kemenkes, 2024).")

    # 2.5.2 Jenis Kelamin Balita
    add_heading_subsection(doc, "2.5.2", "Jenis Kelamin Balita")
    add_para(doc, text_segments=[
        ("Jenis kelamin merupakan variabel penting karena memengaruhi pola pertumbuhan, kerentanan biologis, serta respons terhadap faktor lingkungan. Bukti empiris menunjukkan bahwa balita laki-laki lebih rentan mengalami ", False, False),
        ("stunting", False, True),
        (" dibandingkan perempuan, dan determinan sosioekonomi bekerja berbeda menurut jenis kelamin ", False, False),
        ("(de Onis & Branca, 2016)", False, False),
        (". Karena itu, interpretasi ", False, False),
        ("Z-Score", False, True),
        (" harus selalu mempertimbangkan jenis kelamin, sebagaimana ditunjukkan pada standar antropometri WHO pada Gambar II.4. Dalam sistem fuzzy yang dibangun, jenis kelamin tidak diperlakukan sebagai variabel fuzzy, melainkan sebagai ", False, False),
        ("pre-filter", False, True),
        (" yang membedakan dua jalur perhitungan terpisah: satu set fungsi keanggotaan untuk balita laki-laki, dan satu set lainnya untuk balita perempuan. Pendekatan ini menjamin bahwa interpretasi linguistik (pendek, normal, tinggi, dan seterusnya) konsisten dengan standar pertumbuhan masing-masing jenis kelamin.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 4,
        "Standar Antropometri WHO menurut Jenis Kelamin: (a) Grafik Standar Panjang/Tinggi Badan Perempuan dan (b) Grafik Standar Panjang/Tinggi Badan Laki-laki (UNICEF, 2023).")

    # 2.5.3 Panjang Badan Lahir
    add_heading_subsection(doc, "2.5.3", "Panjang Badan Lahir")
    add_para(doc, text_segments=[
        ("Panjang Badan Lahir (PBL) merupakan indikator penting dalam penilaian risiko ", False, False),
        ("stunting", False, True),
        (" karena secara konsisten terbukti sebagai prediktor kuat pertumbuhan linear. Bayi dengan panjang lahir pendek memiliki kemungkinan lebih tinggi mengalami ", False, False),
        ("stunting", False, True),
        (" dibandingkan bayi dengan panjang lahir normal atau tinggi. Temuan longitudinal seperti ", False, False),
        ("Women First Trial", False, True),
        (" menegaskan bahwa PBL adalah prediktor terkuat status ", False, False),
        ("linear growth", False, True),
        (" hingga 24 bulan, bahkan melampaui berat lahir, sebagaimana terlihat pada Gambar II.5 yang menunjukkan bahwa ", False, False),
        ("birth LAZ < −2", False, True),
        (" dan ", False, False),
        ("< −1", False, True),
        (" memberikan penurunan skor Z panjang menurut umur (LAZ) paling besar dibandingkan faktor lain ", False, False),
        ("(Krebs et al., 2022)", False, False),
        (". Dalam penelitian ini, PBL digunakan dua kali: pertama sebagai input perhitungan ", False, False),
        ("Z-Score", False, True),
        (" saat lahir (HAZ awal), dan kedua sebagai variabel ", False, False),
        ("input", False, True),
        (" pada fuzzy ketika menelaah kondisi awal balita.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 5,
        "Forest plot faktor yang memengaruhi skor Z panjang menurut umur (LAZ) pada usia 24 bulan (Krebs et al., 2022).")

    # 2.5.4 Usia Pengenalan MP-ASI
    add_heading_subsection(doc, "2.5.4", "Usia Pengenalan Makanan Pendamping ASI (MP-ASI)")
    add_para(doc, text_segments=[
        ("Waktu pengenalan Makanan Pendamping ASI (MP-ASI) merupakan faktor penting dalam risiko ", False, False),
        ("stunting", False, True),
        (" karena menentukan kecukupan energi dan zat gizi selama transisi dari ASI eksklusif. Pemberian terlalu dini terbukti meningkatkan risiko malnutrisi, ", False, False),
        ("wasting", False, True),
        (", hingga ", False, False),
        ("stunting", False, True),
        (" akibat ketidaksiapan sistem pencernaan bayi dan kemungkinan tergantinya ASI dengan makanan yang kurang padat gizi ", False, False),
        ("(Muflih & Khasana, 2025)", False, False),
        (". Penelitian ", False, False),
        ("Waqiyah et al. (2023)", False, False),
        (" menunjukkan bahwa anak yang menerima MP-ASI tidak tepat waktu memiliki proporsi ", False, False),
        ("stunting", False, True),
        (" lebih tinggi (39,6%) dibandingkan anak dengan MP-ASI tepat waktu (16,0%), dengan hubungan signifikan secara statistik. Pada penelitian ini, MP-ASI diukur sebagai usia pengenalan dalam satuan bulan dan menjadi salah satu variabel ", False, False),
        ("input", False, True),
        (" fuzzy dengan tiga kategori linguistik: ", False, False),
        ("terlalu cepat", False, True),
        (", ", False, False),
        ("tepat", False, True),
        (", dan ", False, False),
        ("terlalu lambat", False, True),
        (".", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 6,
        "Hubungan Waktu Pemberian MP-ASI dan Pengetahuan Ibu tentang Responsive Feeding dengan Kejadian Stunting pada Anak Usia 6-24 Bulan (Waqiyah et al., 2023).")

    # 2.5.5 Tinggi Badan Saat Pengukuran
    add_heading_subsection(doc, "2.5.5", "Tinggi Badan saat Pengukuran")
    add_para(doc, text_segments=[
        ("Tinggi badan saat pengukuran (TB ukur) menggambarkan kondisi pertumbuhan linear balita pada saat survei berlangsung. Variabel ini menjadi indikator utama untuk menghitung HAZ pada saat ukur (HAZ-ukur), yang dalam penelitian ini diperlakukan sebagai ", False, False),
        ("ground truth", False, True),
        (" fungsional bagi evaluasi sistem fuzzy. Pengukuran dilakukan dengan LAZ untuk balita berusia ≤24 bulan dan HAZ untuk balita >24 bulan sebagaimana diterapkan dalam berbagai studi ", False, False),
        ("(Rosdiana et al., 2024)", False, False),
        (". Variabel ini juga menjadi ", False, False),
        ("input", False, True),
        (" fuzzy dengan tiga kategori linguistik (pendek, normal, tinggi) yang fungsi keanggotaannya dibedakan berdasarkan jenis kelamin.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 7,
        "Grafik Ilustrasi Pertumbuhan Panjang/Tinggi Badan Balita Selama 1000 HPK berdasarkan Usia dan Median Tinggi Badan Rujukan WHO (UNICEF, 2023).")

    # 2.5.6 Berat Badan
    add_heading_subsection(doc, "2.5.6", "Berat Badan saat Pengukuran")
    add_para(doc, text_segments=[
        ("Berat badan saat pengukuran merupakan indikator status nutrisi akut sekaligus prediktor kuat risiko ", False, False),
        ("stunting", False, True),
        (" karena mencerminkan kondisi gizi yang baru saja terjadi serta kecukupan energi harian balita ", False, False),
        ("(Rahmadiani et al., 2025)", False, False),
        (". Temuan penelitian di Puskesmas Kintamani VI yang ditunjukkan pada Gambar II.8 memperlihatkan bahwa balita dengan berat badan rendah memiliki proporsi ", False, False),
        ("stunting", False, True),
        (" lebih tinggi, dan balita dengan riwayat berat rendah memiliki risiko ", False, False),
        ("stunting", False, True),
        (" 4,45 kali lebih besar daripada balita normal ", False, False),
        ("(Dwivayana et al., 2025)", False, False),
        (". Dalam penelitian ini, berat badan saat pengukuran menjadi salah satu variabel ", False, False),
        ("input", False, True),
        (" fuzzy dengan empat kategori linguistik (kurang, normal, lebih, obesitas) yang fungsi keanggotaannya juga dibedakan menurut jenis kelamin.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 8,
        "Grafik Pengaruh Stunting berdasarkan Berat Badan pada Balita (Dwivayana et al., 2025).")

    # 2.5.7 Kelas Baru Hasil Turunan
    add_heading_subsection(doc, "2.5.7", "Kelas Baru Hasil Turunan")
    add_para(doc, text_segments=[
        ("Selain enam variabel utama di atas, penelitian ini menghitung dua variabel turunan yang penting dalam kerangka ", False, False),
        ("monitoring", False, True),
        (", yaitu (a) jarak antara usia balita saat lahir dengan saat pengukuran berlangsung (dalam bulan), dan (b) jarak antara usia pengenalan MP-ASI dengan saat pengambilan survei (dalam bulan). Kedua variabel turunan ini bersifat informatif, digunakan untuk konteks ", False, False),
        ("monitoring", False, True),
        (" trajektori pertumbuhan, dan tidak masuk sebagai ", False, False),
        ("input", False, True),
        (" fuzzy agar interpretabilitas aturan tetap sederhana.", False, False),
    ], first_line_indent=1.27)

    # ========================================================================
    # 2.6 LOGIKA FUZZY
    # ========================================================================
    add_heading_section(doc, "2.6", "Logika Fuzzy")

    add_para(doc, text_segments=[
        ("Fuzzy ", False, False),
        ("logic", False, True),
        (" merupakan kerangka matematis untuk penalaran yang memodelkan derajat kebenaran kontinu antara 0-1 melalui teori himpunan samar, sehingga suatu elemen memiliki tingkat keanggotaan alih-alih status biner. Mekanisme dasarnya bekerja dengan aturan komposisional untuk menghasilkan keputusan dari informasi yang tidak pasti, menjadikannya fleksibel dan selaras dengan cara berpikir manusia. Pendekatan ini terbukti luas digunakan dalam berbagai aplikasi, mulai dari kendali proses hingga diagnosis medis, karena sifatnya yang interpretatif dan mudah diintegrasikan dalam sistem ", False, False),
        ("(Borgwardt & Peñaloza, 2017)", False, False),
        (". Dalam penelitian ini, variabel yang difuzzyfikasi adalah usia, tinggi badan saat pengukuran, berat badan saat pengukuran, panjang badan lahir, dan MP-ASI. Persamaan dasar logika fuzzy ditunjukkan pada Persamaan (II.3).", False, False),
    ], first_line_indent=1.27)

    add_equation(doc, "μ_A(x) ∈ [0, 1]", 3)

    add_para(doc, plain_text="dimana:", first_line_indent=0)
    add_bullet_list(doc, [
        [("μ_A(x)", False, True), (" adalah derajat keanggotaan (", False, False), ("membership degree", False, True), (") elemen ", False, False), ("x", False, True), (" terhadap himpunan samar A; nilainya dinormalisasi ke rentang 0-1.", False, False)],
        [("∈ [0, 1]", False, True), (" menunjukkan rentang nilai yang dapat diambil. Bisa tepat 0, tepat 1, atau nilai apa pun di antaranya seperti 0,13, 0,5, atau 0,97.", False, False)],
    ])

    add_para(doc, text_segments=[
        ("Logika fuzzy juga sering disebut sebagai ", False, False),
        ("Fuzzy Inference System", False, True),
        (" (FIS) atau ", False, False),
        ("Fuzzy Inference Engine", False, True),
        (" ", False, False),
        ("(Maulana et al., 2024)", False, False),
        (". Gambar II.9 menunjukkan alur kerja fundamental sistem logika fuzzy yang terdiri atas tahap fuzzyfikasi, evaluasi aturan, dan defuzzifikasi.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 9, "Alur Kerja Sistem Logika Fuzzy (Maulana et al., 2024).")

    # 2.6.1 Fuzzyfikasi
    add_heading_subsection(doc, "2.6.1", "Fuzzyfikasi")
    add_para(doc, text_segments=[
        ("Pada tahap awal alur sistem fuzzy, nilai ", False, False),
        ("input", False, True),
        (" tegas diubah menjadi derajat keanggotaan menggunakan fungsi keanggotaan, sehingga variabel seperti usia, tinggi badan, berat badan, PBL, dan MP-ASI dapat direpresentasikan dalam bentuk linguistik yang fleksibel ", False, False),
        ("(Gupta et al., 2022)", False, False),
        (". Logika fuzzy menggunakan fungsi keanggotaan untuk memetakan setiap nilai ", False, False),
        ("input", False, True),
        (" ke derajat keanggotaan. Dalam penelitian ini dipilih bentuk trapesium dan segitiga. Representasi grafik kedua bentuk tersebut ditunjukkan pada Gambar II.10.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 10,
        "Representasi Grafik Fungsi Keanggotaan: (a) Bentuk Trapesium (Maulana et al., 2024) dan (b) Bentuk Segitiga (Astonis, 2014).")

    # Trapezoid
    add_heading_subsubsection(doc, "2.6.1.1", "Trapesium")
    add_para(doc, text_segments=[
        ("Trapesium dipilih karena sederhana, mudah ditafsirkan, dan stabil dalam menggambarkan batas kategori variabel risiko ", False, False),
        ("stunting", False, True),
        (". Secara matematis, trapesium ditetapkan melalui empat parameter (", False, False),
        ("a, b, c, d", False, True),
        (") yang membentuk pola naik-datar-turun ", False, False),
        ("(Putri & Saputro, 2021)", False, False),
        (". Persamaan fungsi keanggotaan trapesium dapat dilihat pada Persamaan (II.4).", False, False),
    ], first_line_indent=1.27)

    add_equation(doc,
        "μ_trap(x; a, b, c, d) = { 0 untuk x ≤ a;  (x−a)/(b−a) untuk a < x < b;  "
        "1 untuk b ≤ x ≤ c;  (d−x)/(d−c) untuk c < x < d;  0 untuk x ≥ d }", 4)

    add_para(doc, plain_text="dengan ketentuan:", first_line_indent=0)
    add_bullet_list(doc, [
        [("a", False, True), (" = batas kiri dukungan (", False, False), ("left support", False, True), ("); nilai ≤ a dianggap tidak masuk kategori.", False, False)],
        [("b", False, True), (" = titik mulai keanggotaan penuh; x sudah masuk kategori dengan kekuatan 1.", False, False)],
        [("c", False, True), (" = titik akhir keanggotaan penuh; x masih 100% bagian dari kategori.", False, False)],
        [("d", False, True), (" = batas kanan dukungan (", False, False), ("right support", False, True), ("); nilai ≥ d dianggap keluar kategori.", False, False)],
    ])

    # Triangle
    add_heading_subsubsection(doc, "2.6.1.2", "Segitiga")
    add_para(doc, text_segments=[
        ("Selain trapesium, bentuk representatif lain yang digunakan adalah segitiga. Fungsi segitiga adalah salah satu bentuk fuzzyfikasi sederhana yang terdiri dari tiga titik (", False, False),
        ("a, b, c", False, True),
        ("). Dalam konteks ", False, False),
        ("monitoring", False, True),
        (" ", False, False),
        ("stunting", False, True),
        (", fungsi ini cocok untuk memodelkan variabel yang memiliki satu titik optimum ", False, False),
        ("(Abdy, 2018)", False, False),
        (". Persamaan fungsi keanggotaan segitiga dapat dilihat pada Persamaan (II.5).", False, False),
    ], first_line_indent=1.27)

    add_equation(doc,
        "μ_tri(x; a, b, c) = { 0 untuk x ≤ a atau x ≥ c;  (x−a)/(b−a) untuk a ≤ x ≤ b;  "
        "(c−x)/(c−b) untuk b ≤ x ≤ c }", 5)

    add_para(doc, plain_text="dengan ketentuan:", first_line_indent=0)
    add_bullet_list(doc, [
        [("a", False, True), (" = ambang batas kiri dukungan (", False, False), ("left support", False, True), (").", False, False)],
        [("b", False, True), (" = titik puncak; x = b memiliki keanggotaan penuh (μ = 1) dan paling mewakili kategori.", False, False)],
        [("c", False, True), (" = ambang batas kanan dukungan (", False, False), ("right support", False, True), (").", False, False)],
    ])

    # 2.6.2 Aturan Dasar
    add_heading_subsection(doc, "2.6.2", "Aturan Dasar Fuzzy Sugeno")
    add_para(doc, text_segments=[
        ("Metode fuzzy Takagi-Sugeno (TS) memodelkan sistem nonlinier melalui aturan ", False, False),
        ("IF-THEN", False, True),
        (" dengan konsekuen berbentuk konstanta tegas atau persamaan linear, sehingga mampu menangani ketidakpastian sekaligus menjaga struktur matematis yang efisien. Pendekatan ini relevan untuk kebutuhan kendali dan pengambilan keputusan modern ", False, False),
        ("(Yershov & Ponomarenko, 2018)", False, False),
        (". Bentuk umum aturan Sugeno orde-0 dapat dilihat pada Persamaan (II.6).", False, False),
    ], first_line_indent=1.27)

    add_equation(doc,
        "IF x₁ adalah A₁ᵢ AND x₂ adalah A₂ᵢ AND … AND xₙ adalah Aₙᵢ THEN y = cᵢ", 6)

    add_para(doc, plain_text="dimana:", first_line_indent=0)
    add_bullet_list(doc, [
        [("x₁, x₂, …, xₙ", False, True), (" = variabel masukan (", False, False), ("input", False, True), (") sistem fuzzy.", False, False)],
        [("A₁ᵢ, A₂ᵢ, …, Aₙᵢ", False, True), (" = himpunan fuzzy (label linguistik) yang melekat pada masing-masing variabel masukan pada aturan ke-i.", False, False)],
        [("Operator AND", False, True), (" = operator logika fuzzy yang direalisasikan dengan ", False, False), ("t-norm", False, True), (", umumnya operator ", False, False), ("minimum", False, True), (".", False, False)],
        [("y", False, True), (" = variabel keluaran (", False, False), ("output", False, True), (") sistem fuzzy, dalam penelitian ini berupa ", False, False), ("risk score", False, True), (" stunting yang dipetakan ke ", False, False), ("priority flag", False, True), (".", False, False)],
        [("cᵢ", False, True), (" = nilai konstan (konsekuen tegas) pada aturan Sugeno ke-i.", False, False)],
    ])

    add_para(doc, text_segments=[
        ("Pada penelitian ini, basis aturan dibangun dengan strategi ", False, False),
        ("pre-filter", False, True),
        (" oleh jenis kelamin diikuti dengan aturan komposit antara lima variabel ", False, False),
        ("input", False, True),
        (" fuzzy (usia, tinggi badan, berat badan, PBL, dan MP-ASI). Untuk setiap jenis kelamin, terdapat 108 aturan ", False, False),
        ("IF-THEN", False, True),
        (" yang dipetakan ke empat kelas risiko: ", False, False),
        ("STUNTING HIGH", False, True),
        (" (10 aturan, c = 90), ", False, False),
        ("STUNTING MEDIUM", False, True),
        (" (30 aturan, c = 65), ", False, False),
        ("STUNTING LOW", False, True),
        (" (20 aturan, c = 35), dan ", False, False),
        ("NORMAL", False, True),
        (" (48 aturan, c = 10). Karena terdapat dua jalur jenis kelamin, secara total terdapat 108 × 2 = 216 aturan efektif yang dievaluasi sistem.", False, False),
    ], first_line_indent=1.27)

    # 2.6.3 Defuzzifikasi
    add_heading_subsection(doc, "2.6.3", "Defuzzifikasi")
    add_para(doc, text_segments=[
        ("Tahap terakhir alur sistem fuzzy mengubah hasil agregasi aturan menjadi nilai tegas. Pada metode Sugeno digunakan rata-rata berbobot (", False, False),
        ("weighted average", False, True),
        ("), bukan ", False, False),
        ("centroid", False, True),
        (" seperti pada Mamdani, karena konsekuen Sugeno berbentuk konstanta tegas ", False, False),
        ("(Georgescu et al., 2014)", False, False),
        (". Persamaan defuzzifikasi Sugeno ditunjukkan pada Persamaan (II.7).", False, False),
    ], first_line_indent=1.27)

    add_equation(doc, "y = ( Σᵢ wᵢ · cᵢ ) / ( Σᵢ wᵢ )", 7)

    add_para(doc, plain_text="dimana:", first_line_indent=0)
    add_bullet_list(doc, [
        [("y", False, True), (" = keluaran akhir sistem inferensi fuzzy Sugeno, berupa satu nilai tegas (", False, False), ("crisp", False, True), (") hasil penggabungan semua aturan.", False, False)],
        [("i", False, True), (" = indeks aturan ke-i pada basis aturan (", False, False), ("rule base", False, True), (").", False, False)],
        [("cᵢ", False, True), (" = nilai konsekuen konstan pada aturan ke-i.", False, False)],
        [("wᵢ", False, True), (" = derajat aktivasi (bobot) aturan ke-i, diperoleh dari penggabungan derajat keanggotaan premis melalui operator ", False, False), ("AND", False, True), (" (", False, False), ("minimum", False, True), (").", False, False)],
        [("Σᵢ wᵢ · cᵢ", False, True), (" = penjumlahan semua produk antara derajat aktivasi dan nilai konsekuen.", False, False)],
        [("Σᵢ wᵢ", False, True), (" = penjumlahan seluruh derajat aktivasi aturan.", False, False)],
    ])

    add_para(doc, text_segments=[
        ("Hasil defuzzifikasi pada penelitian ini berupa ", False, False),
        ("risk score", False, True),
        (" dengan rentang 0-100 yang kemudian dipetakan ke ", False, False),
        ("priority flag", False, True),
        (" empat tingkat: ", False, False),
        ("Normal", False, True),
        (" (0-22), ", False, False),
        ("Stunting Low", False, True),
        (" (>22-50), ", False, False),
        ("Stunting Medium", False, True),
        (" (>50-77), dan ", False, False),
        ("Stunting High", False, True),
        (" (>77-100). Pemetaan ini sejalan dengan filosofi nilai konsekuen pada masing-masing kelas risiko.", False, False),
    ], first_line_indent=1.27)

    # ========================================================================
    # 2.7 ALGORITMA GENETIKA (GA)
    # ========================================================================
    add_heading_section(doc, "2.7", "Algoritma Genetika")

    add_para(doc, text_segments=[
        ("Algoritma Genetika (", False, False),
        ("Genetic Algorithm", False, True),
        (", GA) merupakan metode optimasi metaheuristik yang terinspirasi dari proses evolusi biologis Darwin, yaitu seleksi alam, persilangan (", False, False),
        ("crossover", False, True),
        ("), dan mutasi ", False, False),
        ("(Katoch et al., 2021)", False, False),
        (". GA bekerja dengan memelihara populasi individu (kandidat solusi) yang direpresentasikan sebagai kromosom, kemudian secara iteratif mengevolusikan populasi tersebut menuju solusi yang lebih baik dengan dipandu oleh fungsi ", False, False),
        ("fitness", False, True),
        (". Dalam konteks optimasi sistem fuzzy, GA telah terbukti efektif untuk menyesuaikan parameter fungsi keanggotaan terhadap karakteristik empiris data tanpa mengubah struktur aturan ", False, False),
        ("(Suresh et al., 2023; Cordón, 2011)", False, False),
        (". Alur GA secara umum ditunjukkan pada Gambar II.11.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 11,
        "Alur Kerja Algoritma Genetika (Katoch et al., 2021).")

    add_para(doc, text_segments=[
        ("Komponen utama GA meliputi representasi kromosom, fungsi ", False, False),
        ("fitness", False, True),
        (", seleksi, ", False, False),
        ("crossover", False, True),
        (", mutasi, dan ", False, False),
        ("elitism", False, True),
        (". Penjelasan masing-masing komponen diuraikan pada sub-bab berikut.", False, False),
    ], first_line_indent=1.27)

    # 2.7.1 Representasi Kromosom
    add_heading_subsection(doc, "2.7.1", "Representasi Kromosom")
    add_para(doc, text_segments=[
        ("Kromosom merupakan vektor berisi gen-gen yang merepresentasikan satu kandidat solusi. Dalam penelitian ini, kromosom diisi oleh nilai titik batas (", False, False),
        ("boundary points", False, True),
        (") seluruh fungsi keanggotaan ", False, False),
        ("input", False, True),
        (" fuzzy. Dengan demikian, satu individu = satu set parameter lengkap fungsi keanggotaan. Representasi ini bersifat ", False, False),
        ("real-coded", False, True),
        (" (bilangan real), bukan biner, karena nilai ", False, False),
        ("boundary points", False, True),
        (" bersifat kontinu dan presisi rendah pada representasi biner akan mengurangi ruang pencarian GA secara signifikan.", False, False),
    ], first_line_indent=1.27)

    # 2.7.2 Fungsi Fitness
    add_heading_subsection(doc, "2.7.2", "Fungsi Fitness")
    add_para(doc, text_segments=[
        ("Fungsi ", False, False),
        ("fitness", False, True),
        (" merupakan metrik evaluasi yang memandu arah evolusi populasi. Kromosom dengan ", False, False),
        ("fitness", False, True),
        (" lebih tinggi memiliki probabilitas lebih besar untuk terpilih sebagai induk pada generasi berikutnya. Dalam penelitian ini, fungsi ", False, False),
        ("fitness", False, True),
        (" memaksimalkan F1-Score sistem fuzzy terhadap label ", False, False),
        ("ground truth", False, True),
        (" (HAZ saat ukur), karena F1-Score merupakan harmonik antara ", False, False),
        ("precision", False, True),
        (" dan ", False, False),
        ("recall", False, True),
        (" yang seimbang dan tahan terhadap ketidakseimbangan kelas, sebagaimana ditegaskan oleh ", False, False),
        ("Hand et al. (2021)", False, False),
        (".", False, False),
    ], first_line_indent=1.27)

    # 2.7.3 Seleksi
    add_heading_subsection(doc, "2.7.3", "Seleksi")
    add_para(doc, text_segments=[
        ("Seleksi adalah proses memilih individu dari populasi untuk dijadikan induk. Metode seleksi yang umum adalah ", False, False),
        ("tournament selection", False, True),
        (" yang memilih sejumlah ", False, False),
        ("k", False, True),
        (" individu secara acak lalu mengambil yang ber-", False, False),
        ("fitness", False, True),
        (" tertinggi sebagai induk. Pendekatan ini lebih stabil dibandingkan ", False, False),
        ("roulette wheel", False, True),
        (" pada permasalahan dengan rentang ", False, False),
        ("fitness", False, True),
        (" sempit, karena tidak terlalu sensitif terhadap perbedaan ", False, False),
        ("fitness", False, True),
        (" absolut antar individu ", False, False),
        ("(Goldberg & Deb, 1991)", False, False),
        (". Penelitian ini menggunakan ", False, False),
        ("tournament selection", False, True),
        (" dengan ukuran ", False, False),
        ("k = 3", False, True),
        (".", False, False),
    ], first_line_indent=1.27)

    # 2.7.4 Crossover
    add_heading_subsection(doc, "2.7.4", "Crossover")
    add_para(doc, text_segments=[
        ("Crossover", False, True),
        (" atau persilangan adalah proses kombinasi gen dari dua induk untuk menghasilkan keturunan. Untuk representasi ", False, False),
        ("real-coded", False, True),
        (" digunakan ", False, False),
        ("blend crossover", False, True),
        (" (BLX-α), di mana nilai gen anak berada di sekitar rentang gen induknya diperluas dengan faktor α. Operator ini mempertahankan eksplorasi ruang pencarian tanpa kehilangan informasi penting dari induk ", False, False),
        ("(Herrera et al., 2003)", False, False),
        (". Persamaan ", False, False),
        ("blend crossover", False, True),
        (" ditunjukkan pada Persamaan (II.8).", False, False),
    ], first_line_indent=1.27)

    add_equation(doc,
        "x_c = x_p1 + r · (x_p2 − x_p1),    r ~ U(−α, 1 + α)", 8)

    add_para(doc, plain_text="dimana:", first_line_indent=0)
    add_bullet_list(doc, [
        [("x_c", False, True), (" = nilai gen pada keturunan.", False, False)],
        [("x_p1, x_p2", False, True), (" = nilai gen pada kedua induk.", False, False)],
        [("r", False, True), (" = variabel acak ", False, False), ("uniform", False, True), (" dalam rentang (−α, 1 + α).", False, False)],
        [("α", False, True), (" = parameter ekspansi (umumnya 0,3 hingga 0,5).", False, False)],
    ])

    # 2.7.5 Mutasi
    add_heading_subsection(doc, "2.7.5", "Mutasi")
    add_para(doc, text_segments=[
        ("Mutasi merupakan operator yang memperkenalkan keragaman baru pada populasi dengan cara mengubah nilai gen secara acak. Pada representasi ", False, False),
        ("real-coded", False, True),
        (", mutasi ", False, False),
        ("Gaussian", False, True),
        (" sering digunakan, di mana gen yang termutasi diubah dengan menambahkan derau acak berdistribusi normal yang dibatasi oleh ", False, False),
        ("boundary", False, True),
        (" parameter yang valid. Persamaan mutasi ", False, False),
        ("Gaussian", False, True),
        (" ditunjukkan pada Persamaan (II.9).", False, False),
    ], first_line_indent=1.27)

    add_equation(doc, "x'_i = clip( x_i + N(0, σ²) , lb_i, ub_i )", 9)

    add_para(doc, plain_text="dimana:", first_line_indent=0)
    add_bullet_list(doc, [
        [("x'_i", False, True), (" = nilai gen setelah mutasi.", False, False)],
        [("x_i", False, True), (" = nilai gen sebelum mutasi.", False, False)],
        [("N(0, σ²)", False, True), (" = bilangan acak berdistribusi normal dengan rerata 0 dan variansi σ².", False, False)],
        [("lb_i, ub_i", False, True), (" = batas bawah dan batas atas valid untuk gen ke-i.", False, False)],
        [("clip", False, True), (" = operator pemotongan agar gen tetap berada dalam rentang valid.", False, False)],
    ])

    add_para(doc, text_segments=[
        ("Mutasi terjadi dengan probabilitas tertentu per gen (", False, False),
        ("indpb", False, True),
        ("). Probabilitas yang terlalu tinggi membuat GA cenderung berperilaku acak; probabilitas yang terlalu rendah membuat GA terjebak di optimum lokal. Penelitian ini menggunakan ", False, False),
        ("indpb", False, True),
        (" = 0,15 dengan σ proporsional terhadap rentang gen (15% dari rentang).", False, False),
    ], first_line_indent=1.27)

    # 2.7.6 Elitism
    add_heading_subsection(doc, "2.7.6", "Elitism dan Strategi Multi-Seed")
    add_para(doc, text_segments=[
        ("Elitism", False, True),
        (" adalah mekanisme untuk menjamin bahwa individu terbaik dari setiap generasi diwariskan ke generasi berikutnya tanpa melalui ", False, False),
        ("crossover", False, True),
        (" atau mutasi. Strategi ini mencegah hilangnya solusi terbaik karena operator stokastik. Selain itu, untuk meningkatkan robustness terhadap optimum lokal, penelitian ini menerapkan strategi ", False, False),
        ("multi-seed run", False, True),
        (", yaitu menjalankan GA beberapa kali dengan ", False, False),
        ("random seed", False, True),
        (" berbeda lalu memilih individu terbaik dari seluruh ", False, False),
        ("run", False, True),
        (". Pendekatan ini sejalan dengan praktik di literatur GA modern yang menunjukkan bahwa hasil ", False, False),
        ("multi-seed", False, True),
        (" lebih representatif terhadap performa sebenarnya algoritma daripada ", False, False),
        ("single-run", False, True),
        (" ", False, False),
        ("(Cordón, 2011)", False, False),
        (".", False, False),
    ], first_line_indent=1.27)

    # ========================================================================
    # 2.8 MONITORING STUNTING BERBASIS DUAL Z-SCORE
    # ========================================================================
    add_heading_section(doc, "2.8", "Monitoring Stunting Berbasis Dual Z-Score")

    add_para(doc, text_segments=[
        ("Berbeda dengan klasifikasi statis yang hanya memberi label ", False, False),
        ("stunting", False, True),
        (" atau normal berdasarkan satu pengukuran, ", False, False),
        ("monitoring", False, True),
        (" ", False, False),
        ("stunting", False, True),
        (" bertujuan memantau perubahan kondisi balita dari waktu ke waktu. Penelitian ini mengusung kerangka ", False, False),
        ("dual Z-Score", False, True),
        (" yang menghitung HAZ pada dua titik waktu: saat lahir (HAZ awal) dan saat pengukuran survei (HAZ ukur). Kerangka ini memungkinkan empat pola interpretasi seperti ditunjukkan pada Tabel II.3.", False, False),
    ], first_line_indent=1.27)

    add_table_caption(doc, 3, "Empat Pola Interpretasi Monitoring Stunting Dual Z-Score")
    headers3 = [
        [("HAZ Saat Lahir", True, False)],
        [("HAZ Saat Ukur (", True, False), ("Fuzzy Output", True, True), (")", True, False)],
        [("Interpretasi", True, False)],
    ]
    rows3 = [
        ["Normal", "Normal", "Trajektori pertumbuhan baik; balita konsisten normal."],
        ["Normal", "Stunting", "Trajektori memburuk; balita lahir normal namun mengalami penurunan pertumbuhan (alarm intervensi)."],
        ["Stunting", "Normal", "Trajektori membaik; balita lahir pendek namun berkembang baik (perlu monitoring lanjutan)."],
        ["Stunting", "Stunting", "Trajektori buruk; balita konsisten stunting (prioritas intervensi tinggi)."],
    ]
    add_table_from_data(doc, headers3, rows3, col_widths_cm=[3.5, 4.0, 7.0])

    add_para(doc, text_segments=[
        ("Pola ini menunjukkan bahwa ", False, False),
        ("monitoring", False, True),
        (" bukanlah perpanjangan klasifikasi melainkan kerangka berbeda yang menempatkan trajektori pertumbuhan sebagai fokus utama. Dalam evaluasi sistem, HAZ saat ukur digunakan sebagai ", False, False),
        ("ground truth", False, True),
        (" untuk menilai apakah ", False, False),
        ("risk score", False, True),
        (" fuzzy mampu menangkap kondisi balita yang sebenarnya pada saat pengukuran, sehingga sistem dievaluasi secara konsisten dengan tujuan ", False, False),
        ("monitoring", False, True),
        (".", False, False),
    ], first_line_indent=1.27)

    # ========================================================================
    # 2.9 STRATIFIED HOLD-OUT VALIDATION
    # ========================================================================
    add_heading_section(doc, "2.9", "Stratified Hold-Out Validation")

    add_para(doc, text_segments=[
        ("Stratified hold-out validation", False, True),
        (" membagi data ke dalam set pelatihan dan pengujian dengan menjaga proporsi kelas label tetap seimbang di kedua subset, sehingga evaluasi tidak bias akibat ketidakseimbangan kelas yang umum terjadi pada data kesehatan ", False, False),
        ("(Abegaz et al., 2025)", False, False),
        (". Pendekatan ini banyak digunakan pada studi klinis dan klasifikasi modern karena memberikan estimasi kinerja yang lebih representatif dibanding pembagian acak biasa, terutama ketika ", False, False),
        ("outcome", False, True),
        (" seperti ", False, False),
        ("stunting", False, True),
        (" (berdasarkan HAZ) tidak terdistribusi merata ", False, False),
        ("(Andrews et al., 2025)", False, False),
        (". Penelitian ini menggunakan rasio 80:20 dengan stratifikasi pada label ", False, False),
        ("Stunting/Normal", False, True),
        (" hasil HAZ saat ukur sebagai dasar penyusunan ", False, False),
        ("training set", False, True),
        (" dan ", False, False),
        ("test set", False, True),
        (".", False, False),
    ], first_line_indent=1.27)

    # ========================================================================
    # 2.10 CONFUSION MATRIX
    # ========================================================================
    add_heading_section(doc, "2.10", "Confusion Matrix dan Metrik Evaluasi")

    add_para(doc, text_segments=[
        ("Confusion matrix", False, True),
        (" merupakan alat evaluasi dasar yang merangkum hubungan antara prediksi sistem dan kondisi aktual sehingga menghasilkan metrik penting seperti akurasi, ", False, False),
        ("precision", False, True),
        (", ", False, False),
        ("recall", False, True),
        (", dan F1-Score ", False, False),
        ("(Japkowicz & Shah, 2015)", False, False),
        (". Gambar II.12 menampilkan empat komponen utama, yaitu TP, FP, FN, dan TN yang menjadi fondasi seluruh perhitungan metrik.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 12, "Struktur Confusion Matrix (Liu & Yang, 2022).")

    add_para(doc, text_segments=[
        ("Precision", False, True),
        (" mengukur ketepatan prediksi positif, yaitu proporsi prediksi ", False, False),
        ("stunting", False, True),
        (" yang benar-benar ", False, False),
        ("stunting", False, True),
        (", sedangkan ", False, False),
        ("recall", False, True),
        (" mengukur kemampuan model menangkap seluruh kasus positif ", False, False),
        ("(Landgrebe et al., 2006)", False, False),
        (". Persamaan ", False, False),
        ("precision", False, True),
        (" dan ", False, False),
        ("recall", False, True),
        (" ditunjukkan pada Persamaan (II.10) dan Persamaan (II.11).", False, False),
    ], first_line_indent=1.27)

    add_equation(doc, "Precision = TP / ( TP + FP )", 10)
    add_equation(doc, "Recall = TP / ( TP + FN )", 11)

    add_para(doc, text_segments=[
        ("F1-Score merupakan metrik penting karena menyeimbangkan ", False, False),
        ("precision", False, True),
        (" dan ", False, False),
        ("recall", False, True),
        (" melalui rata-rata harmonik, sehingga lebih peka terhadap ", False, False),
        ("false positive", False, True),
        (" dan ", False, False),
        ("false negative", False, True),
        (", terutama pada data tidak seimbang ", False, False),
        ("(Hand et al., 2021)", False, False),
        (". Persamaan F1-Score ditunjukkan pada Persamaan (II.12).", False, False),
    ], first_line_indent=1.27)

    add_equation(doc, "F1-Score = 2 × ( Precision × Recall ) / ( Precision + Recall )", 12)

    add_para(doc, text_segments=[
        ("Akurasi menunjukkan berapa banyak prediksi yang benar dari seluruh prediksi yang dibuat oleh model penelitian, dalam hal ini sistem ", False, False),
        ("fuzzy logic", False, True),
        (" yang dioptimasi GA ", False, False),
        ("(Marsya Finda & Wahyu Utomo, 2024)", False, False),
        (". Persamaan akurasi ditunjukkan pada Persamaan (II.13).", False, False),
    ], first_line_indent=1.27)

    add_equation(doc, "Accuracy = ( TP + TN ) / ( TP + TN + FP + FN )", 13)

    add_para(doc, plain_text="Keterangan:", first_line_indent=0)
    add_bullet_list(doc, [
        [("TP", True, False), (" (", False, False), ("True Positive", False, True), (") = balita yang benar diprediksi ", False, False), ("stunting", False, True), (" dan memang berstatus ", False, False), ("stunting", False, True), (" menurut HAZ ukur.", False, False)],
        [("TN", True, False), (" (", False, False), ("True Negative", False, True), (") = balita yang benar diprediksi normal dan memang normal menurut HAZ ukur.", False, False)],
        [("FP", True, False), (" (", False, False), ("False Positive", False, True), (") = balita diprediksi ", False, False), ("stunting", False, True), (" padahal HAZ ukur normal.", False, False)],
        [("FN", True, False), (" (", False, False), ("False Negative", False, True), (") = balita diprediksi normal padahal HAZ ukur ", False, False), ("stunting", False, True), (".", False, False)],
    ])

    add_para(doc, text_segments=[
        ("Untuk mengevaluasi konsistensi peringkat ", False, False),
        ("risk score", False, True),
        (" terhadap HAZ, penelitian ini juga melaporkan korelasi Spearman antara ", False, False),
        ("risk score", False, True),
        (" dan nilai HAZ ukur. Korelasi Spearman dipilih karena tidak mengasumsikan distribusi normal dan tahan terhadap ", False, False),
        ("outlier", False, True),
        (", sehingga cocok untuk data antropometri balita yang kerap miring (", False, False),
        ("skewed", False, True),
        (") ", False, False),
        ("(Schober et al., 2018)", False, False),
        (".", False, False),
    ], first_line_indent=1.27)

    # ========================================================================
    # 2.11 ALASAN PEMILIHAN METODE
    # ========================================================================
    add_heading_section(doc, "2.11", "Alasan Pemilihan Metode")

    add_para(doc, text_segments=[
        ("Pemilihan topik ", False, False),
        ("stunting", False, True),
        (" balita di Kalimantan Tengah didasari beban masalah yang masih tinggi dan penurunan yang berjalan lebih lambat dibanding target nasional. Berdasarkan data Kemenkes, prevalensi ", False, False),
        ("stunting", False, True),
        (" di Kalimantan Tengah sekitar 26,9% pada 2022 dan 23,5% pada 2023, sebagaimana ditunjukkan pada Gambar II.13, dalam konteks kesenjangan antardaerah yang lebar ", False, False),
        ("(Sekretariat WAPRES RI & Kemenko PMK, n.d.)", False, False),
        (". Kondisi ini menuntut pendekatan pemodelan risiko yang lebih presisi untuk mendukung penajaman sasaran intervensi, bukan sekadar pemantauan prevalensi agregat.", False, False),
    ], first_line_indent=1.27)

    add_figure_caption(doc, 13, "Prevalensi Stunting Berdasarkan Provinsi Tahun 2022 dan 2023 (Kemenkes, 2024).")

    add_para(doc, text_segments=[
        ("Meskipun beberapa provinsi di Indonesia menunjukkan prevalensi ", False, False),
        ("stunting", False, True),
        (" tertinggi secara nasional, penelitian ini tidak bertujuan membandingkan besaran prevalensi antarwilayah melainkan mengembangkan dan mengevaluasi sistem ", False, False),
        ("monitoring", False, True),
        (" berbasis data. Kalimantan Tengah dipilih karena berada pada zona transisi epidemiologis ", False, False),
        ("stunting", False, True),
        (" dengan prevalensi di atas rata-rata nasional namun tetap menunjukkan heterogenitas kondisi anak yang luas. Karakteristik ini memungkinkan sistem fuzzy Takagi-Sugeno mengaktifkan aturan secara parsial, menangkap transisi risiko secara halus, dan menghasilkan ", False, False),
        ("risk score", False, True),
        (" yang lebih informatif, stabil, dan valid secara metodologis dibandingkan wilayah dengan kondisi ekstrem.", False, False),
    ], first_line_indent=1.27)

    add_para(doc, text_segments=[
        ("Penelitian ini memandang ", False, False),
        ("stunting", False, True),
        (" sebagai fenomena yang perlu dimonitor dari sudut pandang ganda: kondisi saat lahir (HAZ awal) sebagai titik berangkat trajektori, dan kondisi saat pengukuran (HAZ ukur) sebagai potret terkini yang menjadi ", False, False),
        ("ground truth", False, True),
        (" sistem. Data SKI 2023 digunakan dengan sampel balita di Kalimantan Tengah yang memiliki enam variabel lengkap (usia, jenis kelamin, PBL, MP-ASI, tinggi badan ukur, dan berat badan ukur). Seluruh kabupaten dan kota di Kalimantan Tengah dimasukkan dalam populasi penelitian karena data SKI 2023 bersifat representatif provinsi. Hal ini memungkinkan model fuzzy Sugeno mempelajari variasi risiko ", False, False),
        ("stunting", False, True),
        (" dari wilayah perkotaan, pesisir, hingga pedalaman, sehingga ", False, False),
        ("risk score", False, True),
        (" yang dihasilkan benar-benar mencerminkan heterogenitas Kalimantan Tengah.", False, False),
    ], first_line_indent=1.27)

    add_para(doc, text_segments=[
        ("Metode fuzzy Takagi-Sugeno dipilih karena memiliki efisiensi komputasi tinggi, mampu merepresentasikan ketidakpastian, serta menghasilkan inferensi yang transparan dan mudah ditelusuri. Sistem ini menggunakan kombinasi fungsi keanggotaan trapesium dan segitiga untuk memodelkan transisi gradual antar status gizi, yang secara biologis mencerminkan perubahan kontinu pada pertumbuhan anak. Strategi ", False, False),
        ("pre-filter", False, True),
        (" oleh jenis kelamin dipilih karena standar antropometri WHO berbeda untuk laki-laki dan perempuan, sehingga pemisahan jalur evaluasi menjamin interpretasi linguistik yang konsisten dengan biologi pertumbuhan masing-masing jenis kelamin. Dibandingkan dengan model kecerdasan buatan berbasis pembelajaran mesin yang bersifat ", False, False),
        ("black-box", False, True),
        (", fuzzy Sugeno memungkinkan setiap keputusan ", False, False),
        ("risk score", False, True),
        (" ditelusuri kembali ke aturan, fungsi keanggotaan, dan variabel biologis yang digunakan.", False, False),
    ], first_line_indent=1.27)

    add_para(doc, text_segments=[
        ("Pemilihan algoritma genetika sebagai metode optimasi didasari oleh empat pertimbangan utama. ", False, False),
        ("Pertama", False, True),
        (", ruang pencarian ", False, False),
        ("boundary points", False, True),
        (" fungsi keanggotaan bersifat kontinu, sehingga GA yang ", False, False),
        ("real-coded", False, True),
        (" lebih sesuai dibandingkan ", False, False),
        ("grid search", False, True),
        (" yang diskrit dan akan kehilangan presisi. ", False, False),
        ("Kedua", False, True),
        (", GA mampu menghindari optimum lokal melalui kombinasi ", False, False),
        ("crossover", False, True),
        (", mutasi, dan strategi ", False, False),
        ("multi-seed", False, True),
        (". ", False, False),
        ("Ketiga", False, True),
        (", GA hanya menyesuaikan ", False, False),
        ("boundary points", False, True),
        (" tanpa mengubah jumlah aturan, jumlah label linguistik, maupun nilai konsekuen, sehingga interpretabilitas pakar terhadap sistem fuzzy tetap utuh. ", False, False),
        ("Keempat", False, True),
        (", GA telah terbukti efektif dalam meningkatkan F1-Score sistem fuzzy pada literatur kesehatan multivariat ", False, False),
        ("(Suresh et al., 2023)", False, False),
        (", termasuk pada klasifikasi gizi balita ", False, False),
        ("(Cordón, 2011)", False, False),
        (".", False, False),
    ], first_line_indent=1.27)

    add_para(doc, text_segments=[
        ("Evaluasi kinerja sistem dilakukan menggunakan ", False, False),
        ("confusion matrix", False, True),
        (" dengan metrik F1-Score, ", False, False),
        ("precision", False, True),
        (", ", False, False),
        ("recall", False, True),
        (", dan akurasi, dilengkapi dengan korelasi Spearman antara ", False, False),
        ("risk score", False, True),
        (" dan HAZ ukur sebagai indikator konsistensi peringkat. Kombinasi keempat metode ini, yaitu fuzzy Sugeno, ", False, False),
        ("pre-filter", False, True),
        (" jenis kelamin, algoritma genetika, dan ", False, False),
        ("monitoring", False, True),
        (" berbasis ", False, False),
        ("dual Z-Score", False, True),
        (", menghasilkan sistem yang tidak hanya akurat secara statistik tetapi juga relevan secara operasional bagi perumusan kebijakan daerah dan dapat dijelaskan secara biologis maupun klinis, sehingga sesuai dengan karakteristik data SKI 2023 dan kebutuhan pengambilan keputusan dalam kesehatan masyarakat.", False, False),
    ], first_line_indent=1.27)


# ============================================================================
# MAIN
# ============================================================================
def main():
    doc = setup_document()
    build_bab2(doc)
    output_path = "/Users/sinitygs/Projects2/train/BAB_II_MODIFIKASI_TERBARU.docx"
    doc.save(output_path)
    print(f"Berhasil menyimpan: {output_path}")

if __name__ == "__main__":
    main()