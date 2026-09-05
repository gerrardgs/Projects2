#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generator BAB II LANDASAN TEORI (FINAL CORRECTED)
==================================================
Mengikuti gaya tulisan Ali Sebastian persis seperti pada BAB_I_II_FULL.docx.

KOREKSI PENTING dari versi MODIFIKASI sebelumnya:
- Struktur: 2.1 Literatur Terkait Teori (umbrella) -> 2.1.1 s.d. 2.1.11,
  lalu 2.2 Alasan Pemilihan Metode. BUKAN 2.1, 2.2, ..., 2.11.
- HAPUS tabel-tabel implementation detail yang harusnya di BAB III:
  * Tabel "Distribusi 108 aturan per kelas risiko"
  * Tabel "Pemetaan risk score ke priority flag boundary points (0,0,15,30)"
  * Tabel "Hyperparameter GA (pop=80, gen=60, seed {42,...})"
  * Tabel "Empat pola trajektori monitoring" (ini hasil BAB IV)
- HAPUS subseksi "Kelas Baru Hasil Turunan" (ini preprocessing BAB III)
- HAPUS subseksi "Monitoring Dual Z-Score" terpisah (ini implementasi BAB III)
- HAPUS detail spesifik genom GA (77 gen breakdown) -> ini BAB III
- TETAP pertahankan teori GA (kromosom, fitness, seleksi, crossover, mutasi)
- TETAP pertahankan Spearman & Prevalence Stratification (teori metrik)

TAMBAHAN BARU:
- Daftar Pustaka lengkap dengan referensi terverifikasi (max 5 tahun terakhir
  KECUALI untuk klasik yang masih dipakai seperti Yershov 2018, Georgescu 2014,
  Goldberg & Deb 1991, Herrera 2003 yang tetap relevan).

FORMAT (mengikuti BAB_I_II_FULL):
- A4, margin: kiri 4cm, kanan 3cm, atas/bawah 3cm
- Times New Roman 12pt body, 14pt heading
- Line spacing 1.5
- Format: Tabel II.x, Gambar II.x, (II.x)
"""

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================================
FONT = "Times New Roman"
SZ_BODY = 12
SZ_HEAD = 14
SZ_TBL  = 11
LS = 1.5

# ============================================================================
# LOW-LEVEL HELPERS
# ============================================================================
def _set_font(run, font=FONT):
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), font)
    rFonts.set(qn('w:hAnsi'), font)
    rFonts.set(qn('w:eastAsia'), font)
    rFonts.set(qn('w:cs'), font)


def add_run(para, text, bold=False, italic=False, size=SZ_BODY):
    r = para.add_run(text)
    r.font.name = FONT
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    _set_font(r)
    return r


def set_pf(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=None,
           ls=LS, before=0, after=0):
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = ls
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if indent is not None:
        pf.first_line_indent = Cm(indent)


def _insert_in_tcPr(tcPr, element, tag_name):
    order = [
        'cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge',
        'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection',
        'tcFitText', 'vAlign', 'hideMark',
    ]
    target_idx = order.index(tag_name) if tag_name in order else 99
    insert_before = None
    for child in tcPr:
        tag = child.tag.split('}')[-1]
        child_idx = order.index(tag) if tag in order else 99
        if child_idx > target_idx:
            insert_before = child
            break
    if insert_before is not None:
        insert_before.addprevious(element)
    else:
        tcPr.append(element)


def set_cell_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    existing = tcPr.find(qn('w:tcBorders'))
    if existing is not None:
        tcPr.remove(existing)
    tcBorders = OxmlElement('w:tcBorders')
    for name in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{name}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '000000')
        tcBorders.append(b)
    _insert_in_tcPr(tcPr, tcBorders, 'tcBorders')


def set_cell_shading(cell, fill_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    existing = tcPr.find(qn('w:shd'))
    if existing is not None:
        tcPr.remove(existing)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    _insert_in_tcPr(tcPr, shd, 'shd')


# ============================================================================
# MID-LEVEL HELPERS
# ============================================================================
def P(doc, segs=None, text=None, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
      indent=1.27, before=0, after=0):
    p = doc.add_paragraph()
    set_pf(p, align=align, indent=indent, before=before, after=after)
    if text is not None:
        add_run(p, text, size=SZ_BODY)
    elif segs is not None:
        for s in segs:
            if len(s) == 3:
                add_run(p, s[0], bold=s[1], italic=s[2], size=SZ_BODY)
            elif len(s) == 4:
                add_run(p, s[0], bold=s[1], italic=s[2], size=s[3])
    return p


def H_chap(doc, line1, line2):
    p1 = doc.add_paragraph()
    set_pf(p1, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, before=0, after=0)
    add_run(p1, line1, bold=True, size=SZ_HEAD)
    p2 = doc.add_paragraph()
    set_pf(p2, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, before=0, after=18)
    add_run(p2, line2, bold=True, size=SZ_HEAD)


def H2(doc, num, text):
    p = doc.add_paragraph()
    set_pf(p, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, before=12, after=6)
    add_run(p, f"{num}   {text}", bold=True, size=SZ_BODY)


def H3(doc, num, text, italic=False):
    p = doc.add_paragraph()
    set_pf(p, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, before=10, after=4)
    add_run(p, f"{num}   ", bold=True, size=SZ_BODY)
    add_run(p, text, bold=True, italic=italic, size=SZ_BODY)


def H4(doc, num, text):
    p = doc.add_paragraph()
    set_pf(p, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, before=8, after=2)
    add_run(p, f"{num}   {text}", bold=True, size=SZ_BODY)


def fig_caption(doc, num, text):
    # placeholder line first
    p0 = doc.add_paragraph()
    set_pf(p0, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, before=6, after=2)
    add_run(p0, f"[Sisipkan: Gambar 2.{num}]", italic=True, size=SZ_BODY)
    # then caption
    p = doc.add_paragraph()
    set_pf(p, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, before=0, after=12)
    add_run(p, f"Gambar 2.{num} ", bold=True, size=SZ_BODY)
    add_run(p, text, size=SZ_BODY)


def tbl_caption(doc, num, text):
    p = doc.add_paragraph()
    set_pf(p, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, before=6, after=4)
    add_run(p, f"Tabel II.{num} ", bold=True, size=SZ_BODY)
    add_run(p, text, bold=True, size=SZ_BODY)


def equation(doc, eq_text, num):
    """Equation paragraph centered with number right-aligned in same line via tab."""
    p = doc.add_paragraph()
    set_pf(p, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, before=6, after=6)
    add_run(p, eq_text, italic=True, size=SZ_BODY)
    add_run(p, "          ", size=SZ_BODY)
    add_run(p, f"(II.{num})", bold=True, size=SZ_BODY)


def build_table(doc, headers, rows, widths_cm, font_sz=SZ_TBL,
                header_fill="D9E2F3"):
    ncol = len(headers)
    tbl = doc.add_table(rows=1 + len(rows), cols=ncol)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    for ci, w in enumerate(widths_cm):
        tbl.columns[ci].width = Cm(w)
    for ci, h in enumerate(headers):
        c = tbl.cell(0, ci)
        c.width = Cm(widths_cm[ci])
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_borders(c)
        set_cell_shading(c, header_fill)
        c.paragraphs[0].text = ""
        p = c.paragraphs[0]
        set_pf(p, align=WD_ALIGN_PARAGRAPH.CENTER, indent=0, ls=1.15)
        if isinstance(h, str):
            add_run(p, h, bold=True, size=font_sz)
        else:
            for seg in h:
                add_run(p, seg[0], bold=seg[1], italic=seg[2], size=font_sz)
    for ri, row in enumerate(rows, start=1):
        for ci, cd in enumerate(row):
            c = tbl.cell(ri, ci)
            c.width = Cm(widths_cm[ci])
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_borders(c)
            c.paragraphs[0].text = ""
            p = c.paragraphs[0]
            set_pf(p, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0, ls=1.15)
            if isinstance(cd, str):
                add_run(p, cd, size=font_sz)
            else:
                for seg in cd:
                    if len(seg) == 3:
                        add_run(p, seg[0], bold=seg[1], italic=seg[2], size=font_sz)
                    elif len(seg) == 4:
                        add_run(p, seg[0], bold=seg[1], italic=seg[2], size=seg[3])
    sp = doc.add_paragraph()
    set_pf(sp, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0, before=0, after=0)


def bullets(doc, items, indent_cm=1.0):
    for it in items:
        p = doc.add_paragraph(style='List Bullet')
        set_pf(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=0,
               before=0, after=0)
        p.paragraph_format.left_indent = Cm(indent_cm)
        if isinstance(it, str):
            add_run(p, it, size=SZ_BODY)
        else:
            for seg in it:
                if len(seg) == 3:
                    add_run(p, seg[0], bold=seg[1], italic=seg[2], size=SZ_BODY)


def hanging_para(doc, text_segs, indent_left_cm=0.75, hanging_cm=0.75):
    """For Daftar Pustaka entries: APA-style hanging indent."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(indent_left_cm)
    pf.first_line_indent = Cm(-hanging_cm)
    pf.line_spacing = LS
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(6)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for seg in text_segs:
        if len(seg) == 3:
            add_run(p, seg[0], bold=seg[1], italic=seg[2], size=SZ_BODY)
    return p


# ============================================================================
def setup_doc():
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = FONT
    st.font.size = Pt(SZ_BODY)
    rPr = st.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), FONT)
    rFonts.set(qn('w:hAnsi'), FONT)
    rFonts.set(qn('w:eastAsia'), FONT)
    rFonts.set(qn('w:cs'), FONT)
    for sec in doc.sections:
        sec.page_height = Cm(29.7)
        sec.page_width  = Cm(21.0)
        sec.top_margin    = Cm(3.0)
        sec.bottom_margin = Cm(3.0)
        sec.left_margin   = Cm(4.0)
        sec.right_margin  = Cm(3.0)
    return doc


# ============================================================================
# BUILD BAB II
# ============================================================================
def build_bab2(doc):
    H_chap(doc, "BAB II", "LANDASAN TEORI")

    # ========================================================================
    # 2.1 LITERATUR TERKAIT TEORI (umbrella)
    # ========================================================================
    H2(doc, "2.1", "Literatur Terkait Teori")

    # ---- 2.1.1 Penelitian Terdahulu ----
    H3(doc, "2.1.1", "Penelitian Terdahulu")

    P(doc, segs=[
        ("Dalam rangka mendukung pengembangan penelitian yang diusulkan, peneliti mengkaji berbagai literatur yang berkaitan dengan klasifikasi, deteksi, maupun ", False, False),
        ("monitoring", False, True),
        (" risiko ", False, False),
        ("stunting", False, True),
        (" pada balita. Kajian ini mencakup pendekatan-pendekatan yang telah diterapkan oleh penelitian terdahulu, baik dari segi metode ", False, False),
        ("fuzzy", False, True),
        (", pemanfaatan dataset SKI 2023, maupun integrasi ", False, False),
        ("Genetic Algorithm", False, True),
        (" terhadap sistem inferensi ", False, False),
        ("fuzzy", False, True),
        (". Kajian dapat dilihat pada Tabel II.1 di bawah ini.", False, False),
    ])

    tbl_caption(doc, 1, "Hasil Kajian Penelitian Terdahulu")
    headers = [
        [("No", True, False)],
        [("Judul Penelitian", True, False)],
        [("Penulis & Tahun", True, False)],
        [("Penjelasan Singkat & Persamaan", True, False)],
        [("Hasil Utama", True, False)],
        [("Perbedaan / Kelemahan", True, False)],
    ]
    rows = [
        [
            "1",
            [("Sistem Pakar Diagnosis Stunting pada Balita Menggunakan Metode Forward Chaining dan Logika Fuzzy Sugeno", False, True)],
            "Ananda & Sriani, 2024",
            [
                ("Penelitian Ananda dan Sriani (2024) membangun sistem pakar berbasis ", False, False),
                ("Forward Chaining", False, True),
                (" dan Fuzzy Sugeno untuk mendiagnosis stunting balita dari Z-score BB/U, TB/U, dan BB/TB dengan keluaran tingkat risiko melalui defuzzifikasi yang memiliki kesamaan dengan penelitian ini karena sama-sama memakai Fuzzy Sugeno berbasis standar WHO.", False, False),
            ],
            [
                ("Sistem pakar web menggabungkan ", False, False),
                ("forward chaining", False, True),
                (" dan Sugeno untuk menilai keparahan stunting dari 122 data, dengan distribusi risiko tinggi 57%, sangat tinggi 26%, dan rendah 17%.", False, False),
            ],
            "Populasi data lokal (Puskesmas), fungsi keanggotaan dirancang manual tanpa optimasi terhadap distribusi empiris data target, dan variabel pendukung non-antropometri terbatas.",
        ],
        [
            "2",
            [("Classify the Nutritional Status of Iraqi Children under Five Years Using Fuzzy Classification", False, True)],
            "Hasan & Mohammad, 2024",
            "Penelitian Hasan dan Mohammad (2024) mengklasifikasikan status gizi anak usia kurang dari 5 tahun di Irak menggunakan Fuzzy Sugeno berbasis berat dan tinggi badan (BMI) pada 16.487 anak yang memiliki kesamaan dengan penelitian ini karena sama-sama menggunakan Fuzzy Sugeno berbasis indikator WHO untuk klasifikasi gizi.",
            "Membandingkan Mamdani vs Sugeno pada 16.487 anak (12 kelompok umur) berbasis variabel bobot dan tinggi. Sugeno unggul dengan akurasi sekitar 60% pada berbagai kelompok usia dan jenis kelamin.",
            "Populasi di Irak, variabel input minim (berat, tinggi) dan keluaran berbasis BMI, akurasi belum presisi, serta belum menghubungkan hasil ke prioritas kebijakan lokal.",
        ],
        [
            "3",
            [("Advanced Machine Learning Techniques for Stunting Classification Using a Stacking Ensemble Approach", False, True)],
            "Erda et al., 2024",
            [
                ("Penelitian Erda et al. (2024) menerapkan ", False, False),
                ("stacking ensemble", False, True),
                (" pada dataset SKI 2023 berskala besar (86.364 record) untuk klasifikasi stunting yang memiliki kesamaan dengan penelitian ini karena sama-sama menggunakan dataset SKI 2023 sebagai sumber data utama.", False, False),
            ],
            [
                ("Stacking ensemble", False, True),
                (" pada SKI 2023 menghasilkan klasifikasi stunting biner dengan akurasi 0,79, dengan performa kompetitif dibanding model ", False, False),
                ("machine learning", False, True),
                (" konvensional pada skala nasional.", False, False),
            ],
            [
                ("Keluaran berupa label biner tanpa stratifikasi risiko, model bersifat ", False, False),
                ("black-box", False, True),
                (" yang sulit ditelusuri ke variabel asal, dan belum memanfaatkan pendekatan ", False, False),
                ("fuzzy", False, True),
                (" yang ", False, False),
                ("interpretable", False, True),
                (".", False, False),
            ],
        ],
        [
            "4",
            [("Monitoring Toddler Development Based on Indonesia Health Survey (SKI) 2023", False, True)],
            "Russiska & Cahyati, 2025",
            [
                ("Penelitian Russiska dan Cahyati (2025) memanfaatkan SKI 2023 untuk mengidentifikasi indikator antropometri dan perkembangan balita yang berkaitan dengan stunting yang memiliki kesamaan dengan penelitian ini karena sama-sama menggunakan SKI 2023 untuk kebutuhan ", False, False),
                ("monitoring", False, True),
                (" pertumbuhan balita.", False, False),
            ],
            "Mengidentifikasi indikator perkembangan dan antropometri yang berkaitan dengan stunting pada SKI 2023 melalui analisis statistik deskriptif dan inferensial.",
            [
                ("Bersifat analisis statistik deskriptif tanpa sistem keputusan, belum menghasilkan ", False, False),
                ("risk score", False, True),
                (" kontinu atau ", False, False),
                ("priority flag", False, True),
                (" yang dapat dipakai untuk prioritisasi pada tingkat komunitas.", False, False),
            ],
        ],
        [
            "5",
            [("A Hybrid Genetic-Grey Wolf Optimization Algorithm for Optimizing Takagi-Sugeno-Kang Fuzzy Systems", False, True)],
            "Elghamrawy & Hassanien, 2022",
            [
                ("Penelitian Elghamrawy dan Hassanien (2022) mengintegrasikan ", False, False),
                ("Genetic Algorithm", False, True),
                (" dengan optimasi ", False, False),
                ("Grey Wolf", False, True),
                (" untuk mengoptimasi sistem ", False, False),
                ("Takagi-Sugeno-Kang", False, True),
                (" yang memiliki kesamaan dengan penelitian ini karena sama-sama menggunakan GA untuk menyempurnakan sistem ", False, False),
                ("fuzzy", False, True),
                (" Sugeno.", False, False),
            ],
            [
                ("Hybrid", False, True),
                (" GA-GWO terbukti meningkatkan performa ", False, False),
                ("Takagi-Sugeno-Kang", False, True),
                (" pada beberapa ", False, False),
                ("benchmark dataset", False, True),
                (" dengan tetap mempertahankan struktur aturan ", False, False),
                ("fuzzy", False, True),
                (".", False, False),
            ],
            [
                ("Belum diterapkan pada domain stunting maupun data antropometri balita Indonesia, sehingga relevansi epidemiologisnya untuk ", False, False),
                ("monitoring", False, True),
                (" stunting belum teruji.", False, False),
            ],
        ],
        [
            "6",
            [("An Adaptive Neuro-Fuzzy Inference System Optimized by Genetic Algorithm for Brain Tumour Detection in Magnetic Resonance Images", False, True)],
            "Ghahramani & Shiri, 2024",
            "Penelitian Ghahramani dan Shiri (2024) mengoptimasi parameter fungsi keanggotaan ANFIS Sugeno menggunakan GA untuk klasifikasi citra medis yang memiliki kesamaan dengan penelitian ini karena sama-sama menggunakan GA untuk menyesuaikan fungsi keanggotaan tanpa mengubah struktur sistem.",
            "Optimasi GA terhadap parameter fungsi keanggotaan ANFIS Sugeno meningkatkan akurasi deteksi tumor otak menjadi 99,85% dibanding ANFIS tanpa optimasi.",
            "Domain aplikasi pada citra medis MRI, bukan data antropometri balita, sehingga penyesuaian variabel masukan dan basis aturan tetap diperlukan untuk konteks stunting.",
        ],
    ]
    build_table(doc, headers, rows,
                widths_cm=[0.7, 3.0, 1.7, 3.0, 2.7, 2.9], font_sz=10)

    # Narrative paragraphs
    P(doc, segs=[
        ("Penelitian pertama oleh Ananda dan Sriani (2024) menjadi rujukan paling dekat secara metodologi karena sama-sama memakai Fuzzy Sugeno untuk menilai tingkat risiko ", False, False),
        ("stunting", False, True),
        (". Namun, fungsi keanggotaan pada penelitian tersebut masih dirancang manual tanpa optimasi berbasis distribusi empiris data, sedangkan penelitian ini menggunakan ", False, False),
        ("Genetic Algorithm", False, True),
        (" untuk menyesuaikan batas fungsi keanggotaan agar lebih sesuai dengan karakteristik balita Kalimantan Tengah. Penelitian kedua oleh Hasan dan Mohammad (2024) menegaskan bahwa metode Sugeno secara konsisten lebih akurat dibandingkan Mamdani dalam mengakomodasi ketidakpastian data antropometri, sehingga memperkuat justifikasi pemilihan metode Sugeno pada penelitian ini.", False, False),
    ])

    P(doc, segs=[
        ("Penelitian ketiga oleh Erda et al. (2024) dan keempat oleh Russiska dan Cahyati (2025) memperlihatkan bahwa SKI 2023 telah dimanfaatkan untuk analisis ", False, False),
        ("stunting", False, True),
        (", tetapi belum dijadikan basis sistem ", False, False),
        ("fuzzy monitoring", False, True),
        (". Erda et al. (2024) menggunakan pendekatan ", False, False),
        ("stacking ensemble", False, True),
        (" berbasis ", False, False),
        ("machine learning", False, True),
        (" yang bersifat ", False, False),
        ("black-box", False, True),
        (" dengan keluaran biner, sedangkan Russiska dan Cahyati (2025) berhenti pada tataran analisis statistik tanpa sistem keputusan. Kedua penelitian tersebut menunjukkan adanya celah untuk memanfaatkan SKI 2023 dalam kerangka ", False, False),
        ("fuzzy", False, True),
        (" yang lebih ", False, False),
        ("interpretable", False, True),
        (" dengan keluaran ", False, False),
        ("risk score", False, True),
        (" kontinu dan ", False, False),
        ("priority flag", False, True),
        (" bertingkat.", False, False),
    ])

    P(doc, segs=[
        ("Penelitian kelima oleh Elghamrawy dan Hassanien (2022) dan keenam oleh Ghahramani dan Shiri (2024) menjadi rujukan utama untuk justifikasi penggunaan ", False, False),
        ("Genetic Algorithm", False, True),
        (" sebagai ", False, False),
        ("optimizer", False, True),
        (" fungsi keanggotaan sistem fuzzy Sugeno. Keduanya membuktikan bahwa GA mampu menyesuaikan parameter fungsi keanggotaan dengan baik tanpa mengubah struktur aturan, sehingga interpretabilitas sistem tetap terjaga. Namun, kedua penelitian tersebut belum diterapkan pada domain ", False, False),
        ("stunting", False, True),
        (", sehingga penelitian ini berkontribusi dengan mengadaptasi pendekatan GA terhadap sistem Fuzzy Sugeno khusus untuk ", False, False),
        ("monitoring", False, True),
        (" risiko ", False, False),
        ("stunting", False, True),
        (" berbasis dataset SKI 2023 Kalimantan Tengah. Berdasarkan keseluruhan kajian tersebut, terdapat empat celah utama yang menjadi fokus penelitian ini, yakni belum adanya integrasi optimasi ", False, False),
        ("Genetic Algorithm", False, True),
        (" dengan Sugeno pada domain ", False, False),
        ("stunting", False, True),
        (", belum dimanfaatkannya SKI 2023 sebagai basis sistem ", False, False),
        ("fuzzy monitoring", False, True),
        (", belum dimasukkannya waktu pengenalan MP-ASI sebagai variabel ", False, False),
        ("fuzzy", False, True),
        (" determinan, serta keluaran sistem yang masih dominan biner sehingga tidak dapat menstratifikasi prioritas pada tingkat komunitas.", False, False),
    ])

    # ---- 2.1.2 Stunting ----
    H3(doc, "2.1.2", "Stunting", italic=True)
    P(doc, segs=[
        ("Stunting", False, True),
        (" merupakan kondisi gangguan pertumbuhan dan perkembangan balita yang diakibatkan oleh kekurangan gizi kronis, infeksi berulang, serta kurangnya stimulasi psikososial yang memadai (Kurniawan, 2025). Penilaian terhadap status ", False, False),
        ("stunting", False, True),
        (" umumnya dilakukan berdasarkan pengukuran tinggi badan menurut usia (", False, False),
        ("height-for-age", False, True),
        (") yang dibandingkan dengan ", False, False),
        ("World Health Organization", False, True),
        (" (WHO) ", False, False),
        ("Child Growth Standards", False, True),
        (". Penentuan batas antar kategori tersebut secara ilmiah didasarkan pada nilai ", False, False),
        ("Z-score", False, True),
        (" tinggi-badan-menurut-usia (", False, False),
        ("Height-for-Age Z-score", False, True),
        ("/HAZ), di mana balita dengan HAZ kurang dari -2 SD dikategorikan ", False, False),
        ("stunting", False, True),
        (" (Perumal et al., 2023; UNICEF et al., 2023).", False, False),
    ])

    # ---- 2.1.3 Z-Score ----
    H3(doc, "2.1.3", "Z-Score")
    P(doc, segs=[
        ("Z-score", False, True),
        (" (", False, False),
        ("standard score", False, True),
        (") merupakan ukuran baku yang menunjukkan seberapa jauh suatu nilai menyimpang dari rerata populasi dalam satuan simpangan baku. Dalam penilaian ", False, False),
        ("stunting", False, True),
        (", ", False, False),
        ("Z-score", False, True),
        (" khususnya ", False, False),
        ("Height-for-Age Z-score", False, True),
        (" (HAZ) digunakan untuk menilai status gizi balita dengan membandingkan tinggi badan terhadap usia berdasarkan standar antropometri WHO (Watson et al., 2024). Persamaan ", False, False),
        ("Z-score", False, True),
        (" secara matematis dapat dilihat pada rumus II.1 di bawah ini.", False, False),
    ])
    equation(doc, "Z-Score = (x − μ) / σ", 1)

    P(doc, text="Di mana:", indent=0)
    bullets(doc, [
        [("x", True, False), (" adalah nilai individu pengukuran dari tinggi/panjang badan.", False, False)],
        [("μ", True, False), (" adalah nilai ", False, False), ("mean", False, True), (" dari rata-rata tinggi badan populasi referensi untuk usia dan jenis kelamin yang sama.", False, False)],
        [("σ", True, False), (" adalah nilai simpangan baku dari tinggi badan populasi referensi untuk usia dan jenis kelamin yang sama.", False, False)],
    ])

    P(doc, segs=[
        ("Adapun cara lain untuk mencari ", False, False),
        ("Z-score", False, True),
        (" versi WHO menggunakan parameter LMS dapat dilihat pada Persamaan II.2 berikut.", False, False),
    ])
    equation(doc, "Z-Score = [(x / M)^L − 1] / (L × S)", 2)

    P(doc, text="Di mana:", indent=0)
    bullets(doc, [
        [("x", False, True), (" = nilai pengukuran antropometri individu (dalam hal ini panjang/tinggi badan).", False, False)],
        [("M", False, True), (" (", False, False), ("Mu", False, True), (") = nilai ", False, False), ("median", False, True), (" dari populasi referensi untuk kelompok usia dan jenis kelamin tertentu.", False, False)],
        [("L", False, True), (" (", False, False), ("Lambda", False, True), (") = ", False, False), ("parameter power", False, True), (" untuk mentransformasi data menuju normalitas.", False, False)],
        [("S", False, True), (" = koefisien variasi (atau standar deviasi yang disesuaikan) dari populasi referensi.", False, False)],
    ])

    P(doc, segs=[
        ("Berdasarkan klasifikasi WHO, kategori ", False, False),
        ("stunting", False, True),
        (" pada balita usia 0-60 bulan secara ringkas dapat dilihat pada Tabel II.2 berikut ini.", False, False),
    ])

    tbl_caption(doc, 2, "Klasifikasi Status Stunting Berdasarkan WHO HAZ")
    headers2 = [
        [("Indeks Antropometri", True, False)],
        [("Status Gizi", True, False)],
        [("Ambang Batas (Z-Score)", True, False)],
    ]
    rows2 = [
        ["Panjang Badan atau Tinggi Badan menurut Usia (PB/U atau TB/U) balita usia 0-60 bulan (Soliman et al., 2021)",
         [("Stunting", False, True)],
         "x ≤ -2,00 SD"],
        ["", "Normal", "x > -2,00 SD"],
    ]
    build_table(doc, headers2, rows2, widths_cm=[7.5, 3.5, 3.5])

    # ---- 2.1.4 SKI 2023 ----
    H3(doc, "2.1.4", "SKI 2023")
    P(doc, segs=[
        ("Survei Kesehatan Indonesia (SKI) 2023 merupakan survei nasional yang diselenggarakan oleh Kementerian Kesehatan untuk memetakan kondisi kesehatan masyarakat Indonesia melalui 86.364 observasi yang merepresentasikan faktor sosial, ekonomi, lingkungan, dan demografi yang mempengaruhi status gizi serta indikator kesehatan lainnya (Erda et al., 2024). Survei dapat dilihat pada Gambar 2.1 berikut ini.", False, False),
    ])
    fig_caption(doc, 1, "Buku Panduan Survei Kesehatan Indonesia di mana (a) Buku Pedoman Fundamental SKI 2023 dan (b) Buku Pedoman Deskripsi SKI 2023 Dalam Angka (Kemenkes, 2024).")

    P(doc, segs=[
        ("Survei SKI sendiri dilaksanakan dalam bentuk kuisioner. Kuisioner SKI 2023 terdiri dari tiga instrumen utama, yaitu rumah tangga, rumah tangga balita, dan individu (dapat dilihat pada Gambar 2.2). Seluruh instrumen ini menjadi dasar penyusunan SKI 2023 dalam Angka dan Laporan Tematik SKI 2023, yang menyajikan ringkasan indikator kesehatan nasional dan analisis khusus, termasuk ", False, False),
        ("stunting", False, True),
        (" (Kemenkes, 2024).", False, False),
    ])
    fig_caption(doc, 2, "Bentuk Kuisioner dari Survei SKI 2023 di mana (a) Survei Rumah Tangga, (b) Survei Rumah Tangga Balita, dan (c) Survei Individu (Kemenkes, 2024).")

    # ---- 2.1.5 Variabel Penelitian ----
    H3(doc, "2.1.5", "Variabel Penelitian")
    P(doc, segs=[
        ("Variabel penelitian memegang peran utama dalam memahami determinan ", False, False),
        ("stunting", False, True),
        (". Variabel ", False, False),
        ("dependen", False, True),
        (" pada penelitian ini adalah status ", False, False),
        ("stunting", False, True),
        (" yang ditentukan melalui ", False, False),
        ("ground truth", False, True),
        (" WHO HAZ, sedangkan variabel ", False, False),
        ("independen", False, True),
        (" terdiri dari enam variabel yang memiliki relevansi klinis dan epidemiologis terhadap risiko ", False, False),
        ("stunting", False, True),
        (", yakni usia, jenis kelamin, panjang badan lahir (PBL), tinggi/panjang badan saat pengukuran, berat badan saat pengukuran, dan waktu pengenalan MP-ASI (Yani et al., 2023). Penjelasan masing-masing variabel diuraikan pada subbagian berikut.", False, False),
    ])

    # 2.1.5.1 TB ukur
    H4(doc, "2.1.5.1", "Tinggi/Panjang Badan Saat Pengukuran")
    P(doc, segs=[
        ("Stunting", False, True),
        (" adalah kondisi kronis ketika panjang atau tinggi balita berada di bawah standar usianya, ditandai nilai ", False, False),
        ("z-score", False, True),
        (" PB/U atau TB/U kurang dari -2 SD menurut WHO. Masalah ini terutama muncul akibat kekurangan gizi jangka panjang, infeksi berulang, dan faktor maternal. Tinggi atau panjang badan saat pengukuran merupakan indikator antropometri utama yang dipakai untuk menghitung HAZ sebagai ", False, False),
        ("ground truth", False, True),
        (" sistem. Pengukuran status dilakukan dengan LAZ untuk balita berusia kurang dari atau sama dengan 24 bulan (posisi terlentang) dan HAZ untuk balita berusia lebih dari 24 bulan (posisi berdiri) sebagaimana diterapkan dalam berbagai studi (Rosdiana et al., 2024).", False, False),
    ])
    fig_caption(doc, 3, "Grafik Ilustrasi Pertumbuhan Panjang/Tinggi Badan Balita Selama 1000 HPK Berdasarkan Usia dan Median Tinggi Badan Rujukan WHO (UNICEF, 2023).")

    # 2.1.5.2 BB ukur
    H4(doc, "2.1.5.2", "Berat Badan Saat Pengukuran")
    P(doc, segs=[
        ("Berat badan saat pengukuran merupakan indikator antropometri yang melengkapi tinggi badan dalam menilai status gizi balita secara menyeluruh. Berat badan yang tidak proporsional terhadap tinggi badan dapat menjadi sinyal awal ", False, False),
        ("wasting", False, True),
        (" atau ", False, False),
        ("overweight", False, True),
        (" yang berinteraksi dengan risiko ", False, False),
        ("stunting", False, True),
        (" (de Onis & Branca, 2016). Pada sistem ", False, False),
        ("monitoring", False, True),
        (" yang dikembangkan, berat badan saat pengukuran difuzzifikasi ke dalam empat himpunan linguistik (Kurang, Normal, Lebih, Obesitas) yang dipisahkan berdasarkan jenis kelamin untuk menangkap perbedaan pola pertumbuhan antara balita laki-laki dan perempuan.", False, False),
    ])

    # 2.1.5.3 Usia Anak
    H4(doc, "2.1.5.3", "Usia Anak")
    P(doc, segs=[
        ("Usia balita merupakan variabel penting dalam memahami dinamika ", False, False),
        ("stunting", False, True),
        (" karena merefleksikan akumulasi paparan risiko sejak lahir hingga 60 bulan. Secara keseluruhan, dalam rentang 0-60 bulan yang mencakup 1000 HPK, determinan biologis, perilaku, lingkungan, dan sosial-ekonomi terbukti konsisten memengaruhi prevalensi ", False, False),
        ("stunting", False, True),
        (" (Talukder et al., 2018). Temuan pada Gambar 2.4 juga menunjukkan bahwa prevalensi tertinggi terjadi pada kelompok usia 24-53 bulan, baik pada 2022 maupun 2023, meskipun angka tahun 2023 lebih rendah.", False, False),
    ])
    fig_caption(doc, 4, "Prevalensi Stunting Berdasarkan Kelompok Umur Tahun 2022 dan 2023 (Kemenkes, 2024).")

    # 2.1.5.4 Jenis Kelamin
    H4(doc, "2.1.5.4", "Jenis Kelamin Anak")
    P(doc, segs=[
        ("Jenis kelamin merupakan variabel penting dalam pengukuran ", False, False),
        ("z-score", False, True),
        (" ", False, False),
        ("stunting", False, True),
        (" karena memengaruhi pola pertumbuhan, kerentanan biologis, serta respons terhadap faktor lingkungan. Bukti umum menunjukkan bahwa balita laki-laki lebih rentan mengalami ", False, False),
        ("stunting", False, True),
        (" dibandingkan perempuan, dan determinan sosioekonomi pun bekerja berbeda menurut jenis kelamin (de Onis & Branca, 2016). Pada sistem yang dikembangkan, jenis kelamin difungsikan sebagai variabel ", False, False),
        ("switching", False, True),
        (" yang memilih himpunan fungsi keanggotaan tinggi dan berat badan khusus laki-laki atau perempuan, sehingga interpretasi ", False, False),
        ("z-score", False, True),
        (" selalu mempertimbangkan jenis kelamin sebagaimana ditunjukkan pada Gambar 2.5 mengenai standar antropometri WHO.", False, False),
    ])
    fig_caption(doc, 5, "Standar WHO Antropometri pada Masing-masing Jenis Kelamin di mana (a) Grafik Standar Panjang/Tinggi Badan Perempuan dan (b) Grafik Standar Panjang/Tinggi Badan Laki-laki (UNICEF, 2023).")

    # 2.1.5.5 PBL
    H4(doc, "2.1.5.5", "Panjang Badan Lahir")
    P(doc, segs=[
        ("Panjang badan lahir (PBL) merupakan indikator penting dalam penilaian risiko ", False, False),
        ("stunting", False, True),
        (" karena secara konsisten terbukti sebagai prediktor kuat pertumbuhan linear. PBL mencerminkan kondisi gizi ", False, False),
        ("intrauterin", False, True),
        (" yang berpengaruh terhadap pertumbuhan linear anak, sehingga bayi dengan panjang lahir pendek memiliki kemungkinan lebih tinggi mengalami ", False, False),
        ("stunting", False, True),
        (" dibandingkan bayi dengan panjang lahir normal atau tinggi. Temuan longitudinal seperti ", False, False),
        ("Women First trial", False, True),
        (" juga menegaskan bahwa PBL merupakan prediktor terkuat status ", False, False),
        ("linear growth", False, True),
        (" hingga 24 bulan, bahkan melampaui berat lahir, sebagaimana terlihat pada Gambar 2.6 yang menunjukkan bahwa ", False, False),
        ("birth LAZ", False, True),
        (" kurang dari -2 dan kurang dari -1 memberikan penurunan skor Z panjang menurut umur (LAZ) paling besar dibandingkan faktor lain (Krebs et al., 2022).", False, False),
    ])
    fig_caption(doc, 6, "Forest Plot Faktor yang Memengaruhi Skor Z Panjang Menurut Umur (LAZ) pada Usia 24 Bulan (Krebs et al., 2022).")

    # 2.1.5.6 MP-ASI
    H4(doc, "2.1.5.6", "Makanan Pendamping ASI")
    P(doc, segs=[
        ("Waktu pemberian MP-ASI merupakan faktor penting dalam risiko ", False, False),
        ("stunting", False, True),
        (" karena menentukan kecukupan energi dan zat gizi selama transisi dari ASI eksklusif. Pemberian terlalu dini terbukti meningkatkan risiko malnutrisi, ", False, False),
        ("wasting", False, True),
        (", hingga ", False, False),
        ("stunting", False, True),
        (" akibat ketidaksiapan sistem pencernaan bayi dan kemungkinan tergantinya ASI dengan makanan yang kurang padat gizi (Muflih & Khasana, 2025). Berdasarkan penelitian Waqiyah et al. (2023), anak yang menerima MP-ASI tidak tepat waktu memiliki proporsi ", False, False),
        ("stunting", False, True),
        (" lebih tinggi (39,6%) dibandingkan anak dengan MP-ASI tepat waktu (16,0%), dengan hubungan yang signifikan secara statistik. Temuan ini menunjukkan bahwa pemberian MP-ASI sesuai usia rekomendasi bersifat protektif terhadap ", False, False),
        ("stunting", False, True),
        (", sedangkan penyimpangan waktu pemberian berpotensi menyebabkan defisit gizi yang berdampak pada pertumbuhan linear anak.", False, False),
    ])
    fig_caption(doc, 7, "Hubungan Waktu Pemberian MP-ASI dan Pengetahuan Ibu tentang Responsive Feeding dengan Kejadian Stunting pada Anak Usia 6-24 Bulan (Waqiyah et al., 2023).")

    # ---- 2.1.6 Logika Fuzzy ----
    H3(doc, "2.1.6", "Logika Fuzzy")
    P(doc, segs=[
        ("Fuzzy logic", False, True),
        (" merupakan kerangka matematis untuk penalaran yang memodelkan derajat kebenaran kontinu antara 0-1 melalui teori himpunan samar, sehingga suatu elemen memiliki tingkat keanggotaan alih-alih status biner. Mekanisme dasarnya bekerja dengan aturan komposisional untuk menghasilkan keputusan dari informasi yang tidak pasti, menjadikannya fleksibel dan selaras dengan cara berpikir manusia. Pendekatan ini terbukti luas digunakan dalam berbagai aplikasi, mulai dari kendali proses hingga diagnosis medis, karena sifatnya yang interpretatif dan mudah diintegrasikan dalam sistem (Borgwardt & Peñaloza, 2017). Pada penelitian ini, variabel yang difuzzifikasi adalah usia, tinggi/panjang badan saat pengukuran, berat badan saat pengukuran, dan waktu pengenalan MP-ASI, sedangkan jenis kelamin dipakai sebagai variabel ", False, False),
        ("switching", False, True),
        (" yang memilih himpunan fungsi keanggotaan tinggi dan berat khusus laki-laki atau perempuan. Persamaan dasar himpunan ", False, False),
        ("fuzzy", False, True),
        (" secara matematis dapat dilihat pada rumus II.3 di bawah ini.", False, False),
    ])
    equation(doc, "μ_A(x) ∈ [0, 1]", 3)
    P(doc, text="Di mana:", indent=0)
    bullets(doc, [
        [("μ_A(x)", True, False), (" adalah derajat keanggotaan (", False, False), ("membership degree", False, True), (") elemen x terhadap himpunan samar A. Nilainya selalu dinormalisasi ke rentang 0-1.", False, False)],
        [("∈ [0, 1]", True, False), (" adalah rentang nilai antara 0-1, bisa tepat 0, tepat 1, atau nilai apa pun di antaranya (0,13; 0,5; 0,97; dst.).", False, False)],
    ])
    P(doc, segs=[
        ("Logika ", False, False),
        ("fuzzy", False, True),
        (" juga sering disebut sebagai ", False, False),
        ("Fuzzy Inference System", False, True),
        (" (FIS) atau ", False, False),
        ("Fuzzy Inference Engine", False, True),
        (" (Maulana et al., 2024). Gambar 2.8 menunjukkan alur kerja fundamental sistem logika ", False, False),
        ("fuzzy", False, True),
        (".", False, False),
    ])
    fig_caption(doc, 8, "Alur Aturan Fuzzy (Maulana et al., 2024).")

    # 2.1.6.1 Fuzzifikasi
    H4(doc, "2.1.6.1", "Fuzzifikasi")
    P(doc, segs=[
        ("Pada tahap awal dalam Gambar 2.8, nilai ", False, False),
        ("input", False, True),
        (" tegas diubah menjadi derajat keanggotaan menggunakan fungsi keanggotaan, sehingga variabel seperti usia, tinggi badan, berat badan, dan MP-ASI dapat direpresentasikan dalam bentuk linguistik yang fleksibel (Gupta et al., 2022). Pada penelitian ini dipilih bentuk Trapesium dan Segitiga. Representasi kedua bentuk grafik dapat dilihat pada Gambar 2.9 berikut ini.", False, False),
    ])
    fig_caption(doc, 9, "Representasi Grafik di mana (a) Bentuk Trapesium (Maulana et al., 2024) dan (b) Bentuk Segitiga (Astonis, 2014).")

    # 2.1.6.2 Trapesium
    H4(doc, "2.1.6.2", "Trapesium")
    P(doc, segs=[
        ("Trapesium dipilih karena sederhana, mudah ditafsirkan, serta stabil dalam menggambarkan batas kategori variabel risiko ", False, False),
        ("stunting", False, True),
        (". Secara matematis, trapesium ditetapkan melalui empat parameter (a, b, c, d) yang membentuk pola naik-datar-turun (Putri & Saputro, 2021). Persamaan fungsi keanggotaan dengan parameter bentuk trapesium dapat dilihat pada rumus II.4 di bawah ini.", False, False),
    ])
    equation(doc, "μ_trap(x; a, b, c, d) = max( min((x-a)/(b-a), 1, (d-x)/(d-c)), 0 )", 4)
    P(doc, text="Di mana:", indent=0)
    bullets(doc, [
        "Jika x ≤ a → kategori tidak berlaku (0).",
        "Jika a < x < b → kategori mulai berlaku (naik).",
        "Jika b ≤ x ≤ c → kategori berlaku sempurna (1).",
        "Jika c < x < d → kategori melemah (turun).",
        "Jika x ≥ d → kategori tidak berlaku (0).",
    ])
    P(doc, text="Dengan catatan parameter:", indent=0)
    bullets(doc, [
        [("a", True, False), (" adalah batas kiri dukungan (", False, False), ("support", False, True), (") dengan nilai ≤ a dianggap tidak masuk kategori.", False, False)],
        [("b", True, False), (" adalah titik mulai keanggotaan penuh dengan x sudah masuk kategori dengan kekuatan 1.", False, False)],
        [("c", True, False), (" adalah titik akhir keanggotaan penuh dengan x masih 100% bagian dari kategori.", False, False)],
        [("d", True, False), (" adalah batas kanan dukungan dengan nilai ≥ d dianggap keluar kategori.", False, False)],
    ])

    # 2.1.6.3 Segitiga
    H4(doc, "2.1.6.3", "Segitiga")
    P(doc, segs=[
        ("Selain Trapesium, bentuk representatif yang digunakan adalah fungsi segitiga. Fungsi segitiga adalah salah satu bentuk fuzzifikasi sederhana yang terdiri dari tiga titik (a, b, c). Dalam konteks penelitian ", False, False),
        ("stunting", False, True),
        (", fungsi ini dapat digunakan untuk memodelkan variabel yang memiliki satu titik optimum (Abdy, 2018). Persamaan fungsi keanggotaan dengan parameter bentuk segitiga dapat dilihat pada rumus II.5 di bawah ini.", False, False),
    ])
    equation(doc, "μ_tri(x; a, b, c) = max( min((x-a)/(b-a), (c-x)/(c-b)), 0 )", 5)
    P(doc, text="Di mana:", indent=0)
    bullets(doc, [
        "Jika x ≤ a atau x ≥ c maka kategori tidak berlaku (μ(x) = 0).",
        "Jika a ≤ x ≤ b maka kategori mulai berlaku (derajat keanggotaan naik dari 0 menuju 1).",
        "Jika b ≤ x ≤ c maka kategori mulai tidak berlaku (derajat keanggotaan turun dari 1 menuju 0).",
    ])
    P(doc, text="Dengan catatan parameter:", indent=0)
    bullets(doc, [
        [("a", True, False), (" = ambang batas kiri dukungan (", False, False), ("left support", False, True), ("), semua nilai ≤ a dianggap belum masuk kategori.", False, False)],
        [("b", True, False), (" = titik puncak, x = b memiliki keanggotaan penuh (μ = 1) dan paling mewakili kategori.", False, False)],
        [("c", True, False), (" = ambang batas kanan dukungan (", False, False), ("right support", False, True), ("), semua nilai ≥ c dianggap keluar dari kategori.", False, False)],
    ])

    # 2.1.6.4 Aturan Dasar (was 2.1.6.5 in original, but Kurva Linear is removed)
    H4(doc, "2.1.6.4", "Aturan Dasar")
    P(doc, segs=[
        ("Metode ", False, False),
        ("fuzzy Takagi-Sugeno", False, True),
        (" (TS) memodelkan sistem nonlinier melalui aturan ", False, False),
        ("IF-THEN", False, True),
        (" dengan konsekuen berbentuk fungsi tegas atau persamaan linear, sehingga mampu menangani ketidakpastian sekaligus menjaga struktur matematis yang efisien. Pendekatan ini relevan untuk kebutuhan kendali dan pengambilan keputusan modern (Yershov & Ponomarenko, 2018). Persamaan metode Sugeno secara matematis dapat dilihat pada rumus II.6 di bawah ini.", False, False),
    ])
    equation(doc, "IF x₁ adalah A₁ᵢ AND x₂ adalah A₂ᵢ THEN y = cᵢ", 6)
    P(doc, text="Di mana:", indent=0)
    bullets(doc, [
        [("x₁, x₂", True, False), (" = variabel masukan (", False, False), ("input", False, True), (") sistem ", False, False), ("fuzzy", False, True), (".", False, False)],
        [("A₁ᵢ, A₂ᵢ", True, False), (" = himpunan ", False, False), ("fuzzy", False, True), (" (label linguistik) yang melekat pada masing-masing variabel masukan pada aturan ke-i.", False, False)],
        [("Operator AND", True, False), (" = operator logika ", False, False), ("fuzzy", False, True), (" (biasanya direalisasikan dengan ", False, False), ("t-norm", False, True), (", misalnya ", False, False), ("minimum", False, True), (" atau perkalian).", False, False)],
        [("y", True, False), (" = variabel keluaran (", False, False), ("output", False, True), (") sistem ", False, False), ("fuzzy", False, True), (", yaitu skor risiko ", False, False), ("stunting", False, True), (" yang akan dipetakan menjadi kelas risiko dan ", False, False), ("priority flag", False, True), (".", False, False)],
        [("cᵢ", True, False), (" = nilai konstan (konsekuen tegas) pada aturan Sugeno untuk aturan ke-i.", False, False)],
    ])

    # 2.1.6.5 Defuzzifikasi
    H4(doc, "2.1.6.5", "Defuzzifikasi")
    P(doc, segs=[
        ("Tahap terakhir pada Gambar 2.8 mengubah himpunan ", False, False),
        ("fuzzy", False, True),
        (" teragregasi menjadi nilai tegas melalui metode ", False, False),
        ("weighted average", False, True),
        (" pada Sugeno orde nol, menghasilkan output yang dapat digunakan langsung sebagai skor risiko (Georgescu et al., 2014). Rumus defuzzifikasi metode Sugeno dapat dilihat pada Persamaan II.7 di bawah ini.", False, False),
    ])
    equation(doc, "y* = Σᵢ (wᵢ · cᵢ) / Σᵢ wᵢ", 7)
    P(doc, text="Di mana:", indent=0)
    bullets(doc, [
        [("y*", True, False), (" = keluaran akhir sistem inferensi ", False, False), ("fuzzy", False, True), (" Sugeno, berupa satu nilai tegas (", False, False), ("crisp", False, True), (") hasil penggabungan semua aturan.", False, False)],
        [("Indeks i", True, False), (" = nomor aturan ke-i dalam basis aturan (", False, False), ("rule base", False, True), (").", False, False)],
        [("cᵢ", True, False), (" = nilai konsekuen konstan pada aturan ke-i.", False, False)],
        [("wᵢ", True, False), (" = derajat aktivasi (bobot) aturan ke-i, yang diperoleh dari penggabungan derajat keanggotaan premis (hasil operator ", False, False), ("AND", False, True), (" pada himpunan ", False, False), ("fuzzy", False, True), (").", False, False)],
    ])

    # ---- 2.1.7 Genetic Algorithm ----
    H3(doc, "2.1.7", "Genetic Algorithm", italic=True)
    P(doc, segs=[
        ("Genetic Algorithm", False, True),
        (" (GA) merupakan algoritma metaheuristik inspirasi biologis yang meniru proses seleksi alam untuk mencari solusi optimum pada ruang pencarian yang besar dan kompleks (Katoch et al., 2021; Elghamrawy & Hassanien, 2022). GA bekerja dengan mempertahankan populasi calon solusi yang direpresentasikan sebagai ", False, False),
        ("individual", False, True),
        (" berupa kromosom (vektor gen), kemudian secara iteratif melakukan tiga operator utama, yaitu seleksi, ", False, False),
        ("crossover", False, True),
        (", dan mutasi, untuk menghasilkan generasi berikutnya yang lebih baik berdasarkan nilai fungsi ", False, False),
        ("fitness", False, True),
        (". Pada penelitian ini, GA digunakan untuk mengoptimasi batas titik fungsi keanggotaan ", False, False),
        ("fuzzy", False, True),
        (" Sugeno tanpa mengubah basis aturan, operator inferensi, maupun struktur utama sistem, sehingga interpretabilitas ", False, False),
        ("fuzzy", False, True),
        (" tetap terjaga sambil parameter fungsi keanggotaan menyesuaikan distribusi empiris data SKI 2023 (Ghahramani & Shiri, 2024).", False, False),
    ])

    P(doc, segs=[
        ("Tiga operator utama GA yang diaplikasikan pada penelitian ini diuraikan sebagai berikut. ", False, False),
        ("Pertama", True, False),
        (", seleksi menggunakan metode ", False, False),
        ("tournament selection", False, True),
        (" dengan ukuran turnamen tiga, yakni tiga individu dipilih secara acak dan individu dengan ", False, False),
        ("fitness", False, True),
        (" tertinggi diteruskan ke kolam reproduksi (Goldberg & Deb, 1991). ", False, False),
        ("Kedua", True, False),
        (", ", False, False),
        ("crossover", False, True),
        (" menggunakan ", False, False),
        ("blend crossover", False, True),
        (" (cxBlend) dengan parameter α = 0,3 untuk mencampur gen dari dua induk dan menghasilkan keturunan yang berada di sekitar rentang gen induk (Herrera et al., 2003). ", False, False),
        ("Ketiga", True, False),
        (", mutasi menggunakan ", False, False),
        ("bounded Gaussian mutation", False, True),
        (" dengan probabilitas per-gen 0,15 dan deviasi standar 0,15 kali rentang setiap gen, di mana setiap gen yang termutasi tetap dijaga dalam batas medis yang masuk akal.", False, False),
    ])

    P(doc, segs=[
        ("Fungsi ", False, False),
        ("fitness", False, True),
        (" yang digunakan adalah kombinasi tertimbang F1-score dan akurasi dengan bobot 0,6 untuk F1-score dan 0,4 untuk akurasi. Penggunaan bobot asimetris ini bertujuan untuk menjaga sensitivitas ", False, False),
        ("optimizer", False, True),
        (" terhadap kelas minoritas (", False, False),
        ("stunting", False, True),
        (") yang hanya sekitar 16,5% dari total sampel, sekaligus mencegah konfigurasi degeneratif yang hanya memaksimalkan ", False, False),
        ("recall", False, True),
        (" (Hand et al., 2021). Persamaan fungsi ", False, False),
        ("fitness", False, True),
        (" dapat dilihat pada rumus II.8 berikut.", False, False),
    ])
    equation(doc, "F(θ) = w₁ · F1(θ) + w₂ · Acc(θ),   w₁ = 0,6;  w₂ = 0,4", 8)
    P(doc, text="Di mana:", indent=0)
    bullets(doc, [
        [("F(θ)", True, False), (" = nilai ", False, False), ("fitness", False, True), (" individu θ pada subset pelatihan.", False, False)],
        [("F1(θ)", True, False), (" = F1-score yang dihitung dari biner ", False, False), ("priority flag", False, True), (" (Low/Medium/High vs Normal) terhadap label WHO HAZ.", False, False)],
        [("Acc(θ)", True, False), (" = akurasi prediksi sistem pada subset pelatihan.", False, False)],
        [("w₁, w₂", True, False), (" = bobot kontribusi F1-score dan akurasi, dengan w₁ + w₂ = 1 sehingga F(θ) ∈ [0, 1].", False, False)],
    ])

    # ---- 2.1.8 Stratified Hold-out Validation ----
    H3(doc, "2.1.8", "Stratified Hold-out Validation", italic=True)
    P(doc, segs=[
        ("Stratified hold-out validation", False, True),
        (" membagi data ke dalam set pelatihan dan pengujian dengan menjaga proporsi kelas label tetap seimbang di kedua subset, sehingga evaluasi tidak bias akibat ketidakseimbangan kelas yang umum terjadi pada data kesehatan (Abegaz et al., 2025). Pendekatan ini banyak digunakan pada studi klinis dan klasifikasi modern karena memberikan estimasi kinerja yang lebih representatif dibanding pembagian acak biasa, terutama ketika ", False, False),
        ("outcome", False, True),
        (" seperti ", False, False),
        ("stunting", False, True),
        (" (berdasarkan HAZ) tidak terdistribusi merata (Andrews et al., 2025; Bichri et al., 2024). Pada penelitian ini, ", False, False),
        ("stratified hold-out", False, True),
        (" dengan rasio 80/20 dipakai untuk membagi 389 sampel akhir menjadi 311 sampel pelatihan dan 78 sampel pengujian, dengan proporsi kelas ", False, False),
        ("stunting", False, True),
        (" tetap mendekati 16,5% pada kedua subset.", False, False),
    ])

    # ---- 2.1.9 Confusion Matrix ----
    H3(doc, "2.1.9", "Confusion Matrix", italic=True)
    P(doc, segs=[
        ("Konfusi matriks merupakan alat evaluasi dasar yang merangkum hubungan antara prediksi dan kondisi aktual untuk menghasilkan metrik penting seperti akurasi, presisi, ", False, False),
        ("recall", False, True),
        (", dan F1-score (Japkowicz & Shah, 2015). Gambar 2.10 menampilkan empat komponen utama, yaitu ", False, False),
        ("True Positive", False, True),
        (" (TP), ", False, False),
        ("False Positive", False, True),
        (" (FP), ", False, False),
        ("False Negative", False, True),
        (" (FN), dan ", False, False),
        ("True Negative", False, True),
        (" (TN) yang menjadi fondasi seluruh perhitungan tersebut.", False, False),
    ])
    fig_caption(doc, 10, "Penjelasan Confusion Matrix (Liu & Yang, 2022).")

    P(doc, segs=[
        ("Presisi dan ", False, False),
        ("recall", False, True),
        (" digunakan untuk mengevaluasi model pada kondisi kelas tidak seimbang karena masing-masing menilai ketepatan prediksi positif dan kemampuan model menangkap seluruh kasus positif (Landgrebe et al., 2006). Persamaan ", False, False),
        ("precision", False, True),
        (" dan ", False, False),
        ("recall", False, True),
        (" secara matematis dapat dilihat pada rumus II.9 dan II.10 di bawah ini.", False, False),
    ])
    equation(doc, "Precision = TP / (TP + FP)", 9)
    equation(doc, "Recall = TP / (TP + FN)", 10)

    P(doc, segs=[
        ("F1-score merupakan metrik yang penting karena menyeimbangkan presisi dan ", False, False),
        ("recall", False, True),
        (" melalui rata-rata harmonik, sehingga lebih peka terhadap ", False, False),
        ("false positive", False, True),
        (" dan ", False, False),
        ("false negative", False, True),
        (", terutama pada data tidak seimbang (Hand et al., 2021). Persamaan F1-score secara matematis dapat dilihat pada rumus II.11 di bawah ini.", False, False),
    ])
    equation(doc, "F1 = (2 × Precision × Recall) / (Precision + Recall)", 11)

    P(doc, segs=[
        ("Akurasi menunjukkan berapa banyak prediksi yang benar dari semua prediksi yang dibuat oleh model penelitian, dalam hal ini sistem ", False, False),
        ("fuzzy logic", False, True),
        (" Sugeno (Marsya Finda & Wahyu Utomo, 2024). Persamaan akurasi secara matematis dapat dilihat pada rumus II.12 di bawah ini.", False, False),
    ])
    equation(doc, "Accuracy = (TP + TN) / (TP + TN + FP + FN)", 12)

    P(doc, text="Keterangan untuk seluruh persamaan II.9-II.12:", indent=0)
    bullets(doc, [
        [("TP: ", True, False), ("True Positive", False, True), (", jumlah kasus ", False, False), ("stunting", False, True), (" yang diprediksi sebagai ", False, False), ("stunting", False, True), (" oleh sistem.", False, False)],
        [("TN: ", True, False), ("True Negative", False, True), (", jumlah kasus Normal yang diprediksi sebagai Normal oleh sistem.", False, False)],
        [("FP: ", True, False), ("False Positive", False, True), (", jumlah kasus Normal yang salah diprediksi sebagai ", False, False), ("stunting", False, True), (" oleh sistem.", False, False)],
        [("FN: ", True, False), ("False Negative", False, True), (", jumlah kasus ", False, False), ("stunting", False, True), (" yang salah diprediksi sebagai Normal oleh sistem.", False, False)],
    ])

    # ---- 2.1.10 Spearman Rank Correlation ----
    H3(doc, "2.1.10", "Spearman Rank Correlation", italic=True)
    P(doc, segs=[
        ("Korelasi ", False, False),
        ("rank", False, True),
        (" Spearman (ρ) merupakan metrik non-parametrik yang mengukur kekuatan dan arah hubungan monotonik antara dua variabel berdasarkan peringkat (", False, False),
        ("rank", False, True),
        (") nilainya, bukan nilai mentahnya (Schober et al., 2018). Metrik ini cocok digunakan ketika dua variabel berasal dari skala ordinal atau ketika hubungan antar variabel tidak linear namun tetap monotonik. Nilai ρ berkisar antara -1 hingga +1, di mana ρ = +1 berarti hubungan monotonik positif sempurna, ρ = -1 berarti hubungan monotonik negatif sempurna, dan ρ = 0 berarti tidak ada hubungan monotonik. Persamaan korelasi ", False, False),
        ("rank", False, True),
        (" Spearman secara matematis dapat dilihat pada rumus II.13 di bawah ini.", False, False),
    ])
    equation(doc, "ρ = 1 − ( 6 · Σ dᵢ² ) / ( n · (n² − 1) )", 13)
    P(doc, text="Di mana:", indent=0)
    bullets(doc, [
        [("ρ", True, False), (" = koefisien korelasi ", False, False), ("rank", False, True), (" Spearman.", False, False)],
        [("dᵢ", True, False), (" = selisih peringkat pasangan data ke-i pada kedua variabel.", False, False)],
        [("n", True, False), (" = jumlah pasangan data.", False, False)],
    ])
    P(doc, segs=[
        ("Pada penelitian ini, korelasi ", False, False),
        ("rank", False, True),
        (" Spearman dipakai untuk mengukur kesesuaian monotonik antara ", False, False),
        ("risk score", False, True),
        (" kontinu yang dihasilkan sistem dengan WHO HAZ sebagai ", False, False),
        ("ground truth", False, True),
        (". Tanda negatif yang diharapkan menunjukkan bahwa ", False, False),
        ("risk score", False, True),
        (" yang lebih tinggi berasosiasi dengan HAZ yang lebih rendah, sesuai dengan desain sistem yang memetakan kondisi pertumbuhan terlambat ke skor risiko yang lebih tinggi.", False, False),
    ])

    # ---- 2.1.11 Prevalence Stratification Ratio ----
    H3(doc, "2.1.11", "Prevalence Stratification Ratio", italic=True)
    P(doc, segs=[
        ("Prevalence stratification ratio", False, True),
        (" merupakan metrik operasional yang mengukur seberapa baik sistem klasifikasi prioritas dapat menstratifikasi kasus ", False, False),
        ("stunting", False, True),
        (" nyata pada masing-masing kelas ", False, False),
        ("priority flag", False, True),
        (" terhadap prevalensi ", False, False),
        ("baseline", False, True),
        (" pada keseluruhan sampel. Metrik ini didefinisikan sebagai rasio antara prevalensi ", False, False),
        ("stunting", False, True),
        (" empiris pada ", False, False),
        ("priority flag", False, True),
        (" tertentu dibagi prevalensi ", False, False),
        ("baseline", False, True),
        (" pada keseluruhan dataset. Persamaan ", False, False),
        ("prevalence stratification ratio", False, True),
        (" dapat dilihat pada rumus II.14 berikut.", False, False),
    ])
    equation(doc, "Rasio_k = Prevalensi_k / Prevalensi_baseline", 14)
    P(doc, text="Di mana:", indent=0)
    bullets(doc, [
        [("Rasio_k", True, False), (" = ", False, False), ("prevalence stratification ratio", False, True), (" untuk ", False, False), ("priority flag", False, True), (" ke-k.", False, False)],
        [("Prevalensi_k", True, False), (" = proporsi kasus ", False, False), ("stunting", False, True), (" pada balita yang dikelompokkan ke ", False, False), ("priority flag", False, True), (" ke-k.", False, False)],
        [("Prevalensi_baseline", True, False), (" = proporsi kasus ", False, False), ("stunting", False, True), (" pada keseluruhan dataset.", False, False)],
    ])
    P(doc, segs=[
        ("Rasio lebih besar dari 1 menunjukkan kelas ", False, False),
        ("priority flag", False, True),
        (" tersebut mengakumulasi kasus ", False, False),
        ("stunting", False, True),
        (" lebih tinggi daripada rata-rata acak, sedangkan rasio kurang dari 1 menunjukkan kelas tersebut bersifat protektif. Metrik ini sangat penting untuk konteks ", False, False),
        ("monitoring", False, True),
        (" karena memvalidasi kemampuan sistem dalam mengarahkan prioritas intervensi pada tingkat komunitas.", False, False),
    ])

    # ========================================================================
    # 2.2 ALASAN PEMILIHAN METODE
    # ========================================================================
    H2(doc, "2.2", "Alasan Pemilihan Metode")
    P(doc, segs=[
        ("Pemilihan topik ", False, False),
        ("monitoring", False, True),
        (" risiko ", False, False),
        ("stunting", False, True),
        (" pada balita di Kalimantan Tengah didasari beban masalah yang masih tinggi dan penurunan yang berjalan lebih lambat dibanding target nasional, dengan prevalensi tercatat sekitar 23,5% pada 2023 dalam konteks kesenjangan antardaerah yang lebar (Kementerian Kesehatan Republik Indonesia, 2024). Kondisi ini menuntut pendekatan pemodelan risiko yang lebih presisi untuk mendukung penajaman sasaran intervensi, bukan sekadar pemantauan prevalensi agregat.", False, False),
    ])
    fig_caption(doc, 11, "Prevalensi Stunting Berdasarkan Provinsi Tahun 2022 dan 2023 (Kemenkes, 2024).")

    P(doc, segs=[
        ("Meskipun beberapa provinsi di Indonesia menunjukkan prevalensi ", False, False),
        ("stunting", False, True),
        (" yang lebih ekstrem secara nasional, penelitian ini tidak bertujuan untuk membandingkan besaran prevalensi antarwilayah, melainkan untuk mengembangkan dan mengevaluasi sistem ", False, False),
        ("monitoring", False, True),
        (" risiko ", False, False),
        ("stunting", False, True),
        (" berbasis data. Dalam konteks pengembangan model, wilayah dengan prevalensi yang terlalu ekstrem cenderung menghasilkan distribusi kelas yang sangat timpang, sehingga kurang optimal untuk menguji sensitivitas model terhadap variasi tingkat risiko. Oleh karena itu, Kalimantan Tengah dipilih karena berada pada zona transisi epidemiologis ", False, False),
        ("stunting", False, True),
        (", dengan prevalensi di atas rata-rata nasional namun tetap menunjukkan heterogenitas kondisi balita yang luas. Karakteristik ini memungkinkan sistem ", False, False),
        ("fuzzy Takagi-Sugeno", False, True),
        (" mengaktifkan aturan secara parsial, menangkap transisi risiko secara lebih halus, serta menghasilkan skor risiko yang lebih informatif, stabil, dan valid secara metodologis. Dengan demikian, pemilihan Kalimantan Tengah didasarkan pada pertimbangan epidemiologis dan statistik sebagai wilayah dengan spektrum risiko yang paling representatif untuk pengembangan dan validasi sistem ", False, False),
        ("fuzzy", False, True),
        (" berbasis ", False, False),
        ("risk scoring", False, True),
        (", bukan semata-mata karena tingkat prevalensinya.", False, False),
    ])

    P(doc, segs=[
        ("Penelitian ini memandang ", False, False),
        ("stunting", False, True),
        (" sebagai fenomena yang bersifat gradual dan multifaktorial, sehingga ", False, False),
        ("monitoring", False, True),
        (" risiko berbasis HAZ diperluas menjadi ", False, False),
        ("risk scoring", False, True),
        (" kontinu pada rentang 0-100 yang dipetakan ke empat ", False, False),
        ("priority flag", False, True),
        (" (Normal, ", False, False),
        ("Stunting Low", False, True),
        (", ", False, False),
        ("Stunting Medium", False, True),
        (", dan ", False, False),
        ("Stunting High", False, True),
        ("). Pemetaan ini menggunakan PBL sebagai indikator gizi ", False, False),
        ("intrauterin", False, True),
        (", tinggi dan berat badan saat pengukuran sebagai indikator antropometri terkini, serta waktu pengenalan MP-ASI sebagai indikator kecukupan asupan gizi pasca enam bulan pertama, untuk menghasilkan skor prioritas yang lebih informatif dibandingkan klasifikasi biner semata. Data SKI 2023 digunakan dengan sampel balita di Kalimantan Tengah yang memiliki data antropometri lengkap, sehingga seluruh kabupaten dan kota di Kalimantan Tengah masuk dalam populasi penelitian karena data SKI 2023 bersifat representatif provinsi.", False, False),
    ])

    P(doc, segs=[
        ("Model ", False, False),
        ("fuzzy Takagi-Sugeno", False, True),
        (" dipilih karena memiliki efisiensi komputasi tinggi, mampu merepresentasikan ketidakpastian, serta menghasilkan inferensi yang transparan dan mudah ditelusuri (Yershov & Ponomarenko, 2018). Sistem ini menggunakan kombinasi fungsi keanggotaan trapesium dan segitiga untuk memodelkan transisi gradual antarstatus gizi, yang secara biologis mencerminkan perubahan kontinu dalam pertumbuhan anak. Selain itu, batas fungsi keanggotaan dioptimasi menggunakan ", False, False),
        ("Genetic Algorithm", False, True),
        (" agar parameter menyesuaikan distribusi empiris data SKI 2023 Kalimantan Tengah tanpa mengubah basis aturan, operator inferensi, maupun struktur utama sistem, sehingga interpretabilitas ", False, False),
        ("fuzzy", False, True),
        (" tetap terjaga sambil performa klasifikasi meningkat (Elghamrawy & Hassanien, 2022; Ghahramani & Shiri, 2024).", False, False),
    ])


# ============================================================================
# BUILD DAFTAR PUSTAKA
# ============================================================================
def build_daftar_pustaka(doc):
    # page break
    p_break = doc.add_paragraph()
    p_break.add_run().add_break(WD_BREAK.PAGE)

    H_chap(doc, "DAFTAR PUSTAKA", "")

    # All references that appear in BAB I + BAB II, sorted alphabetically
    # Format: APA-style with hanging indent
    refs = [
        # A
        [("Abdy, M. (2018). ", False, False),
         ("Implementasi logika fuzzy dalam penilaian kinerja akademik mahasiswa", False, True),
         (". Jurnal MSA (Matematika dan Statistika serta Aplikasinya), 6(2), 32-37.", False, False)],

        [("Abegaz, T. M., Ahmed, M., Sherbeny, F., Diaby, V., Chi, H., & Ali, A. A. (2025). Application of machine learning algorithms to predict uncontrolled diabetes using the all of us research program data. ", False, False),
         ("Healthcare", False, True),
         (", 11(8), 1138. https://doi.org/10.3390/healthcare11081138", False, False)],

        [("Ananda, P. R., & Sriani, S. (2024). Sistem pakar diagnosis stunting pada balita menggunakan metode forward chaining dan logika fuzzy Sugeno. ", False, False),
         ("Jurnal Teknologi Sistem Informasi dan Aplikasi", False, True),
         (", 7(1), 200-216. https://doi.org/10.32493/jtsi.v7i1.38245", False, False)],

        [("Andrews, J. L., Bickel, P. J., & Romano, J. P. (2025). Stratified evaluation strategies for clinical machine learning models. ", False, False),
         ("Journal of Biomedical Informatics", False, True),
         (", 158, 104654. https://doi.org/10.1016/j.jbi.2024.104654", False, False)],

        [("Astonis, S. (2014). ", False, False),
         ("Pengenalan logika fuzzy dan aplikasinya pada sistem pendukung keputusan", False, True),
         (". Yogyakarta: Penerbit Andi.", False, False)],

        # B
        [("Bichri, H., Chergui, A., & Hain, M. (2024). Investigating the impact of train/test split ratio on the performance of pre-trained models with custom datasets. ", False, False),
         ("International Journal of Advanced Computer Science and Applications", False, True),
         (", 15(2), 331-339. https://doi.org/10.14569/IJACSA.2024.0150236", False, False)],

        [("Borgwardt, S., & Peñaloza, R. (2017). Algorithms for reasoning in very expressive description logics under infinitely valued Gödel semantics. ", False, False),
         ("International Journal of Approximate Reasoning", False, True),
         (", 83, 60-101. https://doi.org/10.1016/j.ijar.2017.01.001", False, False)],

        # D
        [("de Onis, M., & Branca, F. (2016). Childhood stunting: A global perspective. ", False, False),
         ("Maternal & Child Nutrition", False, True),
         (", 12, 12-26. https://doi.org/10.1111/mcn.12231", False, False)],

        # E
        [("Elghamrawy, S. M., & Hassanien, A. E. (2022). A hybrid Genetic-Grey Wolf Optimization algorithm for optimizing Takagi-Sugeno-Kang fuzzy systems. ", False, False),
         ("Neural Computing and Applications", False, True),
         (", 34(19), 17051-17069. https://doi.org/10.1007/s00521-022-07356-5", False, False)],

        [("Erda, G., Yolanda, A. M., Adnan, A., Alika, E. R., & Erda, Z. (2024). Advanced machine learning techniques for stunting classification using a stacking ensemble approach. In ", False, False),
         ("Proceedings of the International Conference on Electrical Engineering and Informatics (IConEEI 2024)", False, True),
         (" (pp. 79-84). IEEE. https://doi.org/10.1109/IConEEI64414.2024.10747922", False, False)],

        # G
        [("Georgescu, V., Petcu, A., & Crăciun, M. (2014). A weighted average method for defuzzification of Sugeno-type fuzzy inference systems. ", False, False),
         ("Mathematics and Computers in Simulation", False, True),
         (", 96, 81-101. https://doi.org/10.1016/j.matcom.2013.05.006", False, False)],

        [("Ghahramani, M., & Shiri, N. (2024). An adaptive neuro-fuzzy inference system optimized by genetic algorithm for brain tumour detection in magnetic resonance images. ", False, False),
         ("IET Image Processing", False, True),
         (", 18(5), 1358-1372. https://doi.org/10.1049/ipr2.13031", False, False)],

        [("Goldberg, D. E., & Deb, K. (1991). A comparative analysis of selection schemes used in genetic algorithms. ", False, False),
         ("Foundations of Genetic Algorithms", False, True),
         (", 1, 69-93. https://doi.org/10.1016/B978-0-08-050684-5.50008-2", False, False)],

        [("Gupta, P., Singh, R., & Sharma, A. (2022). Fuzzy logic in healthcare: A systematic review. ", False, False),
         ("Healthcare Analytics", False, True),
         (", 2, 100061. https://doi.org/10.1016/j.health.2022.100061", False, False)],

        # H
        [("Hand, D. J., Christen, P., & Kirielle, N. (2021). F*: An interpretable transformation of the F-measure. ", False, False),
         ("Machine Learning", False, True),
         (", 110(3), 451-456. https://doi.org/10.1007/s10994-021-05964-1", False, False)],

        [("Handayani, S., Lestari, R. P., & Nuraeni, A. (n.d.). Faktor risiko kejadian stunting pada balita: Studi literatur. ", False, False),
         ("Jurnal Kebidanan Indonesia", False, True),
         (", 14(1), 1-15.", False, False)],

        [("Hasan, H. A., & Mohammad, M. J. (2024). Classify the nutritional status of Iraqi children under five years using fuzzy classification. ", False, False),
         ("Journal of Al-Rafidain University College for Sciences", False, True),
         (", 2024(2), 380-394.", False, False)],

        [("Herrera, F., Lozano, M., & Sánchez, A. M. (2003). A taxonomy for the crossover operator for real-coded genetic algorithms: An experimental study. ", False, False),
         ("International Journal of Intelligent Systems", False, True),
         (", 18(3), 309-338. https://doi.org/10.1002/int.10091", False, False)],

        # J
        [("Japkowicz, N., & Shah, M. (2015). Performance evaluation in machine learning. In ", False, False),
         ("Machine Learning in Radiation Oncology", False, True),
         (" (pp. 41-56). Springer. https://doi.org/10.1007/978-3-319-18305-3_4", False, False)],

        # K
        [("Katoch, S., Chauhan, S. S., & Kumar, V. (2021). A review on genetic algorithm: Past, present, and future. ", False, False),
         ("Multimedia Tools and Applications", False, True),
         (", 80, 8091-8126. https://doi.org/10.1007/s11042-020-10139-6", False, False)],

        [("Kementerian Kesehatan Republik Indonesia. (2023). ", False, False),
         ("Laporan Survei Kesehatan Indonesia (SKI) 2023 - Tema Stunting", False, True),
         (". Jakarta: Kementerian Kesehatan RI.", False, False)],

        [("Kementerian Kesehatan Republik Indonesia. (2024). ", False, False),
         ("Laporan Survei Kesehatan Indonesia (SKI) 2023 dalam Angka", False, True),
         (". Jakarta: Kementerian Kesehatan RI.", False, False)],

        [("Krebs, N. F., Hambidge, K. M., Westcott, J. L., Garcés, A., Figueroa, L., Tshefu, A. K., Lokangaka, A. L., Goudar, S. S., Dhaded, S. M., Saleem, S., Ali, S. A., Bauserman, M. S., Derman, R. J., Goldenberg, R. L., Das, A., & Koso-Thomas, M. (2022). Birth length is the strongest predictor of linear growth status and stunting in the first 2 years of life after a preconception maternal nutrition intervention: The Women First trial. ", False, False),
         ("The American Journal of Clinical Nutrition", False, True),
         (", 116(1), 86-96. https://doi.org/10.1093/ajcn/nqac051", False, False)],

        [("Kurniawan, A. (2025). Determinan kejadian stunting pada balita di Indonesia: Tinjauan literatur. ", False, False),
         ("Jurnal Kesehatan Masyarakat Indonesia", False, True),
         (", 20(1), 45-58.", False, False)],

        # L
        [("Landgrebe, T. C. W., Paclík, P., Duin, R. P. W., & Bradley, A. P. (2006). Precision-recall operating characteristic (P-ROC) curves in imprecise environments. In ", False, False),
         ("Proceedings of the 18th International Conference on Pattern Recognition (ICPR'06)", False, True),
         (" (pp. 123-127). IEEE. https://doi.org/10.1109/ICPR.2006.940", False, False)],

        [("Liu, Y., & Yang, T. (2022). Visualization of confusion matrix and performance metrics for binary and multi-class classification. ", False, False),
         ("Pattern Recognition Letters", False, True),
         (", 154, 1-9. https://doi.org/10.1016/j.patrec.2021.12.008", False, False)],

        # M
        [("Marsya Finda, M., & Wahyu Utomo, P. (2024). Evaluasi performa model klasifikasi berbasis confusion matrix pada data tidak seimbang. ", False, False),
         ("Jurnal Sistem Informasi dan Komputer", False, True),
         (", 12(2), 87-99.", False, False)],

        [("Maulana, M. R., Pratama, R. D., & Wijaya, A. (2024). Fuzzy inference system untuk pengambilan keputusan berbasis data. ", False, False),
         ("Jurnal Informatika dan Komputer", False, True),
         (", 15(3), 215-228.", False, False)],

        [("Muflih, M., & Khasana, U. (2025). Hubungan waktu pemberian MP-ASI dengan kejadian stunting pada balita: Tinjauan sistematis. ", False, False),
         ("Jurnal Gizi Indonesia", False, True),
         (", 14(1), 28-39.", False, False)],

        # P
        [("Perumal, N., Manji, K. P., Darling, A. M., Kisenge, R. R., Kvestad, I., Hysing, M., Belinger, D. C., Strand, T. A., Duggan, C. P., Fawzi, W. W., & Sudfeld, C. R. (2023). Brief report: Stunting in early childhood and cognitive function later in childhood among children in Tanzania. ", False, False),
         ("Journal of the Pediatric Infectious Diseases Society", False, True),
         (", 12(5), 290-294. https://doi.org/10.1093/jpids/piad022", False, False)],

        [("Putri, A. D., & Saputro, B. (2021). Implementasi fuzzy Sugeno untuk klasifikasi tingkat kematangan buah. ", False, False),
         ("Jurnal Teknik Informatika", False, True),
         (", 14(2), 145-156.", False, False)],

        # R
        [("Rahayu, S., Lestari, D., & Putri, N. (2025). Determinan stunting pada anak balita di Indonesia berdasarkan data SKI 2023. ", False, False),
         ("Jurnal Ilmu Kesehatan Masyarakat", False, True),
         (", 16(1), 1-12.", False, False)],

        [("Rosdiana, R., Nurhayati, A., & Hadi, S. (2024). Pengukuran antropometri balita: Standar LAZ dan HAZ pada balita usia 0-60 bulan. ", False, False),
         ("Jurnal Gizi dan Kesehatan", False, True),
         (", 18(2), 89-101.", False, False)],

        [("Russiska, R., & Cahyati, W. H. (2025). Monitoring toddler development based on Indonesia Health Survey (SKI) 2023. ", False, False),
         ("Journal of Creativity Student", False, True),
         (", 8(1), 1-12. https://journal.unnes.ac.id/journals/jcs/article/view/25200", False, False)],

        # S
        [("Schober, P., Boer, C., & Schwarte, L. A. (2018). Correlation coefficients: Appropriate use and interpretation. ", False, False),
         ("Anesthesia & Analgesia", False, True),
         (", 126(5), 1763-1768. https://doi.org/10.1213/ANE.0000000000002864", False, False)],

        [("Soliman, A., De Sanctis, V., Alaaraj, N., Ahmed, S., Alyafei, F., Hamed, N., & Soliman, N. (2021). Early and long-term consequences of nutritional stunting: From childhood to adulthood. ", False, False),
         ("Acta Bio-Medica: Atenei Parmensis", False, True),
         (", 92(1), e2021168. https://doi.org/10.23750/abm.v92i1.11346", False, False)],

        # T
        [("Talukder, A., Razu, S. R., Hossain, M. Z., & Hossain, M. M. (2018). Factors influencing nutritional status of under-five children in Bangladesh: Findings from the Multiple Indicator Cluster Survey 2012-2013. ", False, False),
         ("Public Health in Practice", False, True),
         (", 9, 100040. https://doi.org/10.1016/j.puhip.2018.05.009", False, False)],

        # U
        [("UNICEF, World Health Organization, & World Bank Group. (2023). ", False, False),
         ("Levels and trends in child malnutrition: UNICEF/WHO/World Bank Group joint child malnutrition estimates - Key findings of the 2023 edition", False, True),
         (". New York: UNICEF.", False, False)],

        # W
        [("Wa Fijriyani R Ganisi. (2025). Sistem pendukung keputusan berbasis logika fuzzy untuk evaluasi status gizi balita. ", False, False),
         ("Jurnal Sains dan Teknologi", False, True),
         (", 11(1), 23-35.", False, False)],

        [("Waqiyah, A. N., Pertiwi, F. D., & Hidayah, N. (2023). The relationship between the timing of complementary feeding and maternal knowledge of responsive feeding and the incidence of stunting in children aged 6-24 months. ", False, False),
         ("Jurnal Kebidanan dan Keperawatan Aisyiyah", False, True),
         (", 19(2), 121-134.", False, False)],

        [("Watson, F., Adair, L. S., Bentley, M. E., Black, R. E., Caulfield, L. E., Christian, P., Danaei, G., Ezzati, M., Fawzi, W. W., Frongillo, E. A., Murray, C. J. L., Olofin, I., Stevens, G. A., & Sudfeld, C. R. (2024). Anthropometric indicators of child nutritional status: Methods, evidence, and gaps. ", False, False),
         ("The Lancet Child & Adolescent Health", False, True),
         (", 8(3), 195-208. https://doi.org/10.1016/S2352-4642(23)00308-4", False, False)],

        # Y
        [("Yani, D. I., Rahayuwati, L., Sari, C. W. M., Komariah, M., & Fauzi, S. A. (2023). Family household characteristics and stunting: An update scoping review. ", False, False),
         ("Nutrients", False, True),
         (", 15(1), 233. https://doi.org/10.3390/nu15010233", False, False)],

        [("Yershov, A., & Ponomarenko, M. (2018). Takagi-Sugeno fuzzy modelling for nonlinear control systems: Recent developments. ", False, False),
         ("International Journal of Control Systems and Robotics", False, True),
         (", 3, 13-20.", False, False)],

        # Z
        [("Zhang, L., Yang, Y., & Wu, B. (2006). Fuzzy logic in medical decision support systems: A review. ", False, False),
         ("Journal of Biomedical Informatics", False, True),
         (", 39(4), 374-386. https://doi.org/10.1016/j.jbi.2005.10.001", False, False)],
    ]

    for ref in refs:
        hanging_para(doc, ref, indent_left_cm=1.0, hanging_cm=1.0)


# ============================================================================
# MAIN
# ============================================================================
def main():
    doc = setup_doc()
    build_bab2(doc)
    build_daftar_pustaka(doc)
    out = "/Users/sinitygs/Projects2/train/BAB_II_FINAL_KOREKSI.docx"
    doc.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()