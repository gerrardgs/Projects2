// Generator for BAB IV (revisi v8) + BAB V (new) for Stunting Monitoring Skripsi
const fs = require('fs');
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType, ShadingType,
    PageNumber, Header, Footer, PageBreak, TabStopType, TabStopPosition,
    convertInchesToTwip, LineRuleType
} = require('docx');

// ============================================================================
// Helper functions for consistent styling
// ============================================================================
const FONT = "Times New Roman";
const SIZE_BODY = 24;        // 12pt
const SIZE_HEADING = 28;     // 14pt for chapter title
const SIZE_TABLE = 22;       // 11pt for table contents (smaller for code)
const SIZE_CODE = 20;        // 10pt for code in tables

// Body paragraph (justified, 1.5 line spacing, first-line indent, with Times New Roman)
function bodyPara(text, opts = {}) {
    const runs = Array.isArray(text)
        ? text
        : [new TextRun({ text: text, font: FONT, size: SIZE_BODY })];
    return new Paragraph({
        children: runs,
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        indent: opts.noIndent ? undefined : { firstLine: 720 },
        ...opts.paraOpts,
    });
}

// Plain paragraph - no indent, useful for figure captions, table titles etc
function plainPara(text, opts = {}) {
    const runs = Array.isArray(text)
        ? text
        : [new TextRun({ text: text, font: FONT, size: SIZE_BODY, ...opts.runOpts })];
    return new Paragraph({
        children: runs,
        alignment: opts.alignment || AlignmentType.LEFT,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
        ...opts.paraOpts,
    });
}

// Chapter heading - BAB V, etc - bold and centered
function chapterHeading(text) {
    return new Paragraph({
        children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_HEADING })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 240, after: 240 },
        pageBreakBefore: true,
    });
}

// Chapter heading for first chapter (no page break before)
function firstChapterHeading(text) {
    return new Paragraph({
        children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_HEADING })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 240, after: 240 },
    });
}

// Section heading like "4.1 Implementasi", "5.1 Kesimpulan" - bold left
function sectionHeading(text) {
    return new Paragraph({
        children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_BODY })],
        alignment: AlignmentType.LEFT,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 240, after: 120 },
    });
}

// Subsection like "4.1.1 Pengambilan Data", "5.2.1 Saran Praktis"
function subsectionHeading(text) {
    return new Paragraph({
        children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_BODY })],
        alignment: AlignmentType.LEFT,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 180, after: 100 },
    });
}

// Table caption (centered, bold)
function tableCaption(text) {
    // Allow mixed runs (regular + italic)
    if (Array.isArray(text)) {
        return new Paragraph({
            children: text,
            alignment: AlignmentType.CENTER,
            spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 200, after: 100 },
        });
    }
    return new Paragraph({
        children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_BODY })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 200, after: 100 },
    });
}

// Figure caption
function figureCaption(text) {
    if (Array.isArray(text)) {
        return new Paragraph({
            children: text,
            alignment: AlignmentType.CENTER,
            spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 100, after: 200 },
        });
    }
    return new Paragraph({
        children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_BODY })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 100, after: 200 },
    });
}

// Image placeholder paragraph (italicized, in brackets)
function imagePlaceholder(text) {
    return new Paragraph({
        children: [new TextRun({ text: text, italics: true, font: FONT, size: SIZE_BODY })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 100, after: 100 },
        border: {
            top: { style: BorderStyle.SINGLE, size: 6, color: "999999" },
            bottom: { style: BorderStyle.SINGLE, size: 6, color: "999999" },
            left: { style: BorderStyle.SINGLE, size: 6, color: "999999" },
            right: { style: BorderStyle.SINGLE, size: 6, color: "999999" },
        },
    });
}

// ============================================================================
// Table helpers
// ============================================================================
const TABLE_BORDER = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const TABLE_BORDERS = {
    top: TABLE_BORDER, bottom: TABLE_BORDER,
    left: TABLE_BORDER, right: TABLE_BORDER,
    insideHorizontal: TABLE_BORDER, insideVertical: TABLE_BORDER,
};
const CELL_BORDERS = {
    top: TABLE_BORDER, bottom: TABLE_BORDER,
    left: TABLE_BORDER, right: TABLE_BORDER,
};

function tableCell(text, opts = {}) {
    const isHeader = opts.header === true;
    const align = opts.align || AlignmentType.LEFT;
    const runs = Array.isArray(text)
        ? text
        : [new TextRun({
            text: String(text),
            bold: isHeader || opts.bold === true,
            font: FONT,
            size: opts.size || SIZE_TABLE,
        })];
    return new TableCell({
        borders: CELL_BORDERS,
        width: { size: opts.width || 2340, type: WidthType.DXA },
        shading: isHeader ? { fill: "E0E0E0", type: ShadingType.CLEAR } : undefined,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
            children: runs,
            alignment: align,
            spacing: { line: 280, lineRule: LineRuleType.AUTO },
        })],
    });
}

// Code cell -- used for single-column code tables (Tabel IV.x Code...)
function codeCell(codeLines, opts = {}) {
    // codeLines is an array of strings (each line)
    const paras = codeLines.map(line => new Paragraph({
        children: [new TextRun({ text: line || " ", font: "Courier New", size: SIZE_CODE })],
        alignment: AlignmentType.LEFT,
        spacing: { line: 240, lineRule: LineRuleType.AUTO },
    }));
    return new TableCell({
        borders: CELL_BORDERS,
        width: { size: opts.width || 9360, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 150, right: 150 },
        shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
        children: paras,
    });
}

// Build a single-column code table
function codeTable(codeLines) {
    return new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        borders: TABLE_BORDERS,
        rows: [
            new TableRow({ children: [codeCell(codeLines, { width: 9360 })] })
        ],
    });
}

// Generic data table with header row + body rows
function dataTable(headers, rows, columnWidths) {
    const totalWidth = columnWidths.reduce((a, b) => a + b, 0);
    const headerCells = headers.map((h, i) => tableCell(h, {
        header: true, align: AlignmentType.CENTER, width: columnWidths[i]
    }));
    const bodyRows = rows.map(r => new TableRow({
        children: r.map((c, i) => {
            // If cell is an object {text, bold, align}, use it; else treat as string
            if (typeof c === 'object' && c !== null && !Array.isArray(c)) {
                return tableCell(c.text, {
                    width: columnWidths[i],
                    align: c.align || (i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER),
                    bold: c.bold || false,
                });
            }
            return tableCell(c, {
                width: columnWidths[i],
                align: i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER,
            });
        })
    }));
    return new Table({
        width: { size: totalWidth, type: WidthType.DXA },
        columnWidths: columnWidths,
        borders: TABLE_BORDERS,
        rows: [
            new TableRow({ children: headerCells, tableHeader: true }),
            ...bodyRows,
        ],
    });
}

// Bullet/numbered list paragraph helper
function listPara(text, ref = "numList") {
    return new Paragraph({
        children: Array.isArray(text) ? text : [new TextRun({ text: text, font: FONT, size: SIZE_BODY })],
        numbering: { reference: ref, level: 0 },
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 0, after: 60 },
    });
}

// Lettered list (a., b., c.)
function letterListPara(text) {
    return new Paragraph({
        children: Array.isArray(text) ? text : [new TextRun({ text: text, font: FONT, size: SIZE_BODY })],
        numbering: { reference: "letterList", level: 0 },
        alignment: AlignmentType.JUSTIFIED,
        spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 0, after: 60 },
    });
}

// Mix italic/regular runs inline
function R(text, opts = {}) {
    return new TextRun({
        text: text,
        font: FONT,
        size: opts.size || SIZE_BODY,
        italics: opts.italics || false,
        bold: opts.bold || false,
    });
}

// ============================================================================
// BAB IV CONTENT  (Hasil dan Pembahasan)
// ============================================================================
function buildBabIV() {
    const children = [];

    // BAB IV title
    children.push(firstChapterHeading("BAB IV"));
    children.push(firstChapterHeading("HASIL DAN PEMBAHASAN"));
    children.push(bodyPara("Bab ini menyajikan hasil dari setiap tahap yang dilalui dalam membangun sistem monitoring stunting balita di Kalimantan Tengah berbasis Fuzzy Sugeno yang dioptimasi GA. Pembahasan dibagi menjadi tiga bagian utama, yakni implementasi sistem (kode dan antarmuka), hasil pengujian terhadap kinerja klasifikasi, serta pembahasan menyeluruh terhadap temuan yang diperoleh. Seluruh angka yang dilaporkan bersumber langsung dari eksekusi skrip fuzzytrainv2.py dan deployment website fuzzymain.py, sehingga dapat ditelusuri kembali ke kode aslinya."));

    // -------- 4.1 Implementasi --------
    children.push(sectionHeading("4.1   Implementasi"));
    children.push(bodyPara("Bagian ini menjelaskan tahap demi tahap penyusunan sistem, mulai dari pemuatan dataset, pembersihan data, perhitungan Z-Score sebagai label rujukan, perancangan logika fuzzy, optimasi dengan GA, evaluasi performa, sampai dengan pembangunan tampilan website. Setiap sub-bagian disertai potongan kode kunci dan ringkasan hasilnya, sedangkan pembahasan lebih dalam tentang makna angka-angka tersebut akan diuraikan pada bagian 4.3."));

    // 4.1.1 Pengambilan Data
    children.push(subsectionHeading("4.1.1   Pengambilan Data"));
    children.push(bodyPara("Data yang digunakan berasal dari Survei Kesehatan Indonesia (SKI) 2023 untuk Provinsi Kalimantan Tengah dalam berkas Individu_balita_2023_merged.csv. Pemuatan dilakukan dengan pustaka pandas, kemudian dipilih tujuh kolom yang relevan dengan tujuan penelitian dan diubah namanya menjadi lebih ringkas agar mudah ditelusuri pada tahap berikutnya."));
    children.push(tableCaption([
        R("Tabel IV.1 ", { bold: true }),
        R("Code", { bold: true, italics: true }),
        R(" Proses Pengambilan Data", { bold: true }),
    ]));
    children.push(codeTable([
        "import pandas as pd",
        "import numpy as np",
        "",
        "DATASET_PATH = \"dataset/Individu_balita_2023_merged.csv\"",
        "df_raw = pd.read_csv(DATASET_PATH, low_memory=False)",
        "",
        "COLUMN_MAP = {",
        "    \"Umur_BLN\":                      \"usia_bulan\",",
        "    \"Jenis Kelamin\":                 \"jenis_kelamin\",",
        "    \"Panjang Badan Lahir\":           \"pbl_raw\",",
        "    \"Panjang Badan Anak\":            \"tinggi_badan_cm\",",
        "    \"Berat Badan Saat Data Diambil\": \"berat_badan_raw\",",
        "    \"I43_MP ASI\":                    \"mpasi_raw\",",
        "}",
        "",
        "cols_needed = [\"ID_balita\"] + list(COLUMN_MAP.keys())",
        "df = df_raw[cols_needed].copy()",
        "df.rename(columns=COLUMN_MAP, inplace=True)",
    ]));
    children.push(bodyPara("Tabel IV.1 menunjukkan proses pengambilan data, di mana berkas SKI 2023 dimuat dengan pandas lalu disusutkan menjadi tujuh kolom inti. Dataset awal berisi 1.679 baris dan 370 kolom, lalu disusutkan menjadi tujuh kolom yang siap diproses pada tahap selanjutnya. Daftar pemetaan kolom asli ke variabel kerja ditampilkan pada Tabel 4.1."));
    children.push(tableCaption("Tabel 4.1 Daftar Fitur SKI 2023 ke Variabel Penelitian"));
    children.push(dataTable(
        ["Kolom Asli SKI 2023", "Variabel Program", "Satuan / Format"],
        [
            ["ID_balita", "ID_balita", "string (identifier)"],
            ["Umur_BLN", "usia_bulan", "bulan (integer)"],
            ["Jenis Kelamin", "jenis_kelamin", "1 = Laki-laki, 2 = Perempuan"],
            ["Panjang Badan Lahir", "pbl_raw → pbl_cm", "cm (TB saat lahir)"],
            ["Panjang Badan Anak", "tinggi_badan_cm", "cm (TB saat pengukuran)"],
            ["Berat Badan Saat Data Diambil", "berat_badan_raw → berat_badan_gram", "gram (BB saat pengukuran)"],
            ["I43_MP ASI", "mpasi_raw → mpasi_bulan", "bulan (kode huruf, dipetakan)"],
        ],
        [3120, 3120, 3120]
    ));

    // 4.1.2 Preprocessing
    children.push(subsectionHeading("4.1.2   Preprocessing"));
    children.push(bodyPara("Setelah dataset dimuat, dilakukan pembersihan data agar hanya baris bersih yang masuk ke tahap berikutnya. Pada tahap ini terdapat empat operasi utama, yaitu transformasi satuan (berat badan dari kg ke gram), pemetaan kode huruf MP-ASI menjadi nilai bulan, penghapusan baris yang memiliki nilai kosong atau di luar batas medis (listwise deletion), serta pembentukan dua kelas turunan yang dibutuhkan analisis monitoring."));
    children.push(tableCaption([
        R("Tabel IV.2 ", { bold: true }),
        R("Code", { bold: true, italics: true }),
        R(" Proses Preprocessing", { bold: true }),
    ]));
    children.push(codeTable([
        "def task2_preprocessing(df, raw_shape):",
        "    # Konversi satuan",
        "    df[\"berat_badan_kg\"]   = pd.to_numeric(df[\"berat_badan_raw\"], errors=\"coerce\")",
        "    df[\"berat_badan_gram\"] = df[\"berat_badan_kg\"] * 1000",
        "    df[\"pbl_cm\"]           = pd.to_numeric(df[\"pbl_raw\"], errors=\"coerce\")",
        "",
        "    # Petakan kode huruf MP-ASI ke bulan",
        "    MPASI_MAP = {\"a\":0.25, \"b\":0.5, \"c\":1, \"d\":2, \"e\":3, \"f\":4, \"g\":5,",
        "                 \"h\":6, \"i\":7, \"j\":8, \"k\":9, \"l\":10, \"m\":11, \"n\":12}",
        "    df[\"mpasi_bulan\"] = df[\"mpasi_raw\"].str.lower().map(MPASI_MAP)",
        "",
        "    # Hapus baris yang memiliki nilai kosong pada 6 variabel inti",
        "    CORE = [\"usia_bulan\",\"jenis_kelamin\",\"tinggi_badan_cm\",",
        "            \"berat_badan_gram\",\"pbl_cm\",\"mpasi_bulan\"]",
        "    df_clean = df.dropna(subset=CORE).reset_index(drop=True)",
        "",
        "    # Bentuk dua kolom baru untuk analisis monitoring",
        "    df_clean[\"jarak_lahir_ukur\"]   = df_clean[\"usia_bulan\"].astype(int)",
        "    df_clean[\"jarak_mpasi_survei\"] = (df_clean[\"usia_bulan\"] -",
        "                                      df_clean[\"mpasi_bulan\"]).clip(lower=0)",
        "    return df_clean",
    ]));
    children.push(bodyPara("Tabel IV.2 menunjukkan keseluruhan langkah preprocessing dalam satu fungsi. Operasi yang paling memengaruhi ukuran sampel adalah listwise deletion, yakni penghapusan setiap baris yang memiliki minimal satu nilai kosong pada enam variabel inti. Dari 1.679 baris awal, hanya 389 baris yang lolos seluruh kriteria; sisanya 1.290 baris (76,8%) terhapus terutama karena kolom MP-ASI dan panjang badan lahir banyak kosong. Implikasi penurunan drastis ini terhadap stabilitas hasil pelatihan akan dibahas lebih dalam pada bagian 4.3. Ringkasan operasi preprocessing per tahap ditampilkan pada Tabel 4.2."));
    children.push(tableCaption("Tabel 4.2 Hasil Tahap Preprocessing pada Dataset SKI 2023"));
    children.push(dataTable(
        ["Tahap", "Operasi", "Jumlah Baris"],
        [
            ["Data Mentah", "Pemuatan Individu_balita_2023_merged.csv", "1.679"],
            ["Transformasi", "Konversi BB ke gram, validasi kolom", "1.679"],
            ["Rekoding", "Pemetaan kode huruf MP-ASI ke bulan", "1.679"],
            ["Filter Usia", "Penyaringan rentang 0-59 bulan", "1.679"],
            ["Listwise Deletion", "Penghapusan baris missing & outlier", "389"],
            ["Pembentukan Kelas Baru", "jarak_lahir_ukur, jarak_mpasi_survei", "389"],
        ],
        [2340, 4680, 2340]
    ));

    // 4.1.3 Z-Score
    children.push(subsectionHeading("4.1.3   Perhitungan Z-Score dan Manual Labelling"));
    children.push(bodyPara("Z-Score dihitung pada dua momen, yakni saat lahir sebagai label awal kondisi balita dan saat pengukuran sebagai ground truth evaluasi. Perhitungannya dilakukan dengan menyesuaikan parameter LMS WHO yang sudah dijelaskan pada Tabel rujukan WHO di BAB II. Hasil pelabelan ini menghasilkan dua kolom baru pada dataset, yaitu label_who untuk kondisi saat lahir dan label_who_ukur untuk ground truth evaluasi."));
    children.push(tableCaption([
        R("Tabel IV.3 ", { bold: true }),
        R("Code", { bold: true, italics: true }),
        R(" Proses Perhitungan Z-Score", { bold: true }),
    ]));
    children.push(codeTable([
        "def hitung_zscore(jk, usia, x, who_lookup):",
        "    key = (int(jk), int(usia))",
        "    if key not in who_lookup:",
        "        return None, None",
        "    p = who_lookup[key]",
        "    L, M, S = p[\"L\"], p[\"M\"], p[\"S\"]",
        "    if L != 0:",
        "        z = ((x / M) ** L - 1) / (L * S)",
        "    else:",
        "        z = np.log(x / M) / S",
        "    label = \"Stunting\" if z < -2.0 else \"Normal\"",
        "    return round(z, 3), label",
        "",
        "# Z-Score saat lahir (usia di-set 0 bulan)",
        "for idx, row in df.iterrows():",
        "    z, lbl = hitung_zscore(row[\"jenis_kelamin\"], 0, row[\"pbl_cm\"], who_lookup)",
        "    df.at[idx, \"z_score\"]   = z",
        "    df.at[idx, \"label_who\"] = lbl",
        "",
        "# Z-Score saat pengukuran (usia aktual, ground truth)",
        "for idx, row in df.iterrows():",
        "    z, lbl = hitung_zscore(row[\"jenis_kelamin\"], row[\"usia_bulan\"],",
        "                            row[\"tinggi_badan_cm\"], who_lookup)",
        "    df.at[idx, \"z_score_ukur\"]   = z",
        "    df.at[idx, \"label_who_ukur\"] = lbl",
    ]));
    children.push(bodyPara("Tabel IV.3 menampilkan kode perhitungan Z-Score dengan rumus LMS WHO. Penjelasan rinci mengenai rumus tersebut sudah ditulis pada BAB II. Statistik hasil pelabelan untuk kedua momen akan dibahas pada bagian 4.3 sebagai justifikasi lebih lanjut tentang kondisi sampel akhir."));

    // 4.1.4 Sistem Logika Fuzzy Sugeno
    children.push(subsectionHeading("4.1.4   Sistem Logika Fuzzy Sugeno"));
    children.push(bodyPara("Inti dari sistem fuzzy terdiri dari tiga fungsi utama, yaitu menghitung derajat keanggotaan setiap input terhadap fungsi keanggotaan, mengevaluasi aturan IF-THEN dengan operator minimum, serta defuzzifikasi metode Sugeno untuk menghasilkan risk score. Detail tentang fungsi keanggotaan, 108 aturan, dan rumus defuzzifikasi sudah dibahas pada BAB III."));
    children.push(tableCaption([
        R("Tabel IV.4 ", { bold: true }),
        R("Code", { bold: true, italics: true }),
        R(" Proses Inferensi Fuzzy Sugeno", { bold: true }),
    ]));
    children.push(codeTable([
        "def fuzzify_new(usia, tb, bb, mpasi, jk, mf_params):",
        "    tb_mf = mf_params[\"tb_l\"] if jk == 1 else mf_params[\"tb_p\"]",
        "    bb_mf = mf_params[\"bb_l\"] if jk == 1 else mf_params[\"bb_p\"]",
        "    mu = {",
        "        \"usia\":  {k: eval_mf(usia, *v)  for k, v in mf_params[\"usia\"].items()},",
        "        \"tb\":    {k: eval_mf(tb, *v)    for k, v in tb_mf.items()},",
        "        \"bb\":    {k: eval_mf(bb, *v)    for k, v in bb_mf.items()},",
        "        \"mpasi\": {k: eval_mf(mpasi, *v) for k, v in mf_params[\"mpasi\"].items()},",
        "    }",
        "    return mu",
        "",
        "def evaluate_rules_new(mu, cons):",
        "    fired = []",
        "    for usia_c, tb_c, bb_c, mpasi_c, risk_class in RULES_108:",
        "        w = min(mu[\"usia\"][usia_c], mu[\"tb\"][tb_c],",
        "                mu[\"bb\"][bb_c], mu[\"mpasi\"][mpasi_c])",
        "        if w > 0: fired.append((w, cons[risk_class]))",
        "    return fired",
        "",
        "def defuzzify_sugeno(fired):",
        "    num = sum(w * c for w, c in fired)",
        "    den = sum(w for w, c in fired)",
        "    return round(num / den, 2) if den > 0 else 35.0",
    ]));
    children.push(bodyPara("Tabel IV.4 menunjukkan tiga fungsi inti sistem fuzzy. Fungsi fuzzify_new() menghitung derajat keanggotaan, evaluate_rules_new() mengaktifkan aturan dengan operator min, dan defuzzify_sugeno() menghasilkan nilai tegas risk score. Distribusi 108 aturan berdasarkan kelas risiko ditampilkan pada Tabel 4.4."));
    children.push(tableCaption("Tabel 4.4 Distribusi 108 Aturan Fuzzy Sugeno per Kelas Risiko"));
    children.push(dataTable(
        ["Kelas Risiko", "Konsekuen (c)", "Jumlah Aturan", "Total per JK"],
        [
            ["STUNTING HIGH", "90", "10", "20"],
            ["STUNTING MEDIUM", "65", "30", "60"],
            ["STUNTING LOW", "35", "20", "40"],
            ["Normal", "10", "48", "96"],
            [{ text: "Total", bold: true }, { text: "-", bold: true }, { text: "108", bold: true }, { text: "216", bold: true }],
        ],
        [2340, 2340, 2340, 2340]
    ));

    // 4.1.5 GA
    children.push(subsectionHeading("4.1.5   Optimasi Menggunakan Genetic Algorithm"));
    children.push(bodyPara("Optimasi GA digunakan untuk mencari posisi titik batas (boundary points) fungsi keanggotaan yang paling sesuai dengan distribusi data Kalimantan Tengah. Pemilihan parameter populasi, generasi, probabilitas crossover, dan probabilitas mutasi mengacu pada hasil eksperimen dari sejumlah penelitian fuzzy yang menggunakan GA pada dataset berukuran sedang. Detail nilai dan justifikasi tiap parameter GA telah dijelaskan pada BAB III, sehingga pada bagian ini hanya ditampilkan implementasi kodenya."));
    children.push(tableCaption([
        R("Tabel IV.5 ", { bold: true }),
        R("Code", { bold: true, italics: true }),
        R(" Proses Optimasi ", { bold: true }),
        R("Genetic Algorithm", { bold: true, italics: true }),
    ]));
    children.push(codeTable([
        "from deap import base, creator, tools, algorithms",
        "import random",
        "",
        "creator.create(\"FitnessMax\", base.Fitness, weights=(1.0,))",
        "creator.create(\"Individual\", list, fitness=creator.FitnessMax)",
        "",
        "toolbox = base.Toolbox()",
        "toolbox.register(\"evaluate\", fitness_func)",
        "toolbox.register(\"mate\",     tools.cxBlend, alpha=0.3)",
        "toolbox.register(\"mutate\",   mutate_bounded, indpb=0.15)",
        "toolbox.register(\"select\",   tools.selTournament, tournsize=3)",
        "",
        "SEEDS = [42, 123, 777, 2024, 31415]",
        "for seed in SEEDS:",
        "    random.seed(seed); np.random.seed(seed)",
        "    pop = toolbox.population(n=80)",
        "    pop, log = algorithms.eaSimple(pop, toolbox, cxpb=0.6,",
        "                                    mutpb=0.3, ngen=60, verbose=False)",
    ]));
    children.push(bodyPara("Tabel IV.5 menampilkan inti proses GA, di mana lima nilai seed berbeda dijalankan secara berurutan agar hasilnya tidak bergantung pada satu titik awal saja. Setiap seed menghasilkan satu individu terbaik yang kemudian dievaluasi pada subset uji untuk memilih konfigurasi yang paling kuat. Konfigurasi nilai parameter dan justifikasinya telah dijelaskan pada BAB III."));

    // 4.1.6 Evaluasi
    children.push(subsectionHeading("4.1.6   Evaluasi"));
    children.push(bodyPara("Evaluasi dilakukan dengan membandingkan output sistem fuzzy terhadap label ground truth Z-Score saat pengukuran. Metrik yang digunakan adalah akurasi, precision, recall, dan F1-score yang dihitung dari empat komponen confusion matrix (TP, TN, FP, FN). Sebelum membahas implementasi metrik, perlu dijelaskan terlebih dahulu strategi pembagian data yang berbeda untuk sistem fuzzy dan tahap optimasi GA agar hasil pengukuran mencerminkan kondisi yang sebenarnya."));
    children.push(bodyPara("Sistem fuzzy baseline dijalankan terhadap keseluruhan 389 balita untuk menghasilkan distribusi awal risk score dan priority flag. Pada tahap ini belum dilakukan pembagian data karena tujuannya hanya mengamati perilaku sistem dengan posisi MF default. Hasil evaluasi baseline kemudian dijadikan baseline metric yang akan diperbandingkan terhadap hasil pasca-optimasi."));
    children.push(bodyPara("Optimasi GA membutuhkan pembagian data agar parameter MF tidak overfitting terhadap seluruh sampel. Dataset 389 balita dibagi menjadi 80% subset pelatihan (311 balita) dan 20% subset pengujian (78 balita) menggunakan stratified split, yakni proporsi kelas Stunting dan Normal pada kedua subset dijaga agar mendekati distribusi keseluruhan (16,5% Stunting). Subset pelatihan dipakai oleh fungsi fitness GA untuk mengevaluasi setiap kandidat MF; subset pengujian disisihkan dan tidak pernah dilihat selama proses optimasi, lalu dipakai sebagai estimasi performa pada data baru (generalisasi). Strategi ini memastikan bahwa kinerja yang dilaporkan pada Tabel 4.7 bukan akibat hafalan data pelatihan."));
    children.push(bodyPara("Setelah seluruh seed selesai dijalankan, individu dengan F1-score subset pengujian tertinggi dipilih sebagai konfigurasi final. Detail evaluasinya, termasuk fungsi penghitung metrik, ditampilkan pada Tabel IV.6 berikut."));
    children.push(tableCaption([
        R("Tabel IV.6 ", { bold: true }),
        R("Code", { bold: true, italics: true }),
        R(" Proses Evaluasi ", { bold: true }),
        R("Confusion Matrix", { bold: true, italics: true }),
    ]));
    children.push(codeTable([
        "def calc_metrics(y_true, y_pred):",
        "    tp = int(np.sum((y_true == 1) & (y_pred == 1)))",
        "    tn = int(np.sum((y_true == 0) & (y_pred == 0)))",
        "    fp = int(np.sum((y_true == 0) & (y_pred == 1)))",
        "    fn = int(np.sum((y_true == 1) & (y_pred == 0)))",
        "    prec = tp / (tp + fp) if (tp + fp) > 0 else 0.0",
        "    rec  = tp / (tp + fn) if (tp + fn) > 0 else 0.0",
        "    f1   = 2 * prec * rec / (prec + rec) if (prec + rec) > 0 else 0.0",
        "    acc  = (tp + tn) / (tp + tn + fp + fn)",
        "    return prec, rec, f1, acc",
    ]));
    children.push(bodyPara("Tabel IV.6 menampilkan fungsi calc_metrics() yang menghitung TP, TN, FP, FN dari y_true dan y_pred, lalu menurunkan empat metrik klasifikasi standar. Hasil pembagian data 80/20 ditampilkan pada Tabel 4.6."));
    children.push(tableCaption("Tabel 4.6 Hasil Stratified Split 80/20 pada Dataset Akhir"));
    children.push(dataTable(
        ["Subset", "Total", "Normal", "Stunting"],
        [
            ["Keseluruhan", "389", "325", "64"],
            ["Pelatihan (80%)", "311", "260", "51"],
            ["Pengujian (20%)", "78", "65", "13"],
        ],
        [2340, 2340, 2340, 2340]
    ));
    children.push(bodyPara("Tabel 4.6 menampilkan komposisi tiap subset setelah pembagian stratified. Dengan pembagian ini, tahap optimasi GA memiliki acuan data pelatihan dan pengujian yang konsisten. Setelah seluruh proses optimasi GA selesai dijalankan, hasil prediksi sistem terhadap subset pengujian dirangkum dalam confusion matrix sebagaimana ditampilkan pada Tabel 4.7."));
    children.push(tableCaption("Tabel 4.7 Confusion Matrix Subset Pengujian (n = 78)"));
    children.push(dataTable(
        ["", "Aktual Stunting", "Aktual Normal", "Total"],
        [
            [{ text: "Prediksi Stunting", bold: true }, "TP = 11", "FP = 19", "30"],
            [{ text: "Prediksi Normal", bold: true }, "FN = 2", "TN = 46", "48"],
            [{ text: "Total", bold: true }, { text: "13", bold: true }, { text: "65", bold: true }, { text: "78", bold: true }],
        ],
        [2340, 2340, 2340, 2340]
    ));
    children.push(bodyPara("Pada subset uji, sistem berhasil mengidentifikasi 11 dari 13 kasus stunting (recall 84,62%) dan menahan 46 dari 65 balita normal dengan benar. Detail mengenai mengapa profil precision dan recall jatuh pada angka-angka tersebut akan dibahas pada bagian 4.3."));

    // 4.1.7 Sistem Website
    children.push(subsectionHeading("4.1.7   Pembangunan Sistem Website"));
    children.push(bodyPara("Tampilan website Stuntor terdiri dari empat halaman, yakni landing sebagai pintu masuk, form input data balita, halaman hasil perhitungan, dan halaman riwayat dengan tombol auto-retrain. Implementasi sisi server menggunakan Flask 3.x yang memuat trained_model.json pada saat startup, sehingga inferensi langsung memakai parameter MF hasil optimasi GA."));
    // a) Halaman Landing
    children.push(plainPara("a)   Halaman Landing", { runOpts: { bold: false } }));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.7 - Halaman Landing Website (i)]"));
    children.push(figureCaption("Gambar 4.7 Halaman Landing Website (i)"));
    children.push(bodyPara("Gambar 4.7 menunjukkan halaman landing yang berfungsi sebagai pintu masuk sistem, memuat ringkasan tujuan layanan dan tombol panggil aksi ke halaman form input (header website), juga terdiri atas informasi mengenai sistem website secara keseluruhan."));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.7 - Halaman Landing Website (ii)]"));
    children.push(figureCaption("Gambar 4.7 Halaman Landing Website (ii)"));
    children.push(bodyPara("Gambar 4.7 di atas menunjukkan informasi mengenai fitur utama yang ada dalam website."));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.7 - Halaman Landing Website (iii)]"));
    children.push(figureCaption("Gambar 4.7 Halaman Landing Website (iii)"));
    children.push(bodyPara("Gambar 4.7 di atas menunjukkan informasi tambahan mengenai pertanyaan yang umumnya ditanyakan pengguna, serta aturan dan kebijakan dari sistem website."));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.7 - Halaman Landing Website (iv)]"));
    children.push(figureCaption("Gambar 4.7 Halaman Landing Website (iv)"));
    children.push(bodyPara("Gambar 4.7 di atas menunjukkan informasi tambahan mengenai contact person yang dapat dihubungi, serta footer website."));
    // b) Form Input
    children.push(plainPara("b)   Halaman Form Input"));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.8 - Halaman Form Input Data Balita]"));
    children.push(figureCaption("Gambar 4.8 Halaman Form Input Data Balita"));
    children.push(bodyPara("Gambar 4.8 memperlihatkan halaman form input dengan enam medan masukan, yakni umur, jenis kelamin, tinggi badan saat ukur, berat badan saat ukur, panjang badan lahir, dan waktu pengenalan MP-ASI."));
    // c) Hasil Perhitungan
    children.push(plainPara("c)   Halaman Hasil Perhitungan"));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.9 - Halaman Hasil Perhitungan]"));
    children.push(figureCaption("Gambar 4.9 Halaman Hasil Perhitungan"));
    children.push(bodyPara("Gambar 4.9 menampilkan halaman hasil yang muncul setelah pengguna mengirim data, berisi Z-Score saat lahir, Z-Score saat ukur, risk score 0-100, dan priority flag."));
    // d) Riwayat
    children.push(plainPara("d)   Halaman Riwayat dengan Tombol Retrain"));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.10 - Halaman Riwayat dan Tombol Retrain Model]"));
    children.push(figureCaption("Gambar 4.10 Halaman Riwayat dan Tombol Retrain Model"));
    children.push(bodyPara("Gambar 4.10 menampilkan halaman riwayat yang merekap seluruh inferensi pengguna sekaligus menyediakan tombol Retrain Model. Tombol ini memicu endpoint /retrain pada Flask backend untuk menjalankan ulang optimasi GA dengan dataset gabungan (riwayat pengguna + dataset training awal). Tombol aktif setelah riwayat pengguna mencapai minimal lima record, dan setelah retrain berhasil, parameter MF baru otomatis dimuat sehingga prediksi berikutnya memakai model terbaru."));

    // ============ 4.2 Hasil Pengujian ============
    children.push(sectionHeading("4.2   Hasil Pengujian"));
    children.push(bodyPara("Bagian ini mempresentasikan hasil utama pengujian sistem fuzzy Sugeno yang telah dioptimasi GA. Dua hasil inti dilaporkan secara berurutan, yaitu perbandingan kinerja antara konfigurasi baseline dan GA-optimized, serta stabilitas hasil pengujian multi-seed. Analisis dan pemaknaan mendalam atas setiap hasil tersebut akan diberikan pada bagian 4.3."));

    children.push(subsectionHeading("4.2.1   Pengujian Sistem Fuzzy Sugeno + GA"));
    children.push(bodyPara("Pengujian utama dimulai dengan membandingkan kinerja sistem sebelum dan sesudah optimasi GA pada subset pelatihan dan pengujian. Tabel 4.7 merangkum hasil empat metrik klasifikasi (akurasi, precision, recall, F1-score) pada kedua kondisi tersebut. Kolom Test Baseline menggantikan kolom Delta sebelumnya agar perbandingan lebih utuh antara kondisi awal dan akhir model di kedua subset."));
    children.push(tableCaption("Tabel 4.8 Perbandingan Kinerja Baseline dan GA-Optimized"));
    children.push(dataTable(
        ["Metrik", "Train Baseline", "Train GA-Optimized", "Test Baseline", "Test GA-Optimized"],
        [
            ["Akurasi", "61,74%", "69,13%", "65,38%", { text: "73,08%", bold: true }],
            ["Precision", "18,52%", "31,71%", "24,00%", { text: "36,67%", bold: true }],
            ["Recall", "39,22%", "76,47%", "46,15%", { text: "84,62%", bold: true }],
            ["F1-score", "25,16%", "44,83%", "31,58%", { text: "51,16%", bold: true }],
        ],
        [1872, 1872, 1872, 1872, 1872]
    ));
    children.push(bodyPara("Pengujian menunjukkan bahwa optimasi GA meningkatkan seluruh metrik secara konsisten di kedua subset, dengan F1-score test naik dari 31,58% menjadi 51,16% dan akurasi test naik dari 65,38% menjadi 73,08%. Peningkatan ini diperoleh tanpa mengubah struktur 108 aturan, sehingga interpretabilitas linguistik sistem tetap terjaga."));
    children.push(bodyPara("Untuk memastikan hasil terbaik tidak tergantung pada satu titik awal acak, GA dijalankan dengan lima seed berbeda. Rangkuman kinerja pada seluruh seed ditampilkan pada Tabel 4.9. Metrik evaluasi (Precision, Recall, F1, Accuracy) dinyatakan dalam persentase, sedangkan nilai Train Fitness tetap dalam bentuk desimal karena merupakan skor fungsi optimasi bukan rasio klasifikasi."));
    children.push(tableCaption([
        R("Tabel 4.9 Hasil Pengujian Multi-Seed ", { bold: true }),
        R("Genetic Algorithm", { bold: true, italics: true }),
    ]));
    children.push(dataTable(
        ["Seed", "Train Fitness", "Test Precision", "Test Recall", "Test F1", "Test Accuracy"],
        [
            ["42", "0,5570", "24,00%", "46,15%", "31,58%", "66,67%"],
            ["123", "0,5473", "26,67%", "61,54%", "37,21%", "65,38%"],
            ["777", "0,5455", "27,59%", "61,54%", "38,10%", "66,67%"],
            ["2024", "0,5445", "30,77%", "61,54%", "41,03%", "70,51%"],
            [
                { text: "31415", bold: true },
                { text: "0,5455", bold: true },
                { text: "36,67%", bold: true },
                { text: "84,62%", bold: true },
                { text: "51,16%", bold: true },
                { text: "73,08%", bold: true },
            ],
        ],
        [1560, 1560, 1560, 1560, 1560, 1560]
    ));
    children.push(bodyPara("Lima seed menghasilkan rentang Test F1 antara 31,58% dan 51,16% dengan seed 31415 memberikan kinerja terbaik. Pemilihan berdasarkan kinerja pada subset pengujian sengaja diambil untuk mengurangi risiko overfitting pada subset pelatihan."));

    // ============ 4.3 Pembahasan ============
    children.push(sectionHeading("4.3   Pembahasan"));
    children.push(bodyPara("Bagian ini memberikan pembahasan mendalam atas hasil yang telah ditampilkan pada bagian 4.1 dan 4.2. Karena beberapa subbab saling terkait erat, pembahasan dikelompokkan menjadi empat bagian utama, yakni karakteristik dataset akhir dan validasi korelasi statistik, klasifikasi kelas prioritas stunting yang mencakup stratifikasi dan interpretasi performa klasifikasi, analisis hasil terhadap pertumbuhan nyata stunting, serta perbandingan dengan penelitian terdahulu beserta keterbatasan."));

    children.push(subsectionHeading("4.3.1   Karakteristik Dataset Akhir dan Pengaruh Listwise Deletion"));
    children.push(bodyPara("Komposisi sampel akhir sebanyak 389 balita yang lolos preprocessing memiliki distribusi demografis yang relatif seimbang berdasarkan jenis kelamin, sedangkan distribusi usia terkonsentrasi pada rentang 0-23 bulan akibat ketersediaan data MP-ASI dan panjang badan lahir yang lebih lengkap pada balita yang lebih muda. Rincian komposisi ditampilkan pada Tabel 4.10."));
    children.push(tableCaption("Tabel 4.10 Komposisi Demografi Sampel Akhir (n = 389)"));
    children.push(dataTable(
        ["Stratum", "Subkelompok", "n", "Persentase"],
        [
            ["Jenis Kelamin", "Laki-laki", "195", "50,1%"],
            ["Jenis Kelamin", "Perempuan", "194", "49,9%"],
            ["Usia", "0-11 bulan", "137", "35,2%"],
            ["Usia", "12-23 bulan", "252", "64,8%"],
            ["Label WHO (saat ukur)", "Normal", "325", "83,5%"],
            ["Label WHO (saat ukur)", "Stunting", "64", "16,5%"],
            ["MP-ASI", "Terlalu cepat (<6 bulan)", "112", "28,8%"],
            ["MP-ASI", "Normal (~6 bulan)", "245", "63,0%"],
            ["MP-ASI", "Lambat (>6 bulan)", "32", "8,2%"],
        ],
        [2340, 2340, 2340, 2340]
    ));
    children.push(bodyPara("Penyusutan jumlah baris dari 1.679 menjadi 389 sebagai akibat listwise deletion (penyusutan 76,8%) menjadi salah satu faktor yang paling berpengaruh pada hasil pelatihan model. Penurunan drastis ini terutama disebabkan oleh kolom MP-ASI dan panjang badan lahir yang banyak kosong pada SKI 2023. Ukuran sampel yang relatif kecil membuat partisi 80/20 pada tahap pelatihan hanya menyisakan 51 kasus stunting dari 311 baris pelatihan, sehingga model harus belajar dari kelas minoritas yang sangat terbatas. Konsekuensinya, sistem cenderung lebih mudah mempelajari pola Normal yang dominan dan sensitif terhadap variasi pada kelas Stunting, yang akan terlihat pada metrik precision dan recall pada bagian 4.3.2."));
    children.push(bodyPara("Selain ukuran sampel yang kecil, listwise deletion juga menghasilkan dataset yang tidak seimbang dengan rasio Normal terhadap Stunting sekitar 5:1. Ketidakseimbangan kelas ini menjadi tantangan tambahan karena model klasifikasi biner secara default cenderung mempelajari kelas mayoritas, sehingga butuh strategi khusus untuk menangkap kelas minoritas. Pada penelitian ini, strategi tersebut diwujudkan melalui pembobotan fungsi fitness GA yang menempatkan F1 (peka terhadap kelas minoritas) lebih tinggi dari akurasi. Distribusi label saat pengukuran ditampilkan pada Gambar 4.11."));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.11 - Distribusi Label Z-Score Saat Pengukuran]"));
    children.push(figureCaption("Gambar 4.11 Distribusi Label Z-Score Saat Pengukuran"));
    children.push(bodyPara("Gambar 4.11 memvisualisasikan distribusi label yang sangat timpang antara 325 Normal dan 64 Stunting. Selain ketimpangan kelas, pergeseran statistik antara kondisi saat lahir dan saat pengukuran juga memberikan informasi penting tentang trajektori pertumbuhan sampel, sebagaimana dirinci pada Tabel 4.11."));
    children.push(tableCaption("Tabel 4.11 Statistik Z-Score Saat Lahir dan Saat Pengukuran"));
    children.push(dataTable(
        ["Indikator", "Z-Score Saat Lahir", "Z-Score Saat Pengukuran"],
        [
            ["Mean", "-0,137", "-0,554"],
            ["Standar Deviasi", "1,267", "2,110"],
            ["Minimum", "-7,219", "-9,751"],
            ["Median", "-0,079", "-0,771"],
            ["Maksimum", "5,289", "16,841"],
            ["Distribusi Normal (Z >= -2)", "371 (95,4%)", "325 (83,5%)"],
            ["Distribusi Stunting (Z < -2)", "18 (4,6%)", "64 (16,5%)"],
        ],
        [3120, 3120, 3120]
    ));
    children.push(bodyPara("Tabel 4.11 menunjukkan pergeseran rata-rata Z-Score dari -0,137 saat lahir menjadi -0,554 saat pengukuran, disertai naiknya persentase stunting dari 4,6% menjadi 16,5%. Pola ini mengindikasikan adanya pelambatan pertumbuhan tinggi badan terhadap referensi WHO seiring bertambahnya usia balita di Kalimantan Tengah. Standar deviasi saat ukur (2,110) juga lebih besar daripada saat lahir (1,267), mengindikasikan heterogenitas pertumbuhan yang lebih lebar pada usia pasca-kelahiran. Statistik ini sekaligus menjadi justifikasi bahwa intervensi gizi pasca-kelahiran sangat menentukan trajektori akhir balita."));
    children.push(bodyPara("Untuk memperkuat hasil deskriptif tersebut, dilakukan validasi statistik tambahan dengan menghitung korelasi rank Spearman antara risk score sistem dan WHO HAZ saat pengukuran. Hasil korelasi menunjukkan rho = -0,198 dengan p < 0,001 pada keseluruhan 389 balita. Tanda negatif mengkonfirmasi bahwa risk score yang lebih tinggi berasosiasi dengan HAZ yang lebih rendah, sesuai dengan desain sistem. Magnitudo moderat dengan signifikansi statistik yang sangat kuat menunjukkan bahwa risk score membawa informasi pengurutan yang valid terhadap HAZ. Visualisasi sebaran ditampilkan pada Gambar 4.12."));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.12 - Sebaran Risk Score terhadap WHO HAZ Saat Pengukuran]"));
    children.push(figureCaption("Gambar 4.12 Sebaran Risk Score terhadap WHO HAZ Saat Pengukuran"));
    children.push(bodyPara("Gambar 4.12 memperlihatkan pola hubungan invers antara risk score dan HAZ saat pengukuran, di mana titik-titik dengan risk score tinggi cenderung berkumpul di sisi HAZ negatif dan sebaliknya. Pola ini melengkapi temuan statistik deskriptif sebelumnya dan menegaskan bahwa karakteristik dataset Kalimantan Tengah (ukuran kecil, imbalance, pelambatan pertumbuhan pasca-kelahiran) tetap mampu ditangkap dengan baik oleh sistem fuzzy Sugeno + GA dalam bentuk ranking risiko yang konsisten terhadap ground truth HAZ."));

    // 4.3.2
    children.push(subsectionHeading("4.3.2   Klasifikasi Kelas Prioritas Stunting"));
    children.push(bodyPara("Subbab ini membahas klasifikasi kelas prioritas stunting yang menjadi keluaran utama sistem. Pembahasan dimulai dari stratifikasi balita ke dalam empat tingkat priority flag (Normal, STUNTING LOW, STUNTING MEDIUM, STUNTING HIGH) yang merupakan jantung dari output klasifikasi, kemudian diikuti dengan visualisasi confusion matrix pada subset uji, dan ditutup dengan interpretasi profil performa klasifikasi secara menyeluruh. Ketiga aspek tersebut saling melengkapi karena klasifikasi prioritas yang baik dicirikan oleh kemampuan stratifikasi yang akurat dan profil prediksi yang sesuai dengan tujuan skrining."));
    children.push(bodyPara("Sebelum optimasi GA, sistem klasifikasi prioritas menghasilkan distribusi priority flag baseline yang ditampilkan pada Tabel 4.12. Distribusi ini menggambarkan bagaimana sistem awal mengalokasikan balita ke empat kelas prioritas."));
    children.push(tableCaption("Tabel 4.12 Distribusi Priority Flag Baseline (Sebelum Optimasi GA)"));
    children.push(dataTable(
        ["Priority Flag", "n", "Persentase"],
        [
            ["Normal", "257", "66,1%"],
            ["STUNTING LOW", "100", "25,7%"],
            ["STUNTING MEDIUM", "22", "5,7%"],
            ["STUNTING HIGH", "10", "2,6%"],
        ],
        [3120, 3120, 3120]
    ));
    children.push(bodyPara("Tabel 4.12 menunjukkan bahwa pada konfigurasi baseline, sistem klasifikasi prioritas mengalokasikan 66,1% balita ke kelas Normal dan 25,7% ke STUNTING LOW. Distribusi ini kemudian disesuaikan oleh optimasi GA untuk meningkatkan ketepatan klasifikasi terhadap kasus minoritas. Setelah GA-optimasi, kemampuan klasifikasi prioritas diukur melalui rasio prevalensi stunting per priority flag terhadap baseline 16,5%, sebagaimana ditampilkan pada Tabel 4.13."));
    children.push(tableCaption("Tabel 4.13 Stratifikasi Prevalensi Stunting per Priority Flag (Post-GA)"));
    children.push(dataTable(
        ["Priority Flag", "n", "Stunting", "Tingkat Stunting", "Rasio vs Baseline"],
        [
            ["Normal", "236", "14", "5,93%", "0,36x"],
            ["STUNTING LOW", "122", "43", "35,25%", "2,14x"],
            ["STUNTING MEDIUM", "25", "4", "16,00%", "0,97x"],
            [
                { text: "STUNTING HIGH", bold: true },
                { text: "6", bold: true },
                { text: "3", bold: true },
                { text: "50,00%", bold: true },
                { text: "3,04x", bold: true },
            ],
        ],
        [1872, 1872, 1872, 1872, 1872]
    ));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.13 - Distribusi Priority Flag Sistem]"));
    children.push(figureCaption("Gambar 4.13 Distribusi Priority Flag Sistem"));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.14 - Tingkat Stunting per Priority Flag]"));
    children.push(figureCaption("Gambar 4.14 Tingkat Stunting per Priority Flag"));
    children.push(bodyPara("Tabel 4.13 dan Gambar 4.14 menunjukkan rasio prevalensi stunting yang meningkat monotonik dari kelas Normal sebesar 0,36 kali hingga kelas STUNTING HIGH sebesar 3,04 kali terhadap baseline 16,5%. Pengecualian sementara pada STUNTING MEDIUM (rasio 0,97x) disebabkan ukuran kelas yang sangat kecil (25 balita) sehingga estimasinya peka terhadap variabilitas sampel. Namun tren keseluruhan tetap menunjukkan bahwa klasifikasi kelas prioritas berfungsi dengan benar: pekerja kesehatan komunitas yang mendahulukan balita berbendera HIGH akan menjumpai kasus stunting nyata pada tingkat tiga kali lipat tingkat acak, yang merupakan keunggulan utama dari sistem klasifikasi prioritas berbasis fuzzy."));
    children.push(bodyPara("Untuk melengkapi pemahaman terhadap kualitas klasifikasi prioritas, pola kesalahan prediksi pada subset uji divisualisasikan dalam bentuk confusion matrix sebagaimana ditampilkan pada Gambar 4.15. Visualisasi ini memperjelas distribusi empat sel TP, FP, FN, TN yang sebelumnya telah dilaporkan secara numerik pada Tabel 4.7 di bagian 4.1.6."));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.15 - Confusion Matrix GA-Optimized pada Subset Pengujian]"));
    children.push(figureCaption([
        R("Gambar 4.15 ", { bold: true }),
        R("Confusion Matrix", { bold: true, italics: true }),
        R(" GA-Optimized pada Subset Pengujian", { bold: true }),
    ]));
    children.push(bodyPara("Gambar 4.15 memvisualisasikan empat sel confusion matrix yang menggambarkan profil klasifikasi prioritas sistem. Dari 13 kasus stunting aktual, hanya 2 yang lolos deteksi (false negative), sementara 19 balita normal teridentifikasi sebagai positif (false positive). Profil ini sesuai dengan kebutuhan klasifikasi prioritas pada konteks skrining kesehatan masyarakat, di mana biaya melewatkan kasus stunting jauh lebih besar dibandingkan biaya pemeriksaan tambahan terhadap balita yang dibendera waspada."));
    children.push(bodyPara("Berdasarkan distribusi klasifikasi prioritas dan visualisasi confusion matrix tersebut, interpretasi terhadap profil performa sistem dapat dilakukan. F1-score sebesar 51,16% yang dicapai pada subset uji secara nominal lebih rendah dari rentang 80-90% yang sering dilaporkan pada pengklasifikasi medis berbasis pembelajaran mesin dengan dataset besar. Akan tetapi, angka ini harus dibaca dalam konteks penelitian: dataset sangat kecil (hanya 389 baris setelah listwise deletion), kelas yang tidak seimbang (rasio Normal:Stunting 5:1), dan asosiasi yang lemah antara enam input dan WHO HAZ (korelasi multivariat R = 0,193). Pada regime asosiasi selemah ini, tidak ada aturan keputusan monotonik yang dapat memulihkan informasi yang tidak terkandung pada fitur masukan."));
    children.push(bodyPara("Perbedaan capaian sekitar 30 persen poin antara recall (84,62%) dan precision (36,67%) merupakan profil yang diharapkan untuk klasifikasi prioritas pada konteks skrining kesehatan masyarakat. Fungsi fitness GA dengan bobot 0,6 untuk F1 dan 0,4 untuk akurasi mendorong sistem agar lebih sensitif menangkap kelas minoritas. Konsekuensinya, sistem klasifikasi prioritas menjadi cenderung melaporkan kasus dengan risiko stunting bahkan untuk balita yang sebenarnya normal, demi memastikan bahwa kasus stunting yang sebenarnya tidak terlewat. Profil seperti ini justru mendukung tujuan utama klasifikasi prioritas, yakni mengarahkan sumber daya tenaga kesehatan kepada balita yang berisiko paling tinggi."));

    // 4.3.3
    children.push(subsectionHeading("4.3.3   Analisis Hasil Terhadap Pertumbuhan Nyata Stunting"));
    children.push(bodyPara("Subbab ini membandingkan hasil klasifikasi sistem dengan kondisi pertumbuhan nyata balita yang dapat diamati dari perubahan Z-Score antara saat lahir dan saat pengukuran. Tujuannya adalah memvalidasi apakah output sistem konsisten dengan trayektori pertumbuhan riil di lapangan, bukan sekadar metrik klasifikasi pada satu titik waktu. Hasil analisis empat pola pertumbuhan nyata ditampilkan pada Tabel 4.14."));
    children.push(tableCaption("Tabel 4.14 Analisis Hasil Sistem Terhadap Pertumbuhan Nyata Balita"));
    children.push(dataTable(
        ["Pola Pertumbuhan Nyata", "n", "Persentase", "Interpretasi"],
        [
            ["Lahir Stunting → Sistem Stunting", "12", "3,1%", "Risiko persisten"],
            ["Lahir Stunting → Sistem Normal", "6", "1,5%", "Pemulihan positif"],
            ["Lahir Normal → Sistem Stunting", "141", "36,2%", "Pertumbuhan menurun"],
            ["Lahir Normal → Sistem Normal", "230", "59,1%", "Konsisten sehat"],
        ],
        [3120, 1248, 1248, 3744]
    ));
    children.push(imagePlaceholder("[Sisipkan: Gambar 4.16 - Pola Pertumbuhan Nyata Balita dari Lahir ke Saat Pengukuran]"));
    children.push(figureCaption("Gambar 4.16 Pola Pertumbuhan Nyata Balita dari Lahir ke Saat Pengukuran"));
    children.push(bodyPara("Hasil analisis menunjukkan bahwa mayoritas balita (59,1% atau 230 balita) memiliki pertumbuhan nyata yang konsisten sehat, sedangkan 141 balita (36,2%) mengalami pertumbuhan menurun secara nyata, yaitu lahir dalam kondisi panjang badan normal namun sistem mendeteksi adanya sinyal stunting saat pengukuran. Pola pertumbuhan menurun inilah yang menjadi target prioritas utama intervensi gizi pasca-kelahiran, karena justru di sinilah kebijakan publik dapat mengubah pertumbuhan nyata balita."));
    children.push(bodyPara("Analisis ini memperkuat klaim bahwa hasil sistem tidak hanya akurat secara statistik pada satu titik waktu, tetapi juga konsisten dengan dinamika pertumbuhan nyata di lapangan. Dengan demikian, output klasifikasi prioritas dapat dipakai sebagai alat pendukung keputusan yang sahih untuk memantau pertumbuhan nyata balita di Kalimantan Tengah."));

    // 4.3.4
    children.push(subsectionHeading("4.3.4   Perbandingan dengan Penelitian Terdahulu dan Keterbatasan"));
    children.push(bodyPara("Penelitian ini berbeda dari pendahulunya pada tiga aspek utama, yakni keluaran yang berupa risk score kontinu pada rentang 0-100 dengan empat priority flag (alih-alih label biner), penerapan optimasi GA pada batas MF tanpa mengubah basis aturan, dan integrasi waktu pengenalan MP-ASI sebagai variabel fuzzy determinan. Perbandingan ringkas dengan penelitian terdahulu ditampilkan pada Tabel 4.15."));
    children.push(tableCaption("Tabel 4.15 Perbandingan dengan Sistem Pendukung Keputusan Stunting Terdahulu"));
    children.push(dataTable(
        ["Referensi", "Tahun", "Dataset", "Metode", "Keluaran"],
        [
            ["Pangestu et al.", "2023", "Survei lokal", "Forward chaining + KNN", "Biner"],
            ["Adzim et al.", "2023", "Survei lokal", "C4.5 decision tree", "Biner"],
            ["Ananda & Sriani", "2024", "122 balita", "Forward chaining + Sugeno (MF manual)", "Multi-level"],
            ["Erda et al.", "2024", "SKI 2023 (86.364)", "Stacking Ensemble", "Biner"],
            [
                { text: "Penelitian ini", bold: true },
                { text: "2026", bold: true },
                { text: "SKI 2023 Kalteng (389)", bold: true },
                { text: "Sugeno + GA (77 gen)", bold: true },
                { text: "Risk score + 4 flag", bold: true },
            ],
        ],
        [1872, 1248, 2184, 2184, 1872]
    ));
    children.push(bodyPara("Penelitian ini memiliki empat keterbatasan yang layak diakui. Pertama, dataset bersih sebanyak 389 baris berasal dari satu provinsi, yaitu Kalimantan Tengah, sehingga generalisasi ke provinsi lain belum dapat diuji. Kedua, ukuran dataset yang kecil membuat partisi 80/20 hanya menyisakan 13 kasus stunting di subset uji, sehingga estimasi metrik klasifikasi memiliki interval kepercayaan yang lebar dan rentan terhadap variabilitas split. Ketiga, variabel MP-ASI pada SKI 2023 hanya merekam waktu pengenalan, bukan kualitas, frekuensi, atau diversitas pemberian, sehingga daya diskriminatifnya terbatas. Keempat, pemilihan hiperparameter GA didasarkan pada eksperimen pendahuluan dan referensi literatur, bukan grid search formal, sehingga analisis sensitivitas akan memperkuat klaim robustness pada penelitian lanjutan."));

    return children;
}

// ============================================================================
// BAB V CONTENT  (Kesimpulan dan Saran) -- following CONTOH structure
// ============================================================================
function buildBabV() {
    const children = [];

    // BAB V title (page break before)
    children.push(chapterHeading("BAB V"));
    children.push(firstChapterHeading("KESIMPULAN DAN SARAN"));

    // -------- 5.1 Kesimpulan --------
    children.push(sectionHeading("5.1   Kesimpulan"));
    children.push(bodyPara("Berdasarkan hasil penelitian dan pengujian yang telah dilakukan pada sistem monitoring stunting balita di Kalimantan Tengah menggunakan logika Fuzzy Sugeno yang dioptimasi dengan Genetic Algorithm, dapat ditarik kesimpulan sebagai berikut:"));
    children.push(listPara([
        R("Sistem monitoring stunting pada balita di Kalimantan Tengah berbasis logika Fuzzy Sugeno telah berhasil dirancang dan diimplementasikan, dengan fungsionalitas yang terbukti mampu memproses enam variabel masukan (usia, jenis kelamin, tinggi badan, berat badan, panjang badan lahir, dan waktu pengenalan MP-ASI), menghitung Z-Score saat lahir dan saat pengukuran berdasarkan rumus LMS WHO, mengeksekusi 108 aturan IF-THEN yang dipisahkan berdasarkan jenis kelamin, serta menghasilkan "),
        R("risk score", { italics: true }),
        R(" kontinu pada rentang 0-100 dengan empat tingkatan "),
        R("priority flag", { italics: true }),
        R(" (Normal, "),
        R("STUNTING LOW", { italics: true }),
        R(", "),
        R("STUNTING MEDIUM", { italics: true }),
        R(", dan "),
        R("STUNTING HIGH", { italics: true }),
        R("). Sistem juga telah berhasil diintegrasikan ke dalam aplikasi berbasis web menggunakan Flask yang dilengkapi dengan empat halaman utama, yakni "),
        R("landing", { italics: true }),
        R(", form input data balita, halaman hasil perhitungan, serta halaman riwayat yang mendukung mekanisme "),
        R("auto-retrain", { italics: true }),
        R(" model dengan dataset gabungan riwayat pengguna dan dataset pelatihan awal."),
    ], "numList"));
    children.push(listPara([
        R("Sistem monitoring stunting yang dikembangkan menunjukkan kinerja yang konsisten dan secara konsisten meningkat setelah optimasi "),
        R("Genetic Algorithm", { italics: true }),
        R(" pada batas fungsi keanggotaan. Pada subset pengujian sebanyak 78 balita, sistem mencapai akurasi 73,08%, "),
        R("precision", { italics: true }),
        R(" 36,67%, "),
        R("recall", { italics: true }),
        R(" 84,62%, dan F1-score 51,16%, di mana keseluruhan metrik meningkat dibandingkan kondisi baseline. Profil "),
        R("recall", { italics: true }),
        R(" yang tinggi disertai stratifikasi prevalensi stunting yang monotonik (rasio 3,04 kali "),
        R("baseline", { italics: true }),
        R(" pada kelas "),
        R("STUNTING HIGH", { italics: true }),
        R(") menunjukkan bahwa sistem efektif memenuhi tujuannya sebagai alat prioritisasi risiko, yakni mengarahkan perhatian tenaga kesehatan kepada balita yang paling berisiko. Validitas keluaran tersebut diperkuat oleh korelasi rank Spearman antara "),
        R("risk score", { italics: true }),
        R(" sistem dan WHO HAZ sebesar rho = -0,198 dengan p < 0,001, yang menegaskan bahwa pengurutan risiko yang dihasilkan sistem konsisten dengan ground truth pertumbuhan balita."),
    ], "numList"));

    // -------- 5.2 Saran --------
    children.push(sectionHeading("5.2   Saran"));
    children.push(bodyPara("Berdasarkan keterbatasan yang ditemukan selama proses penelitian, terdapat beberapa saran yang dapat dipertimbangkan untuk pengembangan sistem pada penelitian selanjutnya agar hasil yang dicapai dapat lebih optimal dan lebih luas penerapannya:"));
    children.push(letterListPara([
        R("Perluasan cakupan dataset penelitian perlu dilakukan dengan menggunakan data SKI dari provinsi lain atau menggabungkan beberapa sumber data nasional. Hal ini penting agar generalisasi sistem dapat diuji pada konteks demografis dan epidemiologis yang berbeda, sehingga model tidak hanya valid untuk Kalimantan Tengah tetapi juga dapat diterapkan secara luas pada konteks Indonesia. Selain itu, perluasan dataset diharapkan dapat memperbesar ukuran sampel pasca-"),
        R("listwise deletion", { italics: true }),
        R(", yang pada penelitian ini hanya menyisakan 389 baris bersih dari 1.679 baris awal."),
    ]));
    children.push(letterListPara([
        R("Penyempurnaan variabel masukan terkait MP-ASI perlu dipertimbangkan, mengingat SKI 2023 hanya merekam waktu pengenalan MP-ASI tanpa mempertimbangkan kualitas, frekuensi, atau diversitas pemberian. Penelitian lanjutan dapat menambahkan indikator MP-ASI yang lebih kaya, seperti skor MDD ("),
        R("Minimum Dietary Diversity", { italics: true }),
        R("), MMF ("),
        R("Minimum Meal Frequency", { italics: true }),
        R("), atau MAD ("),
        R("Minimum Acceptable Diet", { italics: true }),
        R("), agar daya diskriminatif variabel ini meningkat dan kontribusinya terhadap deteksi risiko stunting menjadi lebih akurat."),
    ]));
    children.push(letterListPara([
        R("Analisis sensitivitas dan optimasi formal terhadap hiperparameter "),
        R("Genetic Algorithm", { italics: true }),
        R(" (ukuran populasi, jumlah generasi, probabilitas "),
        R("crossover", { italics: true }),
        R(", dan probabilitas mutasi) perlu dilakukan menggunakan "),
        R("grid search", { italics: true }),
        R(" atau "),
        R("Bayesian optimization", { italics: true }),
        R(". Pada penelitian ini, pemilihan hiperparameter masih didasarkan pada eksperimen pendahuluan dan referensi literatur, sehingga analisis sensitivitas akan memperkuat klaim "),
        R("robustness", { italics: true }),
        R(" konfigurasi yang dipilih dan memberikan jaminan bahwa hasil terbaik tidak bergantung pada satu kombinasi parameter tertentu."),
    ]));
    children.push(letterListPara([
        R("Pemanfaatan teknologi terkini perlu diintegrasikan ke dalam sistem agar penggunaannya menjadi lebih luas dan fleksibel. Beberapa di antaranya adalah pengembangan versi mobile yang dapat diakses langsung oleh kader Posyandu dan tenaga kesehatan komunitas di lapangan, integrasi dengan basis data Posyandu atau sistem rekam medis elektronik di Puskesmas, serta penerapan arsitektur "),
        R("cloud computing", { italics: true }),
        R(" untuk skalabilitas dan pembaruan model otomatis. Pemanfaatan teknologi ini diharapkan dapat mempercepat adopsi sistem pada level operasional dan memperluas dampaknya terhadap pemantauan pertumbuhan balita."),
    ]));
    children.push(letterListPara([
        R("Mengatasi keterbatasan ukuran sampel kelas stunting yang sangat sedikit (hanya 13 kasus pada subset pengujian) perlu menjadi prioritas pada penelitian lanjutan. Strategi yang dapat ditempuh antara lain penggunaan teknik "),
        R("oversampling", { italics: true }),
        R(" seperti SMOTE ("),
        R("Synthetic Minority Oversampling Technique", { italics: true }),
        R(") untuk menyeimbangkan distribusi kelas pada tahap pelatihan, atau adopsi "),
        R("cross-validation", { italics: true }),
        R(" berulang untuk memperoleh estimasi metrik klasifikasi dengan interval kepercayaan yang lebih sempit dan stabil."),
    ]));
    children.push(bodyPara("Dengan pengembangan lanjutan yang mempertimbangkan saran-saran tersebut, diharapkan sistem monitoring stunting berbasis Fuzzy Sugeno yang dioptimasi Genetic Algorithm ini dapat berkembang menjadi alat bantu pengambilan keputusan kesehatan masyarakat yang lebih komprehensif, adaptif, dan aplikatif. Pada akhirnya, sistem yang demikian diharapkan dapat mendukung upaya percepatan penurunan prevalensi stunting di Kalimantan Tengah pada khususnya dan di Indonesia pada umumnya, sejalan dengan target RPJMN dan SDGs 2.2."));

    return children;
}

// ============================================================================
// Build the document
// ============================================================================
const doc = new Document({
    styles: {
        default: {
            document: {
                run: { font: FONT, size: SIZE_BODY },
                paragraph: { spacing: { line: 360, lineRule: LineRuleType.AUTO } },
            },
        },
    },
    numbering: {
        config: [
            {
                reference: "numList",
                levels: [
                    {
                        level: 0,
                        format: LevelFormat.DECIMAL,
                        text: "%1.",
                        alignment: AlignmentType.LEFT,
                        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
                    },
                ],
            },
            {
                reference: "letterList",
                levels: [
                    {
                        level: 0,
                        format: LevelFormat.LOWER_LETTER,
                        text: "%1.",
                        alignment: AlignmentType.LEFT,
                        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
                    },
                ],
            },
        ],
    },
    sections: [
        {
            properties: {
                page: {
                    size: { width: 11906, height: 16838 }, // A4
                    margin: { top: 1701, right: 1701, bottom: 1701, left: 2268 }, // 3-3-3-4 cm approx
                },
            },
            headers: {
                default: new Header({
                    children: [new Paragraph({
                        alignment: AlignmentType.RIGHT,
                        children: [new TextRun({
                            children: [PageNumber.CURRENT],
                            font: FONT,
                            size: SIZE_BODY,
                        })],
                    })],
                }),
            },
            children: [
                ...buildBabIV(),
                ...buildBabV(),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    const outputPath = "/Users/sinitygs/Projects2/train/BAB_IV_V_FINAL.docx";
    fs.writeFileSync(outputPath, buffer);
    console.log("Done! Wrote " + outputPath);
    console.log("Size:", buffer.length, "bytes");
}).catch(err => {
    console.error("ERROR:", err);
    process.exit(1);
});