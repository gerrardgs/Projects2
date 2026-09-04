// Generator for BAB II (Landasan Teori) revisi - Stunting Monitoring Skripsi
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat,
  BorderStyle, WidthType, PageNumber, Header, LineRuleType, HeadingLevel,
  Table, TableRow, TableCell, ShadingType
} = require('docx');

const FONT = "Times New Roman";
const SIZE_BODY = 24;        // 12pt
const SIZE_HEADING = 28;     // 14pt for chapter title
const SIZE_TABLE = 22;       // 11pt for table contents
const SIZE_TABLE_SM = 20;    // 10pt for smaller table contents

// Body paragraph (justified, 1.5 line spacing, first-line indent)
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

function plainPara(text, opts = {}) {
  const runs = Array.isArray(text)
    ? text
    : [new TextRun({ text: text, font: FONT, size: SIZE_BODY, ...(opts.runOpts || {}) })];
  return new Paragraph({
    children: runs,
    alignment: opts.alignment || AlignmentType.LEFT,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 0, after: 0 },
    ...opts.paraOpts,
  });
}

// Chapter heading - BAB II - bold and centered
function firstChapterHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_HEADING })],
    alignment: AlignmentType.CENTER,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 240, after: 240 },
  });
}

// Section heading like "2.1 Literatur Terkait Teori"
function sectionHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_BODY })],
    alignment: AlignmentType.LEFT,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 240, after: 120 },
  });
}

// Subsection like "2.1.1 Penelitian Terdahulu"
function subsectionHeading(text) {
  if (Array.isArray(text)) {
    return new Paragraph({
      children: text,
      alignment: AlignmentType.LEFT,
      spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 200, after: 100 },
    });
  }
  return new Paragraph({
    children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_BODY })],
    alignment: AlignmentType.LEFT,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 200, after: 100 },
  });
}

// Sub-subsection like "Panjang Badan"
function subSubHeading(text) {
  if (Array.isArray(text)) {
    return new Paragraph({
      children: text,
      alignment: AlignmentType.LEFT,
      spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 180, after: 80 },
    });
  }
  return new Paragraph({
    children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_BODY })],
    alignment: AlignmentType.LEFT,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 180, after: 80 },
  });
}

// Caption (centered)
function caption(text) {
  if (Array.isArray(text)) {
    return new Paragraph({
      children: text,
      alignment: AlignmentType.CENTER,
      spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 100, after: 100 },
    });
  }
  return new Paragraph({
    children: [new TextRun({ text: text, bold: true, font: FONT, size: SIZE_BODY })],
    alignment: AlignmentType.CENTER,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 100, after: 100 },
  });
}

// Image placeholder paragraph (no border to avoid pBdr schema ordering issue)
function imagePlaceholder(text) {
  return new Paragraph({
    children: [new TextRun({ text: text, italics: true, font: FONT, size: SIZE_BODY, color: "808080" })],
    alignment: AlignmentType.CENTER,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 100, after: 100 },
  });
}

// Equation paragraph (centered, italic-like)
function equation(text, label) {
  return new Paragraph({
    children: [
      new TextRun({ text: text, font: FONT, size: SIZE_BODY }),
      new TextRun({ text: "          ", font: FONT, size: SIZE_BODY }),
      new TextRun({ text: label, bold: true, font: FONT, size: SIZE_BODY }),
    ],
    alignment: AlignmentType.CENTER,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 100, after: 100 },
  });
}

// Bullet/numbered list paragraph helper
function bulletPara(text) {
  return new Paragraph({
    children: Array.isArray(text) ? text : [new TextRun({ text: text, font: FONT, size: SIZE_BODY })],
    numbering: { reference: "bulletList", level: 0 },
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 360, lineRule: LineRuleType.AUTO, before: 0, after: 40 },
  });
}

// Inline run helper
function R(text, opts = {}) {
  return new TextRun({
    text: text,
    font: FONT,
    size: opts.size || SIZE_BODY,
    italics: opts.italics || false,
    bold: opts.bold || false,
  });
}

// Table helpers
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

function tableCell(content, opts = {}) {
  const isHeader = opts.header === true;
  const align = opts.align || AlignmentType.LEFT;
  let paras;

  if (Array.isArray(content) && content.length > 0 && content[0] instanceof Paragraph) {
    paras = content;
  } else {
    const runs = Array.isArray(content)
      ? content
      : [new TextRun({
          text: String(content),
          bold: isHeader || opts.bold === true,
          italics: opts.italics || false,
          font: FONT,
          size: opts.size || SIZE_TABLE_SM,
        })];
    paras = [new Paragraph({
      children: runs,
      alignment: align,
      spacing: { line: 260, lineRule: LineRuleType.AUTO },
    })];
  }

  return new TableCell({
    borders: CELL_BORDERS,
    width: { size: opts.width || 2340, type: WidthType.DXA },
    shading: isHeader ? { fill: "E0E0E0", type: ShadingType.CLEAR } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: paras,
  });
}

function dataTable(headers, rows, columnWidths) {
  const totalWidth = columnWidths.reduce((a, b) => a + b, 0);
  const headerCells = headers.map((h, i) => tableCell(h, {
    header: true, align: AlignmentType.CENTER, width: columnWidths[i]
  }));
  const bodyRows = rows.map(r => new TableRow({
    children: r.map((c, i) => {
      if (typeof c === 'object' && c !== null && !Array.isArray(c) && !(c instanceof Paragraph)) {
        return tableCell(c.text || c.runs, {
          width: columnWidths[i],
          align: c.align || (i === 0 ? AlignmentType.LEFT : AlignmentType.LEFT),
          bold: c.bold || false,
        });
      }
      return tableCell(c, {
        width: columnWidths[i],
        align: i === 0 ? AlignmentType.LEFT : AlignmentType.LEFT,
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

// ============================================================================
// BAB II CONTENT
// ============================================================================
function buildBabII() {
  const children = [];

  // BAB II Title
  children.push(firstChapterHeading("BAB II"));
  children.push(firstChapterHeading("LANDASAN TEORI"));

  // 2.1 Literatur Terkait Teori
  children.push(sectionHeading("2.1   Literatur Terkait Teori"));

  // ====================================================================
  // 2.1.1 Penelitian Terdahulu  -- MODIFIED with additional entries
  // ====================================================================
  children.push(subsectionHeading("2.1.1   Penelitian Terdahulu"));
  children.push(bodyPara([
    R("Dalam rangka mendukung pengembangan penelitian yang diusulkan, peneliti mengkaji berbagai literatur yang berkaitan dengan klasifikasi, deteksi, maupun "),
    R("monitoring", { italics: true }),
    R(" risiko "),
    R("stunting", { italics: true }),
    R(" pada balita. Kajian ini mencakup pendekatan-pendekatan yang telah diterapkan oleh penelitian terdahulu, baik dari segi metode "),
    R("fuzzy", { italics: true }),
    R(", pemanfaatan dataset SKI 2023, maupun integrasi "),
    R("Genetic Algorithm", { italics: true }),
    R(" terhadap sistem inferensi "),
    R("fuzzy", { italics: true }),
    R(". Kajian dapat dilihat pada Tabel II.1 di bawah ini."),
  ]));
  children.push(caption("Tabel II.1 Hasil Kajian Penelitian Terdahulu"));
  children.push(dataTable(
    ["No", "Judul Penelitian", "Penulis & Tahun", "Penjelasan Singkat & Persamaan", "Hasil Utama", "Perbedaan / Kelemahan"],
    [
      [
        "1",
        { runs: [R("Sistem Pakar Diagnosis Stunting pada Balita Menggunakan Metode Forward Chaining dan Logika Fuzzy Sugeno", { italics: true })] },
        "Ananda & Sriani, 2024",
        "Penelitian Ananda dan Sriani (2024) membangun sistem pakar berbasis Forward Chaining dan Fuzzy Sugeno untuk mendiagnosis stunting balita dari Z-score BB/U, TB/U, dan BB/TB dengan keluaran tingkat risiko melalui defuzzifikasi yang memiliki kesamaan dengan penelitian ini karena sama-sama memakai Fuzzy Sugeno berbasis standar WHO.",
        "Sistem pakar web menggabungkan forward chaining dan Sugeno untuk menilai keparahan stunting dari 122 data, dengan distribusi risiko tinggi 57%, sangat tinggi 26%, dan rendah 17%.",
        "Populasi data lokal (Puskesmas), fungsi keanggotaan dirancang manual tanpa optimasi terhadap distribusi empiris data target, dan variabel pendukung non-antropometri terbatas.",
      ],
      [
        "2",
        { runs: [R("Classify the Nutritional Status of Iraqi Children under Five Years Using Fuzzy Classification", { italics: true })] },
        "Hasan & Mohammad, 2024",
        "Penelitian Hasan dan Mohammad (2024) mengklasifikasikan status gizi anak usia <5 tahun di Irak menggunakan Fuzzy Sugeno berbasis berat dan tinggi badan (BMI) pada 16.487 anak yang memiliki kesamaan dengan penelitian ini karena sama-sama menggunakan Fuzzy Sugeno berbasis indikator WHO untuk klasifikasi gizi.",
        "Membandingkan Mamdani vs Sugeno pada 16.487 anak (12 kelompok umur) berbasis variabel bobot dan tinggi. Sugeno unggul dengan akurasi sekitar 60% pada berbagai kelompok usia dan jenis kelamin.",
        "Populasi di Irak, variabel input minim (berat, tinggi) dan keluaran berbasis BMI, akurasi belum presisi, serta belum menghubungkan hasil ke prioritas kebijakan lokal.",
      ],
      [
        "3",
        { runs: [R("Advanced Machine Learning Techniques for Stunting Classification Using a Stacking Ensemble Approach", { italics: true })] },
        "Erda et al., 2024",
        "Penelitian Erda et al. (2024) menerapkan stacking ensemble pada dataset SKI 2023 berskala besar (86.364 record) untuk klasifikasi stunting yang memiliki kesamaan dengan penelitian ini karena sama-sama menggunakan dataset SKI 2023 sebagai sumber data utama.",
        "Stacking ensemble pada SKI 2023 menghasilkan klasifikasi stunting biner dengan performa kompetitif dibanding model machine learning konvensional pada skala nasional.",
        "Keluaran berupa label biner tanpa stratifikasi risiko, model bersifat black-box yang sulit ditelusuri ke variabel asal, dan belum memanfaatkan pendekatan fuzzy yang interpretable.",
      ],
      [
        "4",
        { runs: [R("Monitoring Toddler Development Based on Indonesia Health Survey (SKI) 2023", { italics: true })] },
        "Russiska & Cahyati, 2025",
        "Penelitian Russiska dan Cahyati (2025) memanfaatkan SKI 2023 untuk mengidentifikasi indikator antropometri dan perkembangan balita yang berkaitan dengan stunting yang memiliki kesamaan dengan penelitian ini karena sama-sama menggunakan SKI 2023 untuk kebutuhan monitoring pertumbuhan balita.",
        "Mengidentifikasi indikator perkembangan dan antropometri yang berkaitan dengan stunting pada SKI 2023 melalui analisis statistik deskriptif dan inferensial.",
        "Bersifat analisis statistik deskriptif tanpa sistem keputusan, belum menghasilkan risk score kontinu atau priority flag yang dapat dipakai untuk prioritisasi pada tingkat komunitas.",
      ],
      [
        "5",
        { runs: [R("A hybrid Genetic–Grey Wolf Optimization algorithm for optimizing Takagi–Sugeno–Kang fuzzy systems", { italics: true })] },
        "Elghamrawy & Hassanien, 2022",
        "Penelitian Elghamrawy dan Hassanien (2022) mengintegrasikan Genetic Algorithm dengan optimasi Grey Wolf untuk mengoptimasi sistem Takagi-Sugeno-Kang yang memiliki kesamaan dengan penelitian ini karena sama-sama menggunakan GA untuk menyempurnakan sistem fuzzy Sugeno.",
        "Hybrid GA-GWO terbukti meningkatkan performa Takagi-Sugeno-Kang pada beberapa benchmark dataset dengan tetap mempertahankan struktur aturan fuzzy.",
        "Belum diterapkan pada domain stunting maupun data antropometri balita Indonesia, sehingga relevansi epidemiologisnya untuk monitoring stunting belum teruji.",
      ],
      [
        "6",
        { runs: [R("An adaptive neuro-fuzzy inference system optimized by genetic algorithm for brain tumour detection in magnetic resonance images", { italics: true })] },
        "Ghahramani & Shiri, 2024",
        "Penelitian Ghahramani dan Shiri (2024) mengoptimasi parameter fungsi keanggotaan ANFIS Sugeno menggunakan GA untuk klasifikasi citra medis yang memiliki kesamaan dengan penelitian ini karena sama-sama menggunakan GA untuk menyesuaikan fungsi keanggotaan tanpa mengubah struktur sistem.",
        "Optimasi GA terhadap parameter fungsi keanggotaan ANFIS Sugeno meningkatkan akurasi deteksi tumor otak dibanding ANFIS tanpa optimasi.",
        "Domain aplikasi pada citra medis MRI, bukan data antropometri balita, sehingga penyesuaian variabel masukan dan basis aturan tetap diperlukan untuk konteks stunting.",
      ],
    ],
    [400, 1700, 1100, 2400, 2000, 1800]
  ));
  children.push(bodyPara([
    R("Penelitian pertama oleh Ananda dan Sriani (2024) menjadi rujukan paling dekat secara metodologi karena sama-sama memakai Fuzzy Sugeno untuk menilai tingkat risiko "),
    R("stunting", { italics: true }),
    R(". Namun, fungsi keanggotaan pada penelitian tersebut masih dirancang manual tanpa optimasi berbasis distribusi empiris data, sedangkan penelitian ini menggunakan "),
    R("Genetic Algorithm", { italics: true }),
    R(" untuk menyesuaikan batas fungsi keanggotaan agar lebih sesuai dengan karakteristik balita Kalimantan Tengah. Penelitian kedua oleh Hasan dan Mohammad (2024) menegaskan bahwa metode Sugeno secara konsisten lebih akurat dibandingkan Mamdani dalam mengakomodasi ketidakpastian data antropometri, sehingga memperkuat justifikasi pemilihan metode Sugeno pada penelitian ini."),
  ]));
  children.push(bodyPara([
    R("Penelitian ketiga oleh Erda et al. (2024) dan keempat oleh Russiska dan Cahyati (2025) memperlihatkan bahwa SKI 2023 telah dimanfaatkan untuk analisis "),
    R("stunting", { italics: true }),
    R(", tetapi belum dijadikan basis sistem "),
    R("fuzzy monitoring", { italics: true }),
    R(". Erda et al. (2024) menggunakan pendekatan "),
    R("stacking ensemble", { italics: true }),
    R(" berbasis "),
    R("machine learning", { italics: true }),
    R(" yang bersifat "),
    R("black-box", { italics: true }),
    R(" dengan keluaran biner, sedangkan Russiska dan Cahyati (2025) berhenti pada tataran analisis statistik tanpa sistem keputusan. Kedua penelitian tersebut menunjukkan adanya celah untuk memanfaatkan SKI 2023 dalam kerangka "),
    R("fuzzy", { italics: true }),
    R(" yang lebih interpretable dengan keluaran "),
    R("risk score", { italics: true }),
    R(" kontinu dan "),
    R("priority flag", { italics: true }),
    R(" bertingkat."),
  ]));
  children.push(bodyPara([
    R("Penelitian kelima oleh Elghamrawy dan Hassanien (2022) dan keenam oleh Ghahramani dan Shiri (2024) menjadi rujukan utama untuk justifikasi penggunaan "),
    R("Genetic Algorithm", { italics: true }),
    R(" sebagai optimizer fungsi keanggotaan sistem fuzzy Sugeno. Keduanya membuktikan bahwa GA mampu menyesuaikan parameter fungsi keanggotaan dengan baik tanpa mengubah struktur aturan, sehingga interpretabilitas sistem tetap terjaga. Namun, kedua penelitian tersebut belum diterapkan pada domain "),
    R("stunting", { italics: true }),
    R(", sehingga penelitian ini berkontribusi dengan mengadaptasi pendekatan GA terhadap sistem Fuzzy Sugeno khusus untuk monitoring risiko "),
    R("stunting", { italics: true }),
    R(" berbasis dataset SKI 2023 Kalimantan Tengah. Berdasarkan keseluruhan kajian tersebut, terdapat empat celah utama yang menjadi fokus penelitian ini, yakni belum adanya integrasi optimasi "),
    R("Genetic Algorithm", { italics: true }),
    R(" dengan Sugeno pada domain "),
    R("stunting", { italics: true }),
    R(", belum dimanfaatkannya SKI 2023 sebagai basis sistem "),
    R("fuzzy monitoring", { italics: true }),
    R(", belum dimasukkannya waktu pengenalan MP-ASI sebagai variabel "),
    R("fuzzy", { italics: true }),
    R(" determinan, serta keluaran sistem yang masih dominan biner sehingga tidak dapat menstratifikasi prioritas pada tingkat komunitas."),
  ]));

  // ====================================================================
  // 2.1.2 Stunting -- KEEP existing content
  // ====================================================================
  children.push(subsectionHeading([
    R("2.1.2   ", { bold: true }),
    R("Stunting", { bold: true, italics: true }),
  ]));
  children.push(bodyPara([
    R("Stunting", { italics: true }),
    R(" merupakan kondisi gangguan pertumbuhan dan perkembangan balita yang diakibatkan oleh kekurangan gizi kronis, infeksi berulang, serta kurangnya stimulasi psikososial yang memadai (Kurniawan, 2025). Penilaian terhadap status "),
    R("stunting", { italics: true }),
    R(" umumnya dilakukan berdasarkan pengukuran tinggi badan menurut usia ("),
    R("height-for-age", { italics: true }),
    R(") yang dibandingkan dengan "),
    R("World Health Organization", { italics: true }),
    R(" (WHO) "),
    R("Child Growth Standards", { italics: true }),
    R(". Penentuan batas antar kategori tersebut secara ilmiah didasarkan pada nilai "),
    R("Z-score", { italics: true }),
    R(" tinggi-badan-menurut-usia ("),
    R("Height-for-Age Z-score", { italics: true }),
    R("/HAZ), di mana balita dengan HAZ kurang dari -2 SD dikategorikan "),
    R("stunting", { italics: true }),
    R(" (Perumal et al., 2023; UNICEF et al., 2023)."),
  ]));

  // ====================================================================
  // 2.1.3 Z-Score -- KEEP
  // ====================================================================
  children.push(subsectionHeading("2.1.3   Z-Score"));
  children.push(bodyPara([
    R("Z-score", { italics: true }),
    R(" ("),
    R("standard score", { italics: true }),
    R(") merupakan ukuran baku yang menunjukkan seberapa jauh suatu nilai menyimpang dari rerata populasi dalam satuan simpangan baku. Dalam penilaian "),
    R("stunting", { italics: true }),
    R(", "),
    R("Z-score", { italics: true }),
    R(" khususnya "),
    R("Height-for-Age Z-score", { italics: true }),
    R(" (HAZ) digunakan untuk menilai status gizi balita dengan membandingkan tinggi badan terhadap usia berdasarkan standar antropometri WHO (Watson et al., 2024). Persamaan "),
    R("Z-score", { italics: true }),
    R(" secara matematis dapat dilihat pada rumus II.1 di bawah ini."),
  ]));
  children.push(equation("Z-Score = (x - μ) / σ", "(II.1)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara([
    R("x", { bold: true }),
    R(" adalah nilai individu pengukuran dari tinggi/panjang badan."),
  ]));
  children.push(bulletPara([
    R("μ", { bold: true }),
    R(" adalah nilai "),
    R("mean", { italics: true }),
    R(" dari rata-rata tinggi badan populasi referensi untuk usia dan jenis kelamin yang sama."),
  ]));
  children.push(bulletPara([
    R("σ", { bold: true }),
    R(" adalah nilai simpangan baku dari tinggi badan populasi referensi untuk usia dan jenis kelamin yang sama."),
  ]));
  children.push(bodyPara([
    R("Adapun cara lain untuk mencari "),
    R("Z-score", { italics: true }),
    R(" versi WHO menggunakan parameter LMS dapat dilihat pada Persamaan II.2 berikut."),
  ]));
  children.push(equation("Z-Score = ((x/M)^L - 1) / (L × S)", "(II.2)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara("x = nilai pengukuran antropometri individu (dalam hal ini panjang/tinggi badan)."));
  children.push(bulletPara([
    R("M = ("),
    R("Mu", { italics: true }),
    R(") nilai "),
    R("median", { italics: true }),
    R(" dari populasi referensi untuk kelompok usia dan jenis kelamin tertentu."),
  ]));
  children.push(bulletPara([
    R("L = ("),
    R("Lambda", { italics: true }),
    R(") "),
    R("parameter power", { italics: true }),
    R(" untuk mentransformasi data menuju normalitas."),
  ]));
  children.push(bulletPara("S = koefisien variasi (atau standar deviasi yang disesuaikan) dari populasi referensi."));
  children.push(bodyPara([
    R("Berdasarkan klasifikasi WHO, kategori "),
    R("stunting", { italics: true }),
    R(" pada balita usia 0-60 bulan secara ringkas dapat dilihat pada Tabel II.2 berikut ini."),
  ]));
  children.push(caption([
    R("Tabel II.2 Klasifikasi Status ", { bold: true }),
    R("Stunting", { bold: true, italics: true }),
    R(" Berdasarkan WHO HAZ", { bold: true }),
  ]));
  children.push(dataTable(
    ["Indeks Antropometri", "Status Gizi", "Ambang Batas (Z-Score)"],
    [
      [
        "Panjang Badan atau Tinggi Badan menurut Usia (PB/U atau TB/U) balita usia 0-60 bulan (Soliman et al., 2021)",
        { runs: [R("Stunting", { italics: true })] },
        "x ≤ -2,00 SD",
      ],
      [
        "",
        "Normal",
        "x > -2,00 SD",
      ],
    ],
    [4680, 2340, 2340]
  ));

  // ====================================================================
  // 2.1.4 SKI 2023 -- KEEP
  // ====================================================================
  children.push(subsectionHeading("2.1.4   SKI 2023"));
  children.push(bodyPara([
    R("Survei Kesehatan Indonesia (SKI) 2023 merupakan survei nasional yang diselenggarakan oleh Kementerian Kesehatan untuk memetakan kondisi kesehatan masyarakat Indonesia melalui 86.364 observasi yang merepresentasikan faktor sosial, ekonomi, lingkungan, dan demografi yang mempengaruhi status gizi serta indikator kesehatan lainnya (Erda et al., 2024). Survei dapat dilihat pada Gambar 2.1 berikut ini."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.1 - Buku Panduan Survei Kesehatan Indonesia]"));
  children.push(caption([
    R("Gambar 2.1 Buku Panduan Survei Kesehatan Indonesia di mana (a) Buku Pedoman Fundamental SKI 2023 dan (b) Buku Pedoman Deskripsi SKI 2023 Dalam Angka (Kemenkes, 2024)", { bold: true }),
  ]));
  children.push(bodyPara([
    R("Survei SKI sendiri dilaksanakan dalam bentuk kuisioner. Kuisioner SKI 2023 terdiri dari tiga instrumen utama, yaitu rumah tangga, rumah tangga balita, dan individu (dapat dilihat pada Gambar 2.2). Seluruh instrumen ini menjadi dasar penyusunan SKI 2023 dalam Angka dan Laporan Tematik SKI 2023, yang menyajikan ringkasan indikator kesehatan nasional dan analisis khusus, termasuk "),
    R("stunting", { italics: true }),
    R(" (Kemenkes, 2024)."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.2 - Bentuk Kuisioner SKI 2023]"));
  children.push(caption([
    R("Gambar 2.2 Bentuk Kuisioner dari Survei SKI 2023 di mana (a) Survei Rumah Tangga, (b) Survei Rumah Tangga Balita, dan (c) Survei Individu (Kemenkes, 2024)", { bold: true }),
  ]));

  // ====================================================================
  // 2.1.5 Variabel Penelitian -- MODIFIED (removed BBL, added TB ukur and BB ukur)
  // ====================================================================
  children.push(subsectionHeading("2.1.5   Variabel Penelitian"));
  children.push(bodyPara([
    R("Variabel penelitian memegang peran utama dalam memahami determinan "),
    R("stunting", { italics: true }),
    R(". Variabel "),
    R("dependen", { italics: true }),
    R(" pada penelitian ini adalah status "),
    R("stunting", { italics: true }),
    R(" yang ditentukan melalui ground truth WHO HAZ, sedangkan variabel "),
    R("independen", { italics: true }),
    R(" terdiri dari enam variabel yang memiliki relevansi klinis dan epidemiologis terhadap risiko "),
    R("stunting", { italics: true }),
    R(", yakni usia, jenis kelamin, panjang badan lahir (PBL), tinggi/panjang badan saat pengukuran, berat badan saat pengukuran, dan waktu pengenalan MP-ASI (Yani et al., 2023). Penjelasan masing-masing variabel diuraikan pada subbagian berikut."),
  ]));

  // -- 2.1.5.1 Panjang Badan (Saat Ukur) -- (renamed from "Panjang Badan" for clarity)
  children.push(subSubHeading("2.1.5.1   Tinggi/Panjang Badan Saat Pengukuran"));
  children.push(bodyPara([
    R("Stunting", { italics: true }),
    R(" adalah kondisi kronis ketika panjang atau tinggi balita berada di bawah standar usianya, ditandai nilai "),
    R("z-score", { italics: true }),
    R(" PB/U atau TB/U kurang dari -2 SD menurut WHO. Masalah ini terutama muncul akibat kekurangan gizi jangka panjang, infeksi berulang, dan faktor maternal. Tinggi atau panjang badan saat pengukuran merupakan indikator antropometri utama yang dipakai untuk menghitung HAZ sebagai "),
    R("ground truth", { italics: true }),
    R(" sistem. Pengukuran status dilakukan dengan LAZ untuk balita berusia ≤24 bulan (posisi terlentang) dan HAZ untuk balita berusia >24 bulan (posisi berdiri) sebagaimana diterapkan dalam berbagai studi (Rosdiana et al., 2024)."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.3 - Grafik Pertumbuhan Panjang/Tinggi Badan Balita Selama 1000 HPK]"));
  children.push(caption("Gambar 2.3 Grafik Ilustrasi Pertumbuhan Panjang/Tinggi Badan Balita Selama 1000 HPK Berdasarkan Usia dan Median Tinggi Badan Rujukan WHO (UNICEF, 2023)"));

  // -- 2.1.5.2 Berat Badan Saat Ukur (NEW)
  children.push(subSubHeading("2.1.5.2   Berat Badan Saat Pengukuran"));
  children.push(bodyPara([
    R("Berat badan saat pengukuran merupakan indikator antropometri yang melengkapi tinggi badan dalam menilai status gizi balita secara menyeluruh. Berat badan yang tidak proporsional terhadap tinggi badan dapat menjadi sinyal awal "),
    R("wasting", { italics: true }),
    R(" atau "),
    R("overweight", { italics: true }),
    R(" yang berinteraksi dengan risiko "),
    R("stunting", { italics: true }),
    R(" (de Onis & Branca, 2016). Pada sistem monitoring yang dikembangkan, berat badan saat pengukuran difuzzifikasi ke dalam empat himpunan linguistik (Kurang, Normal, Lebih, Obesitas) yang dipisahkan berdasarkan jenis kelamin untuk menangkap perbedaan pola pertumbuhan antara balita laki-laki dan perempuan."),
  ]));

  // -- 2.1.5.3 Usia Anak -- KEEP
  children.push(subSubHeading("2.1.5.3   Usia Anak"));
  children.push(bodyPara([
    R("Usia balita merupakan variabel penting dalam memahami dinamika "),
    R("stunting", { italics: true }),
    R(" karena merefleksikan akumulasi paparan risiko sejak lahir hingga 60 bulan. Secara keseluruhan, dalam rentang 0-60 bulan yang mencakup 1000 HPK, determinan biologis, perilaku, lingkungan, dan sosial-ekonomi terbukti konsisten memengaruhi prevalensi "),
    R("stunting", { italics: true }),
    R(" (Talukder et al., 2018). Temuan pada Gambar 2.4 juga menunjukkan bahwa prevalensi tertinggi terjadi pada kelompok usia 24-53 bulan, baik pada 2022 maupun 2023, meskipun angka tahun 2023 lebih rendah."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.4 - Prevalensi Stunting Berdasarkan Kelompok Umur 2022-2023]"));
  children.push(caption("Gambar 2.4 Prevalensi Stunting Berdasarkan Kelompok Umur Tahun 2022 dan 2023 (Kemenkes, 2024)"));

  // -- 2.1.5.4 Jenis Kelamin -- KEEP
  children.push(subSubHeading("2.1.5.4   Jenis Kelamin Anak"));
  children.push(bodyPara([
    R("Jenis kelamin merupakan variabel penting dalam pengukuran "),
    R("z-score", { italics: true }),
    R(" "),
    R("stunting", { italics: true }),
    R(" karena memengaruhi pola pertumbuhan, kerentanan biologis, serta respons terhadap faktor lingkungan. Bukti umum menunjukkan bahwa balita laki-laki lebih rentan mengalami "),
    R("stunting", { italics: true }),
    R(" dibandingkan perempuan, dan determinan sosioekonomi pun bekerja berbeda menurut jenis kelamin (de Onis & Branca, 2016). Pada sistem yang dikembangkan, jenis kelamin difungsikan sebagai variabel "),
    R("switching", { italics: true }),
    R(" yang memilih himpunan fungsi keanggotaan tinggi dan berat badan khusus laki-laki atau perempuan, sehingga interpretasi "),
    R("z-score", { italics: true }),
    R(" selalu mempertimbangkan jenis kelamin sebagaimana ditunjukkan pada Gambar 2.5 mengenai standar antropometri WHO."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.5 - Standar WHO Antropometri Per Jenis Kelamin]"));
  children.push(caption("Gambar 2.5 Standar WHO Antropometri pada Masing-masing Jenis Kelamin di mana (a) Grafik Standar Panjang/Tinggi Badan Perempuan dan (b) Grafik Standar Panjang/Tinggi Badan Laki-laki (UNICEF, 2023)"));

  // -- 2.1.5.5 Panjang Badan Lahir -- KEEP
  children.push(subSubHeading("2.1.5.5   Panjang Badan Lahir"));
  children.push(bodyPara([
    R("Panjang badan lahir (PBL) merupakan indikator penting dalam penilaian risiko "),
    R("stunting", { italics: true }),
    R(" karena secara konsisten terbukti sebagai prediktor kuat pertumbuhan linear. PBL mencerminkan kondisi gizi "),
    R("intrauterin", { italics: true }),
    R(" yang berpengaruh terhadap pertumbuhan linear anak, sehingga bayi dengan panjang lahir pendek memiliki kemungkinan lebih tinggi mengalami "),
    R("stunting", { italics: true }),
    R(" dibandingkan bayi dengan panjang lahir normal atau tinggi. Temuan longitudinal seperti "),
    R("Women First trial", { italics: true }),
    R(" juga menegaskan bahwa PBL merupakan prediktor terkuat status linear growth hingga 24 bulan, bahkan melampaui berat lahir, sebagaimana terlihat pada Gambar 2.6 yang menunjukkan bahwa "),
    R("birth LAZ", { italics: true }),
    R(" kurang dari -2 dan kurang dari -1 memberikan penurunan skor Z panjang menurut umur (LAZ) paling besar dibandingkan faktor lain (Krebs et al., 2022)."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.6 - Forest plot faktor yang memengaruhi LAZ pada usia 24 bulan]"));
  children.push(caption("Gambar 2.6 Forest Plot Faktor yang Memengaruhi Skor Z Panjang Menurut Umur (LAZ) pada Usia 24 Bulan (Krebs et al., 2022)"));

  // -- 2.1.5.6 MP-ASI -- KEEP
  children.push(subSubHeading("2.1.5.6   Makanan Pendamping ASI"));
  children.push(bodyPara([
    R("Waktu pemberian MP-ASI merupakan faktor penting dalam risiko "),
    R("stunting", { italics: true }),
    R(" karena menentukan kecukupan energi dan zat gizi selama transisi dari ASI eksklusif. Pemberian terlalu dini terbukti meningkatkan risiko malnutrisi, "),
    R("wasting", { italics: true }),
    R(", hingga "),
    R("stunting", { italics: true }),
    R(" akibat ketidaksiapan sistem pencernaan bayi dan kemungkinan tergantinya ASI dengan makanan yang kurang padat gizi (Muflih & Khasana, 2025). Berdasarkan penelitian Waqiyah et al. (2023), anak yang menerima MP-ASI tidak tepat waktu memiliki proporsi "),
    R("stunting", { italics: true }),
    R(" lebih tinggi (39,6%) dibandingkan anak dengan MP-ASI tepat waktu (16,0%), dengan hubungan yang signifikan secara statistik. Temuan ini menunjukkan bahwa pemberian MP-ASI sesuai usia rekomendasi bersifat protektif terhadap "),
    R("stunting", { italics: true }),
    R(", sedangkan penyimpangan waktu pemberian berpotensi menyebabkan defisit gizi yang berdampak pada pertumbuhan linear anak."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.7 - Hubungan MP-ASI dengan Kejadian Stunting]"));
  children.push(caption("Gambar 2.7 Hubungan Waktu Pemberian MP-ASI dan Pengetahuan Ibu tentang Responsive Feeding dengan Kejadian Stunting pada Anak Usia 6-24 Bulan (Waqiyah et al., 2023)"));

  // ====================================================================
  // 2.1.6 Logika Fuzzy -- KEEP (with small modification: variables list)
  // ====================================================================
  children.push(subsectionHeading("2.1.6   Logika Fuzzy"));
  children.push(bodyPara([
    R("Fuzzy logic", { italics: true }),
    R(" merupakan kerangka matematis untuk penalaran yang memodelkan derajat kebenaran kontinu antara 0-1 melalui teori himpunan samar, sehingga suatu elemen memiliki tingkat keanggotaan alih-alih status biner. Mekanisme dasarnya bekerja dengan aturan komposisional untuk menghasilkan keputusan dari informasi yang tidak pasti, menjadikannya fleksibel dan selaras dengan cara berpikir manusia. Pendekatan ini terbukti luas digunakan dalam berbagai aplikasi, mulai dari kendali proses hingga diagnosis medis, karena sifatnya yang interpretatif dan mudah diintegrasikan dalam sistem (Borgwardt & Peñaloza, 2017). Pada penelitian ini, variabel yang difuzzifikasi adalah usia, tinggi/panjang badan saat pengukuran, berat badan saat pengukuran, dan waktu pengenalan MP-ASI, sedangkan jenis kelamin dipakai sebagai variabel "),
    R("switching", { italics: true }),
    R(" yang memilih himpunan fungsi keanggotaan tinggi dan berat khusus laki-laki atau perempuan. Persamaan dasar himpunan "),
    R("fuzzy", { italics: true }),
    R(" secara matematis dapat dilihat pada rumus II.3 di bawah ini."),
  ]));
  children.push(equation("μ_A(x) ∈ [0, 1]", "(II.3)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara([
    R("μ", { bold: true }),
    R("A(x) adalah derajat keanggotaan ("),
    R("membership degree", { italics: true }),
    R(") elemen x terhadap himpunan samar A. Nilainya selalu dinormalisasi ke rentang 0-1."),
  ]));
  children.push(bulletPara([
    R("∈ [0, 1]", { bold: true }),
    R(" adalah rentang nilai antara 0-1, bisa tepat 0, tepat 1, atau nilai apa pun di antaranya (0,13; 0,5; 0,97; dst.)."),
  ]));
  children.push(bodyPara([
    R("Logika "),
    R("fuzzy", { italics: true }),
    R(" juga sering disebut sebagai "),
    R("Fuzzy Inference System", { italics: true }),
    R(" (FIS) atau "),
    R("Fuzzy Inference Engine", { italics: true }),
    R(" (Maulana et al., 2024). Gambar 2.8 menunjukkan alur kerja fundamental sistem logika "),
    R("fuzzy", { italics: true }),
    R("."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.8 - Alur Aturan Fuzzy]"));
  children.push(caption("Gambar 2.8 Alur Aturan Fuzzy (Maulana et al., 2024)"));

  // -- Fuzzifikasi
  children.push(subSubHeading("2.1.6.1   Fuzzifikasi"));
  children.push(bodyPara([
    R("Pada tahap awal dalam Gambar 2.8, nilai input tegas diubah menjadi derajat keanggotaan menggunakan fungsi keanggotaan, sehingga variabel seperti usia, tinggi badan, berat badan, dan MP-ASI dapat direpresentasikan dalam bentuk linguistik yang fleksibel (Gupta et al., 2022). Pada penelitian ini dipilih bentuk Trapesium, Segitiga, dan Linear. Representasi ketiga bentuk grafik dapat dilihat pada Gambar 2.9 berikut ini."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.9 - Representasi Grafik Fungsi Keanggotaan]"));
  children.push(caption("Gambar 2.9 Representasi Grafik di mana (a) Bentuk Trapesium (Maulana et al., 2024), (b) Bentuk Segitiga (Astonis, 2014), (c) Bentuk Linear (Wahyuni, 2018)"));

  // -- Trapesium
  children.push(subSubHeading("2.1.6.2   Trapesium"));
  children.push(bodyPara([
    R("Trapesium dipilih karena sederhana, mudah ditafsirkan, serta stabil dalam menggambarkan batas kategori variabel risiko "),
    R("stunting", { italics: true }),
    R(". Secara matematis, trapesium ditetapkan melalui empat parameter (a, b, c, d) yang membentuk pola naik-datar-turun (Putri & Saputro, 2021). Persamaan fungsi keanggotaan dengan parameter bentuk trapesium dapat dilihat pada rumus II.4 di bawah ini."),
  ]));
  children.push(equation("μ_trap(x; a, b, c, d) = max(min((x-a)/(b-a), 1, (d-x)/(d-c)), 0)", "(II.4)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara("Jika x ≤ a → kategori tidak berlaku (0)."));
  children.push(bulletPara("Jika a < x < b → kategori mulai berlaku (naik)."));
  children.push(bulletPara("Jika b ≤ x ≤ c → kategori berlaku sempurna (1)."));
  children.push(bulletPara("Jika c < x < d → kategori melemah (turun)."));
  children.push(bulletPara("Jika x ≥ d → kategori tidak berlaku (0)."));
  children.push(bodyPara("Dengan catatan parameter:", { noIndent: true }));
  children.push(bulletPara([
    R("a", { bold: true }),
    R(" adalah batas kiri dukungan ("),
    R("support", { italics: true }),
    R(") dengan nilai ≤ a dianggap tidak masuk kategori."),
  ]));
  children.push(bulletPara([
    R("b", { bold: true }),
    R(" adalah titik mulai keanggotaan penuh dengan x sudah masuk kategori dengan kekuatan 1."),
  ]));
  children.push(bulletPara([
    R("c", { bold: true }),
    R(" adalah titik akhir keanggotaan penuh dengan x masih 100% bagian dari kategori."),
  ]));
  children.push(bulletPara([
    R("d", { bold: true }),
    R(" adalah batas kanan dukungan dengan nilai ≥ d dianggap keluar kategori."),
  ]));

  // -- Segitiga
  children.push(subSubHeading("2.1.6.3   Segitiga"));
  children.push(bodyPara([
    R("Selain Trapesium, bentuk representatif yang digunakan adalah fungsi segitiga. Fungsi segitiga adalah salah satu bentuk fuzzifikasi sederhana yang terdiri dari tiga titik (a, b, c). Dalam konteks penelitian "),
    R("stunting", { italics: true }),
    R(", fungsi ini dapat digunakan untuk memodelkan variabel yang memiliki satu titik optimum (Abdy, 2018). Persamaan fungsi keanggotaan dengan parameter bentuk segitiga dapat dilihat pada rumus II.5 di bawah ini."),
  ]));
  children.push(equation("μ_tri(x; a, b, c) = max(min((x-a)/(b-a), (c-x)/(c-b)), 0)", "(II.5)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara("Jika x ≤ a atau x ≥ c maka kategori tidak berlaku (μ(x) = 0)."));
  children.push(bulletPara("Jika a ≤ x ≤ b maka kategori mulai berlaku (derajat keanggotaan naik dari 0 menuju 1)."));
  children.push(bulletPara("Jika b ≤ x ≤ c maka kategori mulai tidak berlaku (derajat keanggotaan turun dari 1 menuju 0)."));
  children.push(bodyPara("Dengan catatan parameter:", { noIndent: true }));
  children.push(bulletPara([
    R("a", { bold: true }),
    R(" = ambang batas kiri dukungan ("),
    R("left support", { italics: true }),
    R("), semua nilai ≤ a dianggap belum masuk kategori."),
  ]));
  children.push(bulletPara([
    R("b", { bold: true }),
    R(" = titik puncak, x = b memiliki keanggotaan penuh (μ = 1) dan paling mewakili kategori."),
  ]));
  children.push(bulletPara([
    R("c", { bold: true }),
    R(" = ambang batas kanan dukungan ("),
    R("right support", { italics: true }),
    R("), semua nilai ≥ c dianggap keluar dari kategori."),
  ]));

  // -- Kurva Linear Naik
  children.push(subSubHeading("2.1.6.4   Kurva Linear Naik"));
  children.push(bodyPara([
    R("Selain Trapesium dan Segitiga, bentuk representatif yang digunakan adalah kurva linear naik. Kurva linear naik adalah bentuk fungsi keanggotaan yang menggambarkan peningkatan derajat keanggotaan secara bertahap dan proporsional dari nilai 0 menuju 1 sepanjang suatu rentang nilai tertentu (Wahyuni, 2018). Persamaan fungsi keanggotaan dengan parameter bentuk kurva linear naik dapat dilihat pada rumus II.6 di bawah ini."),
  ]));
  children.push(equation("μ_curve(x; a, b) = max(min((x-a)/(b-a), 1), 0)", "(II.6)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara("Jika x ≤ a, maka kategori belum berlaku karena nilai masih berada di bawah ambang minimal, derajat keanggotaan tetap 0."));
  children.push(bulletPara("Jika a ≤ x ≤ b, maka kategori mulai aktif dan derajat keanggotaan meningkat secara bertahap dari 0 menuju 1 sesuai posisi x dalam rentang tersebut."));
  children.push(bulletPara("Jika x ≥ b, maka kategori berlaku sepenuhnya dan derajat keanggotaannya bernilai 1 (keanggotaan penuh)."));

  // -- Aturan Dasar
  children.push(subSubHeading("2.1.6.5   Aturan Dasar"));
  children.push(bodyPara([
    R("Metode "),
    R("fuzzy Takagi-Sugeno", { italics: true }),
    R(" (TS) memodelkan sistem nonlinier melalui aturan IF-THEN dengan konsekuen berbentuk fungsi tegas atau persamaan linear, sehingga mampu menangani ketidakpastian sekaligus menjaga struktur matematis yang efisien. Pendekatan ini relevan untuk kebutuhan kendali dan pengambilan keputusan modern (Yershov & Ponomarenko, 2018). Persamaan metode Sugeno secara matematis dapat dilihat pada rumus II.7 di bawah ini."),
  ]));
  children.push(equation("IF x1 adalah A1i AND x2 adalah A2i THEN y = ci", "(II.7)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara([
    R("x1, x2", { bold: true }),
    R(" = variabel masukan ("),
    R("input", { italics: true }),
    R(") sistem fuzzy."),
  ]));
  children.push(bulletPara([
    R("A1i, A2i", { bold: true }),
    R(" = himpunan "),
    R("fuzzy", { italics: true }),
    R(" (label linguistik) yang melekat pada masing-masing variabel masukan pada aturan ke-i."),
  ]));
  children.push(bulletPara([
    R("Operator AND", { bold: true }),
    R(" = operator logika "),
    R("fuzzy", { italics: true }),
    R(" (biasanya direalisasikan dengan t-norm, misalnya minimum atau perkalian)."),
  ]));
  children.push(bulletPara([
    R("y", { bold: true }),
    R(" = variabel keluaran ("),
    R("output", { italics: true }),
    R(") sistem "),
    R("fuzzy", { italics: true }),
    R(", yaitu skor risiko "),
    R("stunting", { italics: true }),
    R(" yang akan dipetakan menjadi kelas risiko dan "),
    R("priority flag", { italics: true }),
    R("."),
  ]));
  children.push(bulletPara([
    R("ci", { bold: true }),
    R(" = nilai konstan (konsekuen tegas) pada aturan Sugeno untuk aturan ke-i."),
  ]));

  // -- Defuzzifikasi
  children.push(subSubHeading("2.1.6.6   Defuzzifikasi"));
  children.push(bodyPara([
    R("Tahap terakhir pada Gambar 2.8 mengubah himpunan "),
    R("fuzzy", { italics: true }),
    R(" teragregasi menjadi nilai tegas melalui metode "),
    R("weighted average", { italics: true }),
    R(" pada Sugeno orde nol, menghasilkan output yang dapat digunakan langsung sebagai skor risiko (Georgescu et al., 2014). Rumus defuzzifikasi metode Sugeno dapat dilihat pada Persamaan II.8 di bawah ini."),
  ]));
  children.push(equation("y* = Σ(wi · ci) / Σ(wi)", "(II.8)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara([
    R("y*", { bold: true }),
    R(" = keluaran akhir sistem inferensi "),
    R("fuzzy", { italics: true }),
    R(" Sugeno, berupa satu nilai tegas ("),
    R("crisp", { italics: true }),
    R(") hasil penggabungan semua aturan."),
  ]));
  children.push(bulletPara([
    R("Indeks i", { bold: true }),
    R(" = nomor aturan ke-i dalam basis aturan ("),
    R("rule base", { italics: true }),
    R(")."),
  ]));
  children.push(bulletPara([
    R("ci", { bold: true }),
    R(" = nilai konsekuen konstan pada aturan ke-i."),
  ]));
  children.push(bulletPara([
    R("wi", { bold: true }),
    R(" = derajat aktivasi (bobot) aturan ke-i, yang diperoleh dari penggabungan derajat keanggotaan premis (hasil operator AND pada himpunan "),
    R("fuzzy", { italics: true }),
    R(")."),
  ]));

  // ====================================================================
  // 2.1.7 Genetic Algorithm -- NEW SUBSECTION
  // ====================================================================
  children.push(subsectionHeading([
    R("2.1.7   ", { bold: true }),
    R("Genetic Algorithm", { bold: true, italics: true }),
  ]));
  children.push(bodyPara([
    R("Genetic Algorithm", { italics: true }),
    R(" (GA) merupakan algoritma metaheuristik inspirasi biologis yang meniru proses seleksi alam untuk mencari solusi optimum pada ruang pencarian yang besar dan kompleks (Elghamrawy & Hassanien, 2022). GA bekerja dengan mempertahankan populasi calon solusi yang direpresentasikan sebagai "),
    R("individual", { italics: true }),
    R(" berupa kromosom (vektor gen), kemudian secara iteratif melakukan tiga operator utama, yaitu seleksi, "),
    R("crossover", { italics: true }),
    R(", dan mutasi, untuk menghasilkan generasi berikutnya yang lebih baik berdasarkan nilai fungsi "),
    R("fitness", { italics: true }),
    R(". Pada penelitian ini, GA digunakan untuk mengoptimasi batas titik fungsi keanggotaan "),
    R("fuzzy", { italics: true }),
    R(" Sugeno tanpa mengubah basis aturan, operator inferensi, maupun struktur utama sistem, sehingga interpretabilitas "),
    R("fuzzy", { italics: true }),
    R(" tetap terjaga sambil parameter fungsi keanggotaan menyesuaikan distribusi empiris data SKI 2023 (Ghahramani & Shiri, 2024)."),
  ]));
  children.push(bodyPara([
    R("Tiga operator utama GA yang diaplikasikan pada penelitian ini diuraikan sebagai berikut. "),
    R("Pertama", { bold: true }),
    R(", seleksi menggunakan metode "),
    R("tournament selection", { italics: true }),
    R(" dengan ukuran turnamen tiga, yakni tiga individu dipilih secara acak dan individu dengan "),
    R("fitness", { italics: true }),
    R(" tertinggi diteruskan ke kolam reproduksi. "),
    R("Kedua", { bold: true }),
    R(", "),
    R("crossover", { italics: true }),
    R(" menggunakan "),
    R("blend crossover", { italics: true }),
    R(" (cxBlend) dengan parameter α = 0,3 untuk mencampur gen dari dua induk dan menghasilkan keturunan yang berada di sekitar rentang gen induk. "),
    R("Ketiga", { bold: true }),
    R(", mutasi menggunakan "),
    R("bounded Gaussian mutation", { italics: true }),
    R(" dengan probabilitas per-gen 0,15 dan deviasi standar 0,15 kali rentang setiap gen, di mana setiap gen yang termutasi tetap dijaga dalam batas medis yang masuk akal."),
  ]));
  children.push(bodyPara([
    R("Fungsi "),
    R("fitness", { italics: true }),
    R(" yang digunakan adalah kombinasi tertimbang F1-score dan akurasi dengan bobot 0,6 untuk F1-score dan 0,4 untuk akurasi. Penggunaan bobot asimetris ini bertujuan untuk menjaga sensitivitas optimizer terhadap kelas minoritas ("),
    R("stunting", { italics: true }),
    R(") yang hanya sekitar 16,5% dari total sampel, sekaligus mencegah konfigurasi degeneratif yang hanya memaksimalkan "),
    R("recall", { italics: true }),
    R(". Persamaan fungsi "),
    R("fitness", { italics: true }),
    R(" dapat dilihat pada rumus II.9 berikut."),
  ]));
  children.push(equation("F(θ) = w1 · F1(θ) + w2 · Acc(θ),    w1 = 0,6;    w2 = 0,4", "(II.9)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara([
    R("F(θ)", { bold: true }),
    R(" = nilai "),
    R("fitness", { italics: true }),
    R(" individu θ pada subset pelatihan."),
  ]));
  children.push(bulletPara([
    R("F1(θ)", { bold: true }),
    R(" = F1-score yang dihitung dari biner "),
    R("priority flag", { italics: true }),
    R(" (Low/Medium/High vs Normal) terhadap label WHO HAZ."),
  ]));
  children.push(bulletPara([
    R("Acc(θ)", { bold: true }),
    R(" = akurasi prediksi sistem pada subset pelatihan."),
  ]));
  children.push(bulletPara([
    R("w1, w2", { bold: true }),
    R(" = bobot kontribusi F1-score dan akurasi, dengan w1 + w2 = 1 sehingga F(θ) ∈ [0, 1]."),
  ]));

  // ====================================================================
  // 2.1.8 Stratified Hold-out Validation -- KEEP
  // ====================================================================
  children.push(subsectionHeading([
    R("2.1.8   ", { bold: true }),
    R("Stratified Hold-out Validation", { bold: true, italics: true }),
  ]));
  children.push(bodyPara([
    R("Stratified hold-out validation", { italics: true }),
    R(" membagi data ke dalam set pelatihan dan pengujian dengan menjaga proporsi kelas label tetap seimbang di kedua subset, sehingga evaluasi tidak bias akibat ketidakseimbangan kelas yang umum terjadi pada data kesehatan (Abegaz et al., 2025). Pendekatan ini banyak digunakan pada studi klinis dan klasifikasi modern karena memberikan estimasi kinerja yang lebih representatif dibanding pembagian acak biasa, terutama ketika "),
    R("outcome", { italics: true }),
    R(" seperti "),
    R("stunting", { italics: true }),
    R(" (berdasarkan HAZ) tidak terdistribusi merata (Andrews et al., 2025; Bichri et al., 2024). Pada penelitian ini, "),
    R("stratified hold-out", { italics: true }),
    R(" dengan rasio 80/20 dipakai untuk membagi 389 sampel akhir menjadi 311 sampel pelatihan dan 78 sampel pengujian, dengan proporsi kelas "),
    R("stunting", { italics: true }),
    R(" tetap mendekati 16,5% pada kedua subset."),
  ]));

  // ====================================================================
  // 2.1.9 Confusion Matrix -- KEEP
  // ====================================================================
  children.push(subsectionHeading([
    R("2.1.9   ", { bold: true }),
    R("Confusion Matrix", { bold: true, italics: true }),
  ]));
  children.push(bodyPara([
    R("Konfusi matriks merupakan alat evaluasi dasar yang merangkum hubungan antara prediksi dan kondisi aktual untuk menghasilkan metrik penting seperti akurasi, presisi, "),
    R("recall", { italics: true }),
    R(", dan F1-score (Japkowicz & Shah, 2015). Gambar 2.10 menampilkan empat komponen utama, yaitu "),
    R("True Positive", { italics: true }),
    R(" (TP), "),
    R("False Positive", { italics: true }),
    R(" (FP), "),
    R("False Negative", { italics: true }),
    R(" (FN), dan "),
    R("True Negative", { italics: true }),
    R(" (TN) yang menjadi fondasi seluruh perhitungan tersebut."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.10 - Penjelasan Confusion Matrix]"));
  children.push(caption([
    R("Gambar 2.10 Penjelasan ", { bold: true }),
    R("Confusion Matrix", { bold: true, italics: true }),
    R(" (Liu & Yang, 2022)", { bold: true }),
  ]));
  children.push(bodyPara([
    R("Presisi"),
    R(" dan "),
    R("recall", { italics: true }),
    R(" digunakan untuk mengevaluasi model pada kondisi kelas tidak seimbang karena masing-masing menilai ketepatan prediksi positif dan kemampuan model menangkap seluruh kasus positif (Landgrebe et al., 2006). Persamaan "),
    R("precision", { italics: true }),
    R(" dan "),
    R("recall", { italics: true }),
    R(" secara matematis dapat dilihat pada rumus II.10 dan II.11 di bawah ini."),
  ]));
  children.push(equation("Precision = TP / (TP + FP)", "(II.10)"));
  children.push(equation("Recall = TP / (TP + FN)", "(II.11)"));
  children.push(bodyPara([
    R("F1-score merupakan metrik yang penting karena menyeimbangkan presisi dan "),
    R("recall", { italics: true }),
    R(" melalui rata-rata harmonik, sehingga lebih peka terhadap "),
    R("false positive", { italics: true }),
    R(" dan "),
    R("false negative", { italics: true }),
    R(", terutama pada data tidak seimbang (Hand et al., 2021). Persamaan F1-score secara matematis dapat dilihat pada rumus II.12 di bawah ini."),
  ]));
  children.push(equation("F1 = (2 × Precision × Recall) / (Precision + Recall)", "(II.12)"));
  children.push(bodyPara([
    R("Akurasi menunjukkan berapa banyak prediksi yang benar dari semua prediksi yang dibuat oleh model penelitian, dalam hal ini sistem "),
    R("fuzzy logic", { italics: true }),
    R(" Sugeno (Marsya Finda & Wahyu Utomo, 2024). Persamaan akurasi secara matematis dapat dilihat pada rumus II.13 di bawah ini."),
  ]));
  children.push(equation("Accuracy = (TP + TN) / (TP + TN + FP + FN)", "(II.13)"));
  children.push(bodyPara("Keterangan untuk seluruh persamaan II.10-II.13:", { noIndent: true }));
  children.push(bulletPara([
    R("TP: ", { bold: true }),
    R("True Positive", { italics: true }),
    R(", jumlah kasus "),
    R("stunting", { italics: true }),
    R(" yang diprediksi sebagai "),
    R("stunting", { italics: true }),
    R(" oleh sistem."),
  ]));
  children.push(bulletPara([
    R("TN: ", { bold: true }),
    R("True Negative", { italics: true }),
    R(", jumlah kasus Normal yang diprediksi sebagai Normal oleh sistem."),
  ]));
  children.push(bulletPara([
    R("FP: ", { bold: true }),
    R("False Positive", { italics: true }),
    R(", jumlah kasus Normal yang salah diprediksi sebagai "),
    R("stunting", { italics: true }),
    R(" oleh sistem."),
  ]));
  children.push(bulletPara([
    R("FN: ", { bold: true }),
    R("False Negative", { italics: true }),
    R(", jumlah kasus "),
    R("stunting", { italics: true }),
    R(" yang salah diprediksi sebagai Normal oleh sistem."),
  ]));

  // ====================================================================
  // 2.1.10 Spearman Rank Correlation -- NEW
  // ====================================================================
  children.push(subsectionHeading([
    R("2.1.10   ", { bold: true }),
    R("Spearman Rank Correlation", { bold: true, italics: true }),
  ]));
  children.push(bodyPara([
    R("Korelasi rank Spearman (ρ) merupakan metrik non-parametrik yang mengukur kekuatan dan arah hubungan monotonik antara dua variabel berdasarkan peringkat ("),
    R("rank", { italics: true }),
    R(") nilainya, bukan nilai mentahnya. Metrik ini cocok digunakan ketika dua variabel berasal dari skala ordinal atau ketika hubungan antar variabel tidak linear namun tetap monotonik. Nilai ρ berkisar antara -1 hingga +1, di mana ρ = +1 berarti hubungan monotonik positif sempurna, ρ = -1 berarti hubungan monotonik negatif sempurna, dan ρ = 0 berarti tidak ada hubungan monotonik. Persamaan korelasi rank Spearman secara matematis dapat dilihat pada rumus II.14 di bawah ini."),
  ]));
  children.push(equation("ρ = 1 - (6 · Σdi²) / (n · (n² - 1))", "(II.14)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara([
    R("ρ", { bold: true }),
    R(" = koefisien korelasi rank Spearman."),
  ]));
  children.push(bulletPara([
    R("di", { bold: true }),
    R(" = selisih peringkat pasangan data ke-i pada kedua variabel."),
  ]));
  children.push(bulletPara([
    R("n", { bold: true }),
    R(" = jumlah pasangan data."),
  ]));
  children.push(bodyPara([
    R("Pada penelitian ini, korelasi rank Spearman dipakai untuk mengukur kesesuaian monotonik antara "),
    R("risk score", { italics: true }),
    R(" kontinu yang dihasilkan sistem dengan WHO HAZ sebagai "),
    R("ground truth", { italics: true }),
    R(". Tanda negatif yang diharapkan menunjukkan bahwa "),
    R("risk score", { italics: true }),
    R(" yang lebih tinggi berasosiasi dengan HAZ yang lebih rendah, sesuai dengan desain sistem yang memetakan kondisi pertumbuhan terlambat ke skor risiko yang lebih tinggi."),
  ]));

  // ====================================================================
  // 2.1.11 Prevalence Stratification Ratio -- NEW
  // ====================================================================
  children.push(subsectionHeading([
    R("2.1.11   ", { bold: true }),
    R("Prevalence Stratification Ratio", { bold: true, italics: true }),
  ]));
  children.push(bodyPara([
    R("Prevalence stratification ratio", { italics: true }),
    R(" merupakan metrik operasional yang mengukur seberapa baik sistem klasifikasi prioritas dapat menstratifikasi kasus "),
    R("stunting", { italics: true }),
    R(" nyata pada masing-masing kelas "),
    R("priority flag", { italics: true }),
    R(" terhadap prevalensi "),
    R("baseline", { italics: true }),
    R(" pada keseluruhan sampel. Metrik ini didefinisikan sebagai rasio antara prevalensi "),
    R("stunting", { italics: true }),
    R(" empiris pada "),
    R("priority flag", { italics: true }),
    R(" tertentu dibagi prevalensi "),
    R("baseline", { italics: true }),
    R(" pada keseluruhan dataset. Persamaan "),
    R("prevalence stratification ratio", { italics: true }),
    R(" dapat dilihat pada rumus II.15 berikut."),
  ]));
  children.push(equation("Rasio_k = Prevalensi_k / Prevalensi_baseline", "(II.15)"));
  children.push(bodyPara("Di mana:", { noIndent: true }));
  children.push(bulletPara([
    R("Rasio_k", { bold: true }),
    R(" = "),
    R("prevalence stratification ratio", { italics: true }),
    R(" untuk "),
    R("priority flag", { italics: true }),
    R(" ke-k."),
  ]));
  children.push(bulletPara([
    R("Prevalensi_k", { bold: true }),
    R(" = proporsi kasus "),
    R("stunting", { italics: true }),
    R(" pada balita yang dikelompokkan ke "),
    R("priority flag", { italics: true }),
    R(" ke-k."),
  ]));
  children.push(bulletPara([
    R("Prevalensi_baseline", { bold: true }),
    R(" = proporsi kasus "),
    R("stunting", { italics: true }),
    R(" pada keseluruhan dataset (16,5% pada penelitian ini)."),
  ]));
  children.push(bodyPara([
    R("Rasio lebih besar dari 1 menunjukkan kelas "),
    R("priority flag", { italics: true }),
    R(" tersebut mengakumulasi kasus "),
    R("stunting", { italics: true }),
    R(" lebih tinggi daripada rata-rata acak, sedangkan rasio kurang dari 1 menunjukkan kelas tersebut bersifat protektif. Metrik ini sangat penting untuk konteks "),
    R("monitoring", { italics: true }),
    R(" karena memvalidasi kemampuan sistem dalam mengarahkan prioritas intervensi pada tingkat komunitas."),
  ]));

  // ====================================================================
  // 2.2 Alasan Pemilihan Metode -- HEAVILY MODIFIED
  // ====================================================================
  children.push(sectionHeading("2.2   Alasan Pemilihan Metode"));
  children.push(bodyPara([
    R("Pemilihan topik "),
    R("monitoring", { italics: true }),
    R(" risiko "),
    R("stunting", { italics: true }),
    R(" pada balita di Kalimantan Tengah didasari beban masalah yang masih tinggi dan penurunan yang berjalan lebih lambat dibanding target nasional, dengan prevalensi tercatat sekitar 23,5% pada 2023 dalam konteks kesenjangan antardaerah yang lebar (Kementerian Kesehatan Republik Indonesia, 2024). Kondisi ini menuntut pendekatan pemodelan risiko yang lebih presisi untuk mendukung penajaman sasaran intervensi, bukan sekadar pemantauan prevalensi agregat."),
  ]));
  children.push(imagePlaceholder("[Sisipkan: Gambar 2.11 - Prevalensi Stunting Per Provinsi 2022-2023]"));
  children.push(caption("Gambar 2.11 Prevalensi Stunting Berdasarkan Provinsi Tahun 2022 dan 2023 (Kemenkes, 2024)"));
  children.push(bodyPara([
    R("Meskipun beberapa provinsi di Indonesia menunjukkan prevalensi "),
    R("stunting", { italics: true }),
    R(" yang lebih ekstrem secara nasional, penelitian ini tidak bertujuan untuk membandingkan besaran prevalensi antarwilayah, melainkan untuk mengembangkan dan mengevaluasi sistem "),
    R("monitoring", { italics: true }),
    R(" risiko "),
    R("stunting", { italics: true }),
    R(" berbasis data. Dalam konteks pengembangan model, wilayah dengan prevalensi yang terlalu ekstrem cenderung menghasilkan distribusi kelas yang sangat timpang, sehingga kurang optimal untuk menguji sensitivitas model terhadap variasi tingkat risiko. Oleh karena itu, Kalimantan Tengah dipilih karena berada pada zona transisi epidemiologis "),
    R("stunting", { italics: true }),
    R(", dengan prevalensi di atas rata-rata nasional namun tetap menunjukkan heterogenitas kondisi balita yang luas. Karakteristik ini memungkinkan sistem "),
    R("fuzzy Takagi-Sugeno", { italics: true }),
    R(" mengaktifkan aturan secara parsial, menangkap transisi risiko secara lebih halus, serta menghasilkan skor risiko yang lebih informatif, stabil, dan valid secara metodologis. Dengan demikian, pemilihan Kalimantan Tengah didasarkan pada pertimbangan epidemiologis dan statistik sebagai wilayah dengan spektrum risiko yang paling representatif untuk pengembangan dan validasi sistem "),
    R("fuzzy", { italics: true }),
    R(" berbasis "),
    R("risk scoring", { italics: true }),
    R(", bukan semata-mata karena tingkat prevalensinya."),
  ]));
  children.push(bodyPara([
    R("Penelitian ini memandang "),
    R("stunting", { italics: true }),
    R(" sebagai fenomena yang bersifat gradual dan multifaktorial, sehingga "),
    R("monitoring", { italics: true }),
    R(" risiko berbasis HAZ diperluas menjadi "),
    R("risk scoring", { italics: true }),
    R(" kontinu pada rentang 0-100 yang dipetakan ke empat "),
    R("priority flag", { italics: true }),
    R(" (Normal, "),
    R("Stunting Low", { italics: true }),
    R(", "),
    R("Stunting Medium", { italics: true }),
    R(", dan "),
    R("Stunting High", { italics: true }),
    R("). Pemetaan ini menggunakan PBL sebagai indikator gizi "),
    R("intrauterin", { italics: true }),
    R(", tinggi dan berat badan saat pengukuran sebagai indikator antropometri terkini, serta waktu pengenalan MP-ASI sebagai indikator kecukupan asupan gizi pasca enam bulan pertama, untuk menghasilkan skor prioritas yang lebih informatif dibandingkan klasifikasi biner semata. Data SKI 2023 digunakan dengan sampel balita di Kalimantan Tengah yang memiliki data antropometri lengkap, sehingga seluruh kabupaten dan kota di Kalimantan Tengah masuk dalam populasi penelitian karena data SKI 2023 bersifat representatif provinsi."),
  ]));
  children.push(bodyPara([
    R("Model "),
    R("fuzzy Takagi-Sugeno", { italics: true }),
    R(" dipilih karena memiliki efisiensi komputasi tinggi, mampu merepresentasikan ketidakpastian, serta menghasilkan inferensi yang transparan dan mudah ditelusuri (Yershov & Ponomarenko, 2018). Sistem ini menggunakan kombinasi fungsi keanggotaan trapesium, segitiga, dan linear naik untuk memodelkan transisi gradual antarstatus gizi, yang secara biologis mencerminkan perubahan kontinu dalam pertumbuhan anak. Selain itu, batas fungsi keanggotaan dioptimasi menggunakan "),
    R("Genetic Algorithm", { italics: true }),
    R(" agar parameter menyesuaikan distribusi empiris data SKI 2023 Kalimantan Tengah tanpa mengubah basis aturan, operator inferensi, maupun struktur utama sistem, sehingga interpretabilitas "),
    R("fuzzy", { italics: true }),
    R(" tetap terjaga sambil performa klasifikasi meningkat (Elghamrawy & Hassanien, 2022; Ghahramani & Shiri, 2024)."),
  ]));
  children.push(bodyPara([
    R("Evaluasi kinerja sistem dilakukan menggunakan tiga kategori metrik yang saling melengkapi. "),
    R("Pertama", { bold: true }),
    R(", "),
    R("confusion matrix", { italics: true }),
    R(" beserta turunannya (akurasi, presisi, "),
    R("recall", { italics: true }),
    R(", F1-score) untuk mengukur ketepatan klasifikasi pada subset uji. "),
    R("Kedua", { bold: true }),
    R(", korelasi rank Spearman antara "),
    R("risk score", { italics: true }),
    R(" sistem dan WHO HAZ untuk memvalidasi kesesuaian monotonik secara statistik. "),
    R("Ketiga", { bold: true }),
    R(", "),
    R("prevalence stratification ratio", { italics: true }),
    R(" untuk menilai utilitas operasional sistem pada konteks "),
    R("monitoring", { italics: true }),
    R(" komunitas. Dibandingkan model kecerdasan buatan berbasis pembelajaran mesin yang bersifat "),
    R("black-box", { italics: true }),
    R(", "),
    R("fuzzy Takagi-Sugeno", { italics: true }),
    R(" memungkinkan setiap keputusan risiko ditelusuri kembali ke aturan, fungsi keanggotaan, dan variabel biologis yang digunakan. Dengan demikian, hasil "),
    R("monitoring", { italics: true }),
    R(" risiko "),
    R("stunting", { italics: true }),
    R(" tidak hanya bersifat numerik, tetapi juga dapat dijelaskan secara biologis dan kebijakan, sehingga lebih sesuai dengan karakteristik data SKI 2023 dan kebutuhan pengambilan keputusan dalam kesehatan masyarakat."),
  ]));

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
        reference: "bulletList",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
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
      children: buildBabII(),
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
    const outputPath = "/Users/sinitygs/Projects2/train/BAB_II_FINAL.docx";
    fs.writeFileSync(outputPath, buffer);
    console.log("Done! Wrote " + outputPath);
    console.log("Size:", buffer.length, "bytes");
}).catch(err => {
    console.error("ERROR:", err);
    process.exit(1);
});