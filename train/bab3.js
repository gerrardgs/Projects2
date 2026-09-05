const fs = require('fs');
const {
    Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
    AlignmentType, LevelFormat, HeadingLevel,
    BorderStyle, WidthType, ShadingType,
    TabStopType, TabStopPosition
} = require('docx');

const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };
const FONT = "Times New Roman";
const SIZE = 24;
const LINE_SPACING = { line: 360, lineRule: "auto" };

function parseInline(text, opts = {}) {
    const runs = [];
    const re = /\*([^*]+)\*/g;
    let last = 0, m;
    const baseRun = (s, italic = false) => new TextRun({
        text: s, font: FONT, size: SIZE, italics: italic, bold: opts.bold || false,
    });
    while ((m = re.exec(text)) !== null) {
        if (m.index > last) runs.push(baseRun(text.slice(last, m.index), false));
        runs.push(baseRun(m[1], true));
        last = m.index + m[0].length;
    }
    if (last < text.length) runs.push(baseRun(text.slice(last), false));
    if (runs.length === 0) runs.push(baseRun("", false));
    return runs;
}
function P(text, opts = {}) {
    return new Paragraph({
        children: parseInline(text, opts),
        alignment: opts.align || AlignmentType.JUSTIFIED,
        spacing: { ...LINE_SPACING, before: opts.before || 0, after: opts.after || 0 },
        indent: opts.indent || { firstLine: 720 },
    });
}
function H2(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 240, after: 120, ...LINE_SPACING },
        indent: { firstLine: 0 },
        children: [new TextRun({ text, font: FONT, size: 26, bold: true })],
    });
}
function H3(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 180, after: 120, ...LINE_SPACING },
        indent: { firstLine: 0 },
        children: [new TextRun({ text, font: FONT, size: 24, bold: true })],
    });
}
function H4(text) {
    return new Paragraph({
        heading: HeadingLevel.HEADING_4,
        spacing: { before: 160, after: 100, ...LINE_SPACING },
        indent: { firstLine: 0 },
        children: [new TextRun({ text, font: FONT, size: 24, bold: true })],
    });
}
function bulletP(text) {
    return new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        alignment: AlignmentType.JUSTIFIED,
        spacing: { ...LINE_SPACING, after: 60 },
        children: parseInline(text),
    });
}
function captionFig(text) {
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120, ...LINE_SPACING },
        indent: { firstLine: 0 },
        children: [new TextRun({ text, font: FONT, size: SIZE, bold: true })],
    });
}
function captionTab(text) {
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 60, ...LINE_SPACING },
        indent: { firstLine: 0 },
        children: [new TextRun({ text, font: FONT, size: SIZE, bold: true })],
    });
}
function placeholderFig(label) {
    return new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 60, after: 60, ...LINE_SPACING },
        indent: { firstLine: 0 },
        children: [new TextRun({ text: `[Sisipkan: ${label}]`, font: FONT, size: SIZE, italics: true, color: "808080" })],
    });
}
function makeCell(text, opts = {}) {
    const runs = typeof text === 'string' ? [new TextRun({
        text, font: FONT, size: SIZE, bold: opts.bold || false, italics: opts.italics || false
    })] : text;
    return new TableCell({
        borders,
        width: { size: opts.width, type: WidthType.DXA },
        shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
            alignment: opts.align || AlignmentType.LEFT,
            spacing: { line: 276 },
            indent: { firstLine: 0 },
            children: runs,
        })],
    });
}
function makeTable(headers, rows, colWidths) {
    const tableWidth = colWidths.reduce((a, b) => a + b, 0);
    const headerRow = new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => makeCell(h, {
            width: colWidths[i], bold: true, shading: "D9D9D9", align: AlignmentType.CENTER,
        })),
    });
    const dataRows = rows.map(row => new TableRow({
        children: row.map((c, i) => {
            const isObj = typeof c === 'object' && !Array.isArray(c);
            const text = isObj ? c.text : c;
            const align = (isObj && c.align) ? c.align : (i === 0 ? AlignmentType.LEFT : AlignmentType.CENTER);
            const bold = isObj ? c.bold : false;
            return makeCell(text, { width: colWidths[i], align, bold });
        }),
    }));
    return new Table({
        width: { size: tableWidth, type: WidthType.DXA },
        columnWidths: colWidths,
        rows: [headerRow, ...dataRows],
    });
}
function codeLine(text, indent = 0) {
    return new Paragraph({
        alignment: AlignmentType.LEFT,
        spacing: { line: 276, after: 0 },
        indent: { left: 360 + indent * 360, firstLine: 0 },
        children: [new TextRun({ text, font: "Consolas", size: 20 })],
    });
}

const children = [];

// ---- BAB III HEADING ----
children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 }, indent: { firstLine: 0 },
    children: [new TextRun({ text: "BAB III", font: FONT, size: 28, bold: true })],
}));
children.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 360 }, indent: { firstLine: 0 },
    children: [new TextRun({ text: "METODOLOGI PENELITIAN", font: FONT, size: 28, bold: true })],
}));

// ============================================================
// 3.1 ALUR PENELITIAN
// ============================================================
children.push(H2("3.1   Alur Penelitian"));

children.push(P("Alur penelitian disusun secara sistematis mulai dari studi literatur, pengumpulan dan praproses data, perancangan serta implementasi sistem, hingga pengujian, evaluasi, dan penyusunan laporan. Setiap tahap dirancang saling terintegrasi agar tujuan penelitian tercapai secara efektif. Rancangan penelitian yang diusulkan dapat dilihat pada Gambar 3.1."));

children.push(placeholderFig("Gambar 3.1 Alur Penelitian"));
children.push(captionFig("Gambar 3.1 Alur Penelitian"));

children.push(P("Adapun penjelasan mengenai alur penelitian adalah sebagai berikut."));

children.push(bulletP("Studi Literatur. Tahap awal penelitian adalah studi literatur untuk mengumpulkan dan menelaah referensi terkait *monitoring* risiko *stunting*, standar WHO HAZ, logika *fuzzy* Sugeno, *Genetic Algorithm*, serta penelitian terdahulu yang relevan."));

children.push(bulletP("Pengumpulan Data. Setelah dasar teori dipahami, data dikumpulkan dari Survei Kesehatan Indonesia (SKI) 2023 untuk subset Provinsi Kalimantan Tengah, lalu dipraproses melalui penyaringan rentang usia, konversi satuan, pemetaan kode, pembuangan *noise*, dan *listwise deletion* agar siap dipakai oleh model."));

children.push(bulletP("Perancangan Sistem. Tahap perancangan sistem menyusun kerangka kerja *monitoring* risiko *stunting* melalui desain mesin *fuzzy* Sugeno, prosedur penyesuaian batas fungsi keanggotaan menggunakan *Genetic Algorithm*, serta arsitektur antarmuka web berbasis Flask."));

children.push(bulletP("Implementasi Sistem. Implementasi sistem menerapkan desain yang telah dirancang, mengintegrasikan komponen pelatihan model dengan komponen *deployment* berbasis web, serta melakukan penyesuaian agar parameter hasil pelatihan dapat langsung dipakai oleh sistem web."));

children.push(bulletP("Pengujian dan Evaluasi. Setelah sistem dikembangkan, pengujian dan evaluasi dilakukan untuk menilai konsistensi keluaran sistem terhadap *ground truth* WHO HAZ. Evaluasi menggunakan *confusion matrix*, korelasi *rank* Spearman, *prevalence stratification ratio*, dan analisis trayektori pertumbuhan."));

children.push(bulletP("Penulisan Laporan. Tahap akhir merangkum dan menyimpulkan seluruh temuan penelitian, kemudian menyusunnya secara sistematis dalam laporan agar memberikan gambaran utuh mengenai hasil yang dicapai."));

// ============================================================
// 3.2 ALAT DAN BAHAN
// ============================================================
children.push(H2("3.2   Alat dan Bahan"));

children.push(H3("3.2.1   Perangkat Keras"));
children.push(P("Perangkat keras yang digunakan dalam penelitian ini berupa MacBook Pro dengan *chipset* Apple M4, RAM 16 GB, dan GPU terintegrasi 10-*core*. Konfigurasi ini sudah memadai untuk menjalankan praproses data, eksekusi mesin *fuzzy* Sugeno, optimasi *Genetic Algorithm* dengan lima *seed* berbeda, sekaligus menjalankan *server* Flask secara lokal saat pengembangan."));

children.push(H3("3.2.2   Perangkat Lunak"));
children.push(P("Pada penelitian ini, Figma digunakan untuk merancang *user interface* (UI). Visual Studio Code (VS Code) dipilih sebagai *Integrated Development Environment* (IDE). Bahasa pemrograman Python digunakan sebagai *logic program* dalam praproses data dan implementasi model bersama dengan pustaka pendukungnya, yaitu *pandas* dan *numpy* untuk pemrosesan data, *deap* untuk implementasi *Genetic Algorithm*, *scipy* untuk perhitungan korelasi *rank* Spearman, serta *matplotlib* untuk visualisasi. Google Colab digunakan untuk eksperimen dan prototipe, serta *framework* web Flask untuk integrasi antarmuka berbasis web sehingga sistem *monitoring* dapat diakses secara praktis oleh pengguna akhir."));

children.push(H3("3.2.3   Dataset"));
children.push(P("Dataset yang digunakan dalam penelitian ini bersumber dari Survei Kesehatan Indonesia (SKI) 2023, yaitu survei nasional yang merekam kondisi kesehatan masyarakat Indonesia sebagaimana telah dibahas pada Subbab 2.1.4. Data diperoleh melalui portal resmi Kementerian Kesehatan (https://layanandata.kemkes.go.id) dan dibatasi pada subset Provinsi Kalimantan Tengah yang berisi 1.679 *record* dengan 370 variabel. Karena dataset asli sudah dianonimisasi pada tahap pengambilan oleh penyelenggara SKI 2023, tidak diperlukan persetujuan etik tambahan untuk penggunaan kembali pada penelitian ini. Selain dataset SKI 2023, penelitian ini juga memakai empat tabel parameter LMS WHO yang dibedakan menurut jenis kelamin (laki-laki dan perempuan) dan kelompok usia (0-24 bulan dan 25-60 bulan) sebagai rujukan perhitungan Z-Score."));

// ============================================================
// 3.3 PERANCANGAN SISTEM
// ============================================================
children.push(H2("3.3   Perancangan Sistem"));

children.push(P("Perancangan sistem dilakukan untuk mengembangkan kerangka kerja yang sistematis dalam *monitoring* risiko *stunting* secara efektif melalui tahapan terstruktur. Sistem dirancang sebagai *pipeline* yang menerima enam variabel masukan, menghasilkan dua label WHO HAZ (saat lahir dan saat pengukuran), serta memetakan masukan ke *risk score* kontinu pada rentang 0-100 dan empat *priority flag* (Normal, *Stunting Low*, *Stunting Medium*, *Stunting High*) melalui mesin *fuzzy* Sugeno yang batas fungsi keanggotaannya disesuaikan dengan *Genetic Algorithm*. Alur perancangan sistem ditampilkan pada Gambar 3.2 berikut."));

children.push(placeholderFig("Gambar 3.2 Alur Perancangan Sistem"));
children.push(captionFig("Gambar 3.2 Alur Perancangan Sistem Monitoring Risiko Stunting"));

// ---- 3.3.1 Input Data ----
children.push(H3("3.3.1   Input Data"));

children.push(P("Dataset SKI 2023 Kalimantan Tengah dimuat dengan pustaka *pandas* dan diseleksi untuk hanya menyimpan tujuh kolom yang relevan dengan penelitian, yakni satu kolom identitas (ID_balita) dan enam kolom variabel yang ditampilkan pada Tabel III.1 berikut ini."));

children.push(captionTab("Tabel III.1 Variabel Penelitian"));
children.push(makeTable(
    ["No.", "Variabel", "Kolom Asli SKI 2023", "Keterangan"],
    [
        ["1", "Usia Balita", "Umur_BLN", "Usia balita dalam satuan bulan"],
        ["2", "Jenis Kelamin Balita", "Jenis Kelamin", "1 = Laki-laki, 2 = Perempuan"],
        ["3", "Panjang Badan Lahir (PBL)", "Panjang Badan Lahir", "Nilai panjang badan saat lahir (cm)"],
        ["4", "Tinggi Badan", "Panjang Badan Anak", "Nilai tinggi/panjang badan saat pengukuran (cm)"],
        ["5", "Berat Badan", "Berat Badan Saat Data Diambil", "Nilai berat badan saat pengukuran (gram)"],
        ["6", "MP-ASI", "I43_MP ASI", "Waktu pengenalan MP-ASI dalam satuan bulan"],
    ],
    [800, 2600, 2900, 3060]
));
children.push(P("Pemilihan enam variabel ini didasarkan pada relevansi klinis terhadap *stunting* sebagaimana telah dibahas pada Subbab 2.1.5 dan dikonsultasikan dengan ahli kesehatan, yaitu Heti Ira Ayue, SST., M.Keb., kandidat doktor dan dosen Program Studi Profesi Bidan Jurusan Kebidanan Poltekkes Kemenkes Palangka Raya."));

// ---- 3.3.2 Preprocessing ----
children.push(H3("3.3.2   Preprocessing"));

children.push(P("Praproses dilakukan untuk memastikan hanya data yang bersih dan relevan yang masuk ke dalam mesin *fuzzy* Sugeno. Tahapannya meliputi beberapa hal, yakni sebagai berikut."));

children.push(bulletP("Transformasi dan rekoding. Tahap ini memastikan setiap variabel berada dalam satuan yang benar, yaitu usia dalam bulan (dengan label '<1 bulan' untuk balita yang berusia di bawah satu bulan), berat badan dikonversi dari kilogram ke gram, sedangkan kode huruf MP-ASI pada SKI 2023 (a hingga n) dipetakan ke satuan bulan dengan rentang a = 0,25 bulan (7 hari) hingga n = 12 bulan."));

children.push(bulletP("Penyaringan rentang usia. Data dibatasi hanya untuk balita berusia 0-59 bulan sesuai definisi operasional *stunting* dalam SKI 2023, sedangkan baris di luar rentang tersebut dikeluarkan dari analisis."));

children.push(bulletP("*Data reduction* (seleksi fitur). Dari 370 kolom dalam SKI 2023, hanya sembilan kolom inti yang dipertahankan, yaitu ID_balita, usia_bulan, usia_label, jenis_kelamin, tinggi_badan_cm, berat_badan_gram, pbl_cm, mpasi_bulan, dan mpasi_label."));

children.push(bulletP("Pembuangan *noise* dan *listwise deletion*. *Sentinel codes* yang lazim muncul pada SKI 2023 seperti 88, 88.8, 99, dan 99.9 untuk variabel panjang badan lahir dan berat badan diubah menjadi nilai kosong, sedangkan kode MP-ASI di luar daftar a-n diperlakukan sebagai *noise*. Setelah itu, setiap baris yang memiliki minimal satu nilai kosong pada enam variabel inti dihapus secara *listwise* (Tabachnick & Fidell, 2014) karena keenam variabel tersebut diperlukan oleh perhitungan WHO HAZ dan mesin *fuzzy*."));

children.push(bulletP("Pembentukan kelas turunan. Dua kolom turunan dibentuk untuk menopang analisis *monitoring*, yaitu jarak_lahir_ukur (selisih usia antara saat lahir dengan saat pengukuran) dan jarak_mpasi_survei (selisih antara usia saat pengukuran dengan waktu pengenalan MP-ASI), keduanya dalam satuan bulan."));

children.push(P("Setelah seluruh langkah praproses dijalankan, dari 1.679 baris awal hanya 389 baris yang lolos seluruh kriteria. Penyusutan ukuran sampel pada tiap tahap praproses dirangkum pada Tabel III.2 berikut."));

children.push(captionTab("Tabel III.2 Tahapan Praproses dan Penyusutan Sampel"));
children.push(makeTable(
    ["No.", "Tahap Praproses", "Sisa Baris"],
    [
        ["1", "Data mentah (SKI 2023 Kalimantan Tengah)", "1.679"],
        ["2", "Transformasi & rekoding satuan", "1.679"],
        ["3", "Filter usia 0-59 bulan", "1.679"],
        ["4", "Data reduction (seleksi fitur)", "1.679"],
        ["5", "Listwise deletion", "389"],
        ["6", "Pembentukan kelas turunan", "389"],
    ],
    [800, 6500, 1560]
));
children.push(P("Tabel III.2 menunjukkan bahwa *listwise deletion* menjadi tahap yang paling memengaruhi ukuran sampel, terutama akibat banyaknya nilai kosong pada kolom MP-ASI dan panjang badan lahir. Dataset bersih sebanyak 389 balita inilah yang menjadi dasar perhitungan WHO HAZ, evaluasi mesin *fuzzy* Sugeno pada konfigurasi awal, serta tahap penyesuaian batas fungsi keanggotaan dengan bantuan GA."));

// ---- 3.3.3 Perhitungan Z-Score dan Stunting Labelling ----
children.push(H3("3.3.3   Perhitungan Z-Score dan Stunting Labelling"));

children.push(P("Status *stunting* sebagai *ground truth* dihitung secara manual berdasarkan standar WHO *Child Growth Standards* dengan rumus Z-Score WHO berbasis LMS sebagaimana telah dijelaskan pada Subbab 2.1.3 (Persamaan II.2). Langkah perhitungan adalah sebagai berikut."));

children.push(bulletP("Menentukan parameter rujukan (L, M, S) dari tabel antropometri WHO sesuai jenis kelamin dan usia balita. Penelitian ini memakai empat tabel LMS WHO yang dirangkum pada Tabel III.3."));

children.push(bulletP("Menghitung *Height-for-Age Z-score* (HAZ) menggunakan rumus LMS WHO pada Persamaan II.2 dengan masukan panjang/tinggi badan balita."));

children.push(bulletP("Mengonversi nilai HAZ ke dalam dua kategori status gizi, yaitu *Stunting* (Z < −2 SD) dan Normal (Z ≥ −2 SD) sebagaimana ambang WHO pada Tabel II.2."));

children.push(P("Karena penelitian ini bersifat *monitoring* dan bukan klasifikasi pada satu titik waktu, perhitungan Z-Score dilakukan pada dua momen. Z-Score saat lahir dihitung dengan parameter LMS pada usia 0 bulan menggunakan panjang badan lahir sebagai masukan, sedangkan Z-Score saat pengukuran dihitung dengan parameter LMS pada usia aktual menggunakan tinggi badan saat ukur. Z-Score saat pengukuran inilah yang berperan sebagai *ground truth* untuk evaluasi sistem, sedangkan Z-Score saat lahir dipakai untuk analisis trayektori *monitoring* pada Bab IV."));

children.push(captionTab("Tabel III.3 Daftar Tabel LMS WHO Rujukan Perhitungan Z-Score"));
children.push(makeTable(
    ["No.", "Jenis Kelamin", "Rentang Usia", "Nama Berkas"],
    [
        ["1", "Laki-laki", "0-24 bulan", "lhfa_boys_0to2years_zscores.xlsx"],
        ["2", "Laki-laki", "25-60 bulan", "lhfa_boys_2to5years_zscores.xlsx"],
        ["3", "Perempuan", "0-24 bulan", "lhfa_girls_0to2years_zscores.xlsx"],
        ["4", "Perempuan", "25-60 bulan", "lhfa_girls_2to5years_zscores.xlsx"],
    ],
    [800, 2400, 2400, 3760]
));
children.push(P("Tabel III.3 menampilkan empat tabel LMS WHO yang dipakai untuk pencarian parameter L, M, dan S. Pencarian baris LMS dilakukan dengan kunci pasangan jenis kelamin dan usia dalam bulan, sehingga setiap balita memperoleh satu pasangan parameter yang khas sesuai karakteristik tubuhnya."));

// ---- 3.3.4 Sistem Logika Fuzzy Sugeno ----
children.push(H3("3.3.4   Sistem Logika Fuzzy Sugeno"));

children.push(P("Mesin logika *fuzzy* Sugeno menjadi inti penalaran sistem yang menghasilkan *risk score* dan *priority flag*. Proses inferensi *fuzzy* terdiri dari tiga komponen, yakni fuzzifikasi, basis aturan IF-THEN, dan defuzzifikasi sebagaimana dasar teorinya telah diuraikan pada Subbab 2.1.6. Pada penelitian ini, jenis kelamin berperan sebagai variabel pengalih (*switching variable*) yang menentukan himpunan fungsi keanggotaan tinggi badan dan berat badan yang dipakai. Variabel yang difuzzifikasi adalah usia, tinggi badan, berat badan, dan waktu pengenalan MP-ASI, sedangkan panjang badan lahir tidak difuzzifikasi karena perannya sudah terwakili melalui Z-Score saat lahir. Penjelasan tiap komponen mesin *fuzzy* Sugeno dijabarkan pada sub-subbab berikut."));

children.push(H4("3.3.4.1   Fuzzifikasi"));

children.push(P("Fuzzifikasi merupakan proses mengubah setiap nilai *input* tegas menjadi derajat keanggotaan. Penelitian ini menggunakan bentuk fungsi keanggotaan trapesium dan segitiga (Persamaan II.4 dan II.5 pada Bab II) karena keduanya sederhana, mudah ditafsirkan, dan stabil dalam menggambarkan transisi gradual antarstatus pertumbuhan anak. Variabel masukan beserta label linguistik, bentuk fungsi keanggotaan, dan titik batas awal sebelum optimasi ditampilkan pada Tabel III.4 berikut."));

children.push(captionTab("Tabel III.4 Konfigurasi Awal Fungsi Keanggotaan Variabel Masukan"));
children.push(makeTable(
    ["Variabel", "Label Linguistik", "Bentuk", "Titik Batas Awal"],
    [
        ["Usia (bulan)", "Muda, Sedang, Tua", "Trap-Tri-Trap", "(0,0,6,18); (12,30,48); (42,52,59,59)"],
        ["TB Laki-laki (cm)", "Pendek, Normal, Tinggi", "Trap-Trap-Trap", "(38,38,55,70); (60,70,100,110); (100,110,130,130)"],
        ["TB Perempuan (cm)", "Pendek, Normal, Tinggi", "Trap-Trap-Trap", "(38,38,53,68); (58,68,98,108); (98,108,130,130)"],
        ["BB Laki-laki (g)", "Kurang, Normal, Lebih, Obesitas", "Trap-Trap-Tri-Trap", "(1500,1500,5000,8000); (6000,8000,15000,18000); (15000,18000,22000); (20000,23000,30000,30000)"],
        ["BB Perempuan (g)", "Kurang, Normal, Lebih, Obesitas", "Trap-Trap-Tri-Trap", "(1500,1500,4500,7500); (5500,7500,14000,17000); (14000,17000,21000); (19000,22000,28000,28000)"],
        ["MP-ASI (bulan)", "Terlalu Cepat, Normal, Lambat", "Trap-Trap-Trap", "(0,0,3,5); (4,5.5,6.5,8); (7,8,12,12)"],
    ],
    [1800, 2200, 1800, 3560]
));

children.push(P("Tinggi badan dan berat badan sengaja dipisahkan berdasarkan jenis kelamin agar fungsi keanggotaan yang dipakai sesuai dengan pola pertumbuhan balita laki-laki dan perempuan menurut standar WHO. Titik batas awal pada Tabel III.4 disusun berdasarkan rujukan standar pertumbuhan WHO dan rekomendasi MP-ASI ESPGHAN, dan menjadi posisi awal *Genetic Algorithm* sebagaimana dijelaskan pada Sub-subbab 3.3.5. Visualisasi grafik fungsi keanggotaan untuk masing-masing variabel dapat dilihat pada Gambar 3.3 hingga Gambar 3.6 berikut."));

children.push(placeholderFig("Gambar 3.3 Grafik Fungsi Keanggotaan Variabel Usia"));
children.push(captionFig("Gambar 3.3 Grafik Fungsi Keanggotaan Variabel Usia"));
children.push(placeholderFig("Gambar 3.4 Grafik Fungsi Keanggotaan Variabel Tinggi Badan (a) Laki-laki (b) Perempuan"));
children.push(captionFig("Gambar 3.4 Grafik Fungsi Keanggotaan Variabel Tinggi Badan (a) Laki-laki (b) Perempuan"));
children.push(placeholderFig("Gambar 3.5 Grafik Fungsi Keanggotaan Variabel Berat Badan (a) Laki-laki (b) Perempuan"));
children.push(captionFig("Gambar 3.5 Grafik Fungsi Keanggotaan Variabel Berat Badan (a) Laki-laki (b) Perempuan"));
children.push(placeholderFig("Gambar 3.6 Grafik Fungsi Keanggotaan Variabel Waktu Pengenalan MP-ASI"));
children.push(captionFig("Gambar 3.6 Grafik Fungsi Keanggotaan Variabel Waktu Pengenalan MP-ASI"));

children.push(P("Selain fungsi keanggotaan variabel masukan, sistem juga memiliki fungsi keanggotaan keluaran yang memetakan *risk score* tegas ke dalam empat *priority flag* sebagaimana ditampilkan pada Tabel III.5 dan Gambar 3.7. Berbeda dengan fungsi keanggotaan masukan yang akan disesuaikan oleh GA, fungsi keanggotaan keluaran bersifat tetap agar interpretasi *priority flag* tetap konsisten sepanjang penelitian dan sepanjang penggunaan sistem pada tahap *deployment*."));

children.push(captionTab("Tabel III.5 Fungsi Keanggotaan Output (Risk Score → Priority Flag)"));
children.push(makeTable(
    ["No.", "Priority Flag", "Bentuk", "Titik Batas"],
    [
        ["1", "Normal", "Trapesium", "(0, 0, 15, 30)"],
        ["2", "Stunting Low", "Segitiga", "(15, 30, 60)"],
        ["3", "Stunting Medium", "Segitiga", "(45, 60, 75)"],
        ["4", "Stunting High", "Trapesium", "(75, 90, 100, 100)"],
    ],
    [800, 2500, 2200, 3860]
));
children.push(placeholderFig("Gambar 3.7 Grafik Fungsi Keanggotaan Output (Risk Score → Priority Flag)"));
children.push(captionFig("Gambar 3.7 Grafik Fungsi Keanggotaan Output (Risk Score → Priority Flag)"));

children.push(H4("3.3.4.2   Basis Aturan IF-THEN"));

children.push(P("Aturan dasar adalah inti dari sistem *fuzzy*, yaitu kumpulan logika *IF-THEN* yang menggabungkan seluruh variabel *fuzzy* untuk menghasilkan nilai konsekuen berupa *risk score* sebagaimana bentuk umumnya diuraikan pada Persamaan II.6 di Bab II. Penelitian ini menggunakan 108 aturan IF-THEN yang merupakan hasil produk Kartesian dari tiga himpunan usia, tiga himpunan tinggi badan, empat himpunan berat badan, dan tiga himpunan MP-ASI (3 × 3 × 4 × 3 = 108 aturan). Setiap aturan kemudian diinstansiasi dua kali, yakni untuk himpunan tinggi-berat laki-laki dan untuk himpunan tinggi-berat perempuan, sehingga total aturan efektif yang berjalan pada sistem berjumlah 216 aturan. Distribusi 108 aturan ke dalam empat kelas risiko beserta nilai konsekuennya ditampilkan pada Tabel III.6 berikut."));

children.push(captionTab("Tabel III.6 Distribusi 108 Aturan IF-THEN per Kelas Risiko"));
children.push(makeTable(
    ["No.", "Kelas Risiko", "Konsekuen (c)", "Jumlah Aturan", "Total Efektif (× 2 JK)"],
    [
        ["1", "STUNTING HIGH", "90", "10", "20"],
        ["2", "STUNTING MEDIUM", "65", "30", "60"],
        ["3", "STUNTING LOW", "35", "20", "40"],
        ["4", "Normal", "10", "48", "96"],
        [{ text: "Total", bold: true, align: AlignmentType.LEFT }, " - ", " - ", { text: "108", bold: true }, { text: "216", bold: true }],
    ],
    [800, 2600, 1800, 2000, 2160]
));
children.push(P("Mayoritas aturan (48 dari 108) memberikan konsekuen kelas Normal, sedangkan aturan-aturan kelas *Stunting High* terkonsentrasi pada kombinasi yang mengindikasikan risiko paling tinggi seperti Usia Tua, TB Pendek, dan BB Kurang. Nilai konsekuen 10, 35, 65, dan 90 dipilih agar tepat berada di pusat rentang fungsi keanggotaan keluaran pada Tabel III.5, sehingga aktivasi aturan tunggal langsung memetakan ke *priority flag* yang sesuai sedangkan aktivasi banyak aturan menghasilkan distribusi *risk score* yang halus. Sebagai gambaran, aturan \"IF Usia is Tua AND TB is Pendek AND BB is Kurang AND MP-ASI is Terlalu Cepat THEN y = 90\" memetakan ke kelas *Stunting High*, sedangkan aturan \"IF Usia is Sedang AND TB is Normal AND BB is Normal AND MP-ASI is Normal THEN y = 10\" memetakan ke kelas Normal. Operator agregasi premis (*AND*) yang dipakai adalah *minimum* (*Gödel* t-*norm*), sehingga bobot aktivasi tiap aturan diperoleh dari nilai derajat keanggotaan terkecil di antara keempat premisnya."));

children.push(H4("3.3.4.3   Defuzzifikasi"));

children.push(P("Defuzzifikasi merupakan tahap akhir dalam sistem inferensi *fuzzy* yang mengubah keluaran *fuzzy* hasil aktivasi aturan menjadi nilai tegas berupa *risk score*. Penelitian ini menggunakan metode *weighted average* Sugeno orde nol sebagaimana telah dijabarkan pada Persamaan II.7 di Bab II. *Risk score* dinormalisasi pada rentang 0-100 dengan catatan semakin kecil nilai maka semakin baik kondisi balita, begitupun sebaliknya. Skor ini kemudian dipetakan menjadi *priority flag* (Normal, *Stunting Low*, *Stunting Medium*, dan *Stunting High*) melalui fungsi keanggotaan keluaran pada Tabel III.5 dengan kaidah *winner-takes-all*. Apabila tidak ada aturan yang teraktivasi (jumlah bobot sama dengan nol), sistem memberikan nilai *default* sebesar 35 yang setara dengan kelas *Stunting Low* sebagai posisi konservatif agar tidak melewatkan kemungkinan adanya risiko. Konfigurasi mesin *fuzzy* Sugeno pada Sub-subbab 3.3.4 ini berperan sebagai *baseline* sebelum batas fungsi keanggotaannya disesuaikan oleh GA pada Sub-subbab 3.3.5."));

// ---- 3.3.5 GA ----
children.push(H3("3.3.5   Penyesuaian Batas Fungsi Keanggotaan dengan Genetic Algorithm"));

children.push(P("Setelah mesin *fuzzy* Sugeno selesai dirancang, satu-satunya komponen yang masih dapat disesuaikan dengan data adalah posisi titik batas fungsi keanggotaan variabel masukan. *Genetic Algorithm* (GA) dipakai sebagai mekanisme pembantu pelatihan akhir untuk menggeser titik batas tersebut agar batas-batas linguistik yang sebelumnya disusun secara konseptual menjadi lebih cocok dengan distribusi nyata data Kalimantan Tengah. Perlu ditegaskan bahwa GA tidak mengubah 108 aturan pada Tabel III.6 maupun bentuk fungsi keanggotaan pada Tabel III.4; struktur penalaran tetap sepenuhnya milik sistem *fuzzy* Sugeno. Konsep dasar GA dan ketiga operatornya (seleksi, *crossover*, dan mutasi) telah dijelaskan pada Subbab 2.1.7. Bagian ini menguraikan rancangan spesifik GA pada penelitian, yakni representasi individu, fungsi *fitness*, hiperparameter, serta alur optimasi multi-*seed*."));

children.push(H4("3.3.5.1   Representasi Individu (Genome)"));

children.push(P("Setiap individu pada populasi GA direpresentasikan sebagai vektor berisi 77 gen bertipe bilangan riil. Setiap gen berkorespondensi dengan satu titik batas pada salah satu fungsi keanggotaan, dengan rincian sebagaimana ditampilkan pada Tabel III.7."));

children.push(captionTab("Tabel III.7 Komposisi 77 Gen pada Individu Genetic Algorithm"));
children.push(makeTable(
    ["No.", "Variabel", "Jumlah Label × Parameter", "Jumlah Gen"],
    [
        ["1", "Usia", "Muda (4) + Sedang (3) + Tua (4)", "11"],
        ["2", "TB Laki-laki", "Pendek (4) + Normal (4) + Tinggi (4)", "12"],
        ["3", "TB Perempuan", "Pendek (4) + Normal (4) + Tinggi (4)", "12"],
        ["4", "BB Laki-laki", "Kurang (4) + Normal (4) + Lebih (3) + Obesitas (4)", "15"],
        ["5", "BB Perempuan", "Kurang (4) + Normal (4) + Lebih (3) + Obesitas (4)", "15"],
        ["6", "MP-ASI", "Terlalu Cepat (4) + Normal (4) + Lambat (4)", "12"],
        [{ text: "Total", bold: true, align: AlignmentType.LEFT }, " - ", " - ", { text: "77", bold: true }],
    ],
    [800, 3000, 4300, 1060]
));
children.push(P("Setiap gen dibatasi pada rentang yang masuk akal secara klinis agar hasil optimasi tidak menyimpang dari rentang biologis pertumbuhan balita. Pembatasan ini juga menjamin keteraturan urutan titik batas (a ≤ b ≤ c ≤ d pada trapesium) tetap terjaga sepanjang proses optimasi. Sebagai gambaran, contoh rentang batas gen untuk beberapa parameter ditampilkan pada Tabel III.8 berikut."));

children.push(captionTab("Tabel III.8 Contoh Rentang Batas Gen untuk Beberapa Parameter Fungsi Keanggotaan"));
children.push(makeTable(
    ["No.", "Variabel & Label", "Parameter", "Rentang Batas (lo, hi)"],
    [
        ["1", "Usia - Muda", "a, b, c, d", "(0,0); (0,0); (3,12); (12,24)"],
        ["2", "Usia - Sedang", "a, b, c", "(6,18); (20,36); (36,54)"],
        ["3", "TB Laki-laki - Pendek", "a, b, c, d", "(35,42); (35,42); (45,60); (60,80)"],
        ["4", "TB Perempuan - Normal", "a, b, c, d", "(48,65); (58,75); (88,103); (98,113)"],
        ["5", "BB Laki-laki - Kurang", "a, b, c, d", "(1000,2000); (1000,2000); (3000,6000); (6000,10000)"],
        ["6", "MP-ASI - Normal", "a, b, c, d", "(3,5); (4,6); (5.5,7.5); (6,9)"],
    ],
    [800, 2900, 1900, 3560]
));

children.push(H4("3.3.5.2   Fungsi Fitness"));

children.push(P("Fungsi *fitness* yang dipakai pada penelitian ini merupakan kombinasi tertimbang antara F1-score dan akurasi sebagaimana telah dirumuskan pada Persamaan II.8 di Bab II, dengan bobot w₁ = 0,6 untuk F1-score dan w₂ = 0,4 untuk akurasi. Bobot asimetris ini sengaja dipilih karena dataset bersifat tidak seimbang dengan rasio kelas Normal terhadap *Stunting* sekitar 5:1, sehingga F1-score lebih informatif daripada akurasi semata. Penambahan bobot akurasi sebesar 0,4 berperan sebagai penyeimbang yang mencegah konfigurasi degeneratif yang hanya memaksimalkan *recall* dengan mengklasifikasikan hampir semua balita sebagai *Stunting*."));

children.push(H4("3.3.5.3   Hiperparameter GA"));

children.push(P("Nilai hiperparameter GA yang dipakai pada penelitian ini ditampilkan pada Tabel III.9. Ukuran populasi 80 dan jumlah generasi 60 dipilih berdasarkan eksperimen pendahuluan yang menunjukkan konvergensi tercapai pada kisaran tersebut tanpa biaya komputasi berlebih. Probabilitas *crossover* 0,6 dan probabilitas mutasi 0,3 dipilih untuk menjaga keseimbangan antara eksploitasi solusi yang sudah baik dan eksplorasi ruang pencarian baru, dengan probabilitas *crossover* dijaga lebih besar agar arah utama pencarian tetap pada rekombinasi solusi yang sudah promising sedangkan mutasi berperan sebagai diversifier untuk mencegah konvergensi prematur."));

children.push(captionTab("Tabel III.9 Hiperparameter Genetic Algorithm"));
children.push(makeTable(
    ["No.", "Hiperparameter", "Nilai"],
    [
        ["1", "Ukuran populasi", "80"],
        ["2", "Jumlah generasi", "60"],
        ["3", "Probabilitas crossover (cxpb)", "0,6"],
        ["4", "Probabilitas mutasi (mutpb)", "0,3"],
        ["5", "Seleksi", "Tournament (size = 3)"],
        ["6", "Crossover", "Blend crossover (cxBlend), α = 0,3"],
        ["7", "Mutasi", "Bounded Gaussian, indpb = 0,15, σ = 0,15 × range"],
        ["8", "Random seed", "{42, 123, 777, 2024, 31415}"],
        ["9", "Panjang genome", "77 gen"],
        ["10", "Strategi elitisme", "Individu nilai awal disisipkan sebagai P[0]"],
    ],
    [800, 4280, 4280]
));

children.push(H4("3.3.5.4   Alur Optimasi dan Pelaksanaan Multi-Seed"));

children.push(P("Untuk mengurangi ketergantungan hasil pada satu titik awal acak, prosedur GA dijalankan dengan lima nilai *seed* berbeda yaitu {42, 123, 777, 2024, 31415}. Setiap *seed* menghasilkan satu kandidat konfigurasi batas fungsi keanggotaan yang kemudian dievaluasi pada subset uji, dan kandidat dengan F1-score terbaik dipilih sebagai konfigurasi final yang akan disimpan dalam berkas *trained_model.json*. Berkas inilah yang nantinya dimuat oleh sistem web pada tahap *deployment*. Alur prosedur GA per *seed* dirumuskan dalam Algoritma III.1 berikut."));

children.push(codeLine("Algoritma III.1  Optimasi Batas Fungsi Keanggotaan dengan GA"));
children.push(codeLine("Input  : subset pelatihan D_tr, batas tiap gen B = {(lo_i, hi_i)} i=1..77"));
children.push(codeLine("Output : individu terbaik theta*"));
children.push(codeLine("1.  Inisialisasi populasi P berukuran 80 di sekitar nilai awal MF (Tabel III.4),"));
children.push(codeLine("    setiap gen dipotong ke batas (lo_i, hi_i)"));
children.push(codeLine("2.  Sisipkan nilai awal MF sebagai P[0] (elitisme awal)"));
children.push(codeLine("3.  Hitung F(theta) untuk setiap theta di P dengan Persamaan II.8 pada D_tr"));
children.push(codeLine("4.  for g = 1 to 60 do"));
children.push(codeLine("5.      P_sel <- tournament_select(P, k = 3)"));
children.push(codeLine("6.      P_off <- cxBlend(P_sel, alpha = 0,3) dengan cxpb = 0,6"));
children.push(codeLine("7.      P_off <- mutate_gaussian(P_off, indpb = 0,15, sigma = 0,15 x range) dengan mutpb = 0,3"));
children.push(codeLine("8.      Potong setiap gen P_off ke (lo_i, hi_i)"));
children.push(codeLine("9.      Hitung F(theta) untuk individu baru"));
children.push(codeLine("10.     P <- (mu + lambda) replacement antara P dan P_off"));
children.push(codeLine("11. end for"));
children.push(codeLine("12. return arg max F(theta), theta di P"));

// ---- 3.3.6 Web ----
children.push(H3("3.3.6   Integrasi Sistem Berbasis Web"));

children.push(P("Setelah mesin *fuzzy* Sugeno dengan batas fungsi keanggotaan hasil penyesuaian GA tersimpan dalam berkas *trained_model.json*, sistem diintegrasikan ke dalam aplikasi web berbasis *framework* Flask agar dapat diakses lintas perangkat oleh pengguna akhir. Sisi *backend* memuat berkas *trained_model.json* pada saat *startup* sehingga inferensi pada sistem web langsung memakai konfigurasi hasil pelatihan tanpa perlu menjalankan ulang proses GA, sedangkan sisi *frontend* diimplementasikan dengan HTML, CSS, dan JavaScript mengacu pada rancangan tata letak yang sebelumnya dibuat di Figma. Sisi *backend* menyediakan dua *endpoint* utama, yakni *endpoint* inferensi yang menerima enam variabel masukan dan mengembalikan Z-Score saat lahir, Z-Score saat pengukuran, *risk score*, dan *priority flag*, serta *endpoint* /retrain yang memicu prosedur penyesuaian batas fungsi keanggotaan ulang menggunakan GA pada dataset gabungan riwayat pengguna dan dataset pelatihan awal sehingga sistem dapat menyesuaikan diri terhadap data baru. Sistem web terdiri dari empat halaman utama sebagaimana dijabarkan pada Tabel III.10."));

children.push(captionTab("Tabel III.10 Daftar Halaman Sistem Web"));
children.push(makeTable(
    ["No.", "Halaman", "Fungsi"],
    [
        ["1", "Landing", "Pintu masuk sistem yang memuat ringkasan layanan, fitur utama, daftar pertanyaan umum, kebijakan sistem, kontak, dan tombol panggil aksi ke halaman form input."],
        ["2", "Form Input", "Menampilkan formulir enam medan masukan, yakni usia, jenis kelamin, tinggi badan saat ukur, berat badan saat ukur, panjang badan lahir, dan waktu pengenalan MP-ASI, dengan dukungan kontrol popup pemilihan untuk jenis kelamin dan MP-ASI."],
        ["3", "Hasil Perhitungan", "Menampilkan Z-Score saat lahir, Z-Score saat pengukuran, label WHO, *risk score* 0-100, dan *priority flag* hasil sistem *fuzzy* Sugeno."],
        ["4", "Riwayat dan Retrain", "Merekap seluruh inferensi pengguna dan menyediakan tombol *Retrain Model* yang memicu *endpoint* /retrain untuk menjalankan ulang prosedur GA. Tombol aktif setelah riwayat pengguna mencapai minimal lima record."],
    ],
    [800, 2400, 6160]
));
children.push(P("Rancangan tata letak masing-masing halaman dapat dilihat pada Gambar 3.8 hingga Gambar 3.11. Karena hasil inferensi pada halaman Hasil Perhitungan diturunkan langsung dari mesin *fuzzy* Sugeno yang sama dengan tahap pelatihan, koherensi antara tahap pelatihan dan tahap *deployment* tetap terjaga."));

children.push(placeholderFig("Gambar 3.8 Rancangan Tata Letak Halaman Landing"));
children.push(captionFig("Gambar 3.8 Rancangan Tata Letak Halaman Landing"));
children.push(placeholderFig("Gambar 3.9 Rancangan Tata Letak Halaman Form Input"));
children.push(captionFig("Gambar 3.9 Rancangan Tata Letak Halaman Form Input"));
children.push(placeholderFig("Gambar 3.10 Rancangan Tata Letak Halaman Hasil Perhitungan"));
children.push(captionFig("Gambar 3.10 Rancangan Tata Letak Halaman Hasil Perhitungan"));
children.push(placeholderFig("Gambar 3.11 Rancangan Tata Letak Halaman Riwayat dan Retrain"));
children.push(captionFig("Gambar 3.11 Rancangan Tata Letak Halaman Riwayat dan Retrain"));

// ============================================================
// 3.4 PENGUJIAN DAN EVALUASI
// ============================================================
children.push(H2("3.4   Pengujian dan Evaluasi"));

children.push(P("Pengujian dan evaluasi dalam penelitian ini bertujuan untuk menilai performa sistem *monitoring* risiko *stunting* berbasis *fuzzy* Sugeno dengan membandingkan keluaran sistem terhadap *ground truth* WHO HAZ. Karena sistem menghasilkan *risk score* kontinu dan *priority flag* bertingkat, evaluasi tidak cukup hanya pada metrik biner. Tahap evaluasi disusun atas lima aspek yang saling melengkapi, yaitu skema pembagian data, metrik *confusion matrix*, validasi statistik dengan korelasi *rank* Spearman, *prevalence stratification ratio* per *priority flag*, dan analisis trayektori pertumbuhan dari momen lahir ke momen pengukuran."));

children.push(H3("3.4.1   Skema Pembagian Data"));

children.push(P("Penelitian ini menggunakan *stratified hold-out validation* dengan proporsi 80% data pelatihan dan 20% data pengujian, di mana pembagian dilakukan secara terstratifikasi berdasarkan kelas WHO HAZ (Normal dan *Stunting*) dengan *random seed* sebesar 42 agar pembagian dapat direproduksi. Skema ini dipilih karena dataset bersifat tidak seimbang sehingga pembagian acak biasa berisiko menghasilkan subset uji yang tidak representatif. Komposisi pembagian data ditampilkan pada Tabel III.11 berikut."));

children.push(captionTab("Tabel III.11 Komposisi Stratified Hold-out 80/20 pada Dataset Akhir"));
children.push(makeTable(
    ["No.", "Subset", "Total", "Normal", "Stunting"],
    [
        ["1", "Keseluruhan", "389", "325", "64"],
        ["2", "Pelatihan (80%)", "311", "260", "51"],
        ["3", "Pengujian (20%)", "78", "65", "13"],
    ],
    [800, 3200, 1760, 1760, 1760]
));

children.push(P("Strategi pembagian data dipakai berbeda untuk dua kondisi evaluasi yang akan dilaporkan pada Bab IV. Untuk evaluasi mesin *fuzzy* Sugeno pada konfigurasi awal (sebelum batas fungsi keanggotaan disesuaikan oleh GA), pengujian dijalankan terhadap keseluruhan 389 balita tanpa pembagian karena tujuannya hanya mengamati perilaku mesin *fuzzy* dengan posisi titik batas *default* pada Tabel III.4. Hasil pada konfigurasi awal ini menjadi acuan pembanding terhadap kinerja sistem setelah batas fungsi keanggotaan disesuaikan oleh GA. Sementara itu, untuk evaluasi mesin *fuzzy* Sugeno setelah penyesuaian GA, subset pelatihan dipakai oleh fungsi *fitness* GA untuk menilai setiap kandidat konfigurasi batas sedangkan subset pengujian disisihkan sepanjang proses optimasi dan baru dipakai untuk estimasi performa pada data baru. Pendekatan ini memastikan kinerja yang dilaporkan pada Bab IV bukan akibat hafalan data pelatihan."));

children.push(H3("3.4.2   Metrik Confusion Matrix"));

children.push(P("Empat metrik klasifikasi yang dipakai adalah akurasi, *precision*, *recall*, dan F1-score, dengan rumus masing-masing telah dijabarkan pada Subbab 2.1.9 (Persamaan II.9 hingga II.12). Prediksi sistem dipetakan menjadi biner dengan menggabungkan *Stunting Low*, *Stunting Medium*, dan *Stunting High* sebagai kelas *Stunting*, sedangkan *priority flag* Normal sebagai kelas Normal. Adapun dasar interpretasi keempat komponen *confusion matrix* untuk konteks penelitian ini adalah sebagai berikut."));

children.push(bulletP("TP (*True Positive*): balita yang dilabeli *Stunting* oleh WHO HAZ dan diprediksi *Stunting* oleh sistem."));
children.push(bulletP("FP (*False Positive*): balita yang dilabeli Normal oleh WHO HAZ tetapi diprediksi *Stunting* oleh sistem."));
children.push(bulletP("FN (*False Negative*): balita yang dilabeli *Stunting* oleh WHO HAZ tetapi diprediksi Normal oleh sistem."));
children.push(bulletP("TN (*True Negative*): balita yang dilabeli Normal oleh WHO HAZ dan diprediksi Normal oleh sistem."));

children.push(P("Pada konteks *monitoring* risiko *stunting*, F1-score menjadi metrik utama karena lebih peka terhadap kelas minoritas (kelas *Stunting* lebih sedikit dibanding Normal) sedangkan akurasi tetap dilaporkan untuk memastikan sistem tidak hanya tepat dalam mendeteksi risiko tinggi tetapi juga stabil dalam meminimalkan kesalahan klasifikasi pada seluruh kelas. *Recall* dipandang lebih penting daripada *precision* karena biaya melewatkan kasus *stunting* (FN) lebih besar daripada biaya pemeriksaan tambahan terhadap balita yang dibendera waspada (FP)."));

children.push(H3("3.4.3   Validasi Statistik dengan Korelasi Rank Spearman"));

children.push(P("Karena sistem menghasilkan *risk score* kontinu pada rentang 0-100, evaluasi tidak cukup hanya bersandar pada metrik biner dari *confusion matrix*. Korelasi *rank* Spearman sebagaimana telah diuraikan pada Persamaan II.13 di Bab II dipakai untuk mengukur kesesuaian monotonik antara *risk score* yang dihasilkan sistem dengan WHO HAZ saat pengukuran sebagai *ground truth*. Tanda negatif pada nilai ρ yang diharapkan menunjukkan bahwa *risk score* yang lebih tinggi berasosiasi dengan HAZ yang lebih rendah, sejalan dengan desain sistem yang memetakan kondisi pertumbuhan terlambat ke *risk score* yang lebih tinggi. Pengujian signifikansi dilakukan dengan ambang p < 0,05 pada seluruh 389 balita agar memanfaatkan informasi pengurutan secara penuh, bukan sekadar pada subset uji."));

children.push(H3("3.4.4   Prevalence Stratification Ratio per Priority Flag"));

children.push(P("Karena sistem menghasilkan empat *priority flag* yang bertujuan menstratifikasi prioritas intervensi, evaluasi juga mengukur seberapa baik sistem mengakumulasi kasus *stunting* nyata pada *priority flag* yang lebih tinggi. *Prevalence stratification ratio* sebagaimana telah dirumuskan pada Persamaan II.14 di Bab II dihitung untuk setiap *priority flag* dengan prevalensi *baseline* keseluruhan dataset sebesar 16,5%. Rasio lebih besar dari 1 menunjukkan *priority flag* tersebut mengakumulasi kasus *stunting* lebih tinggi daripada rata-rata acak, sedangkan rasio kurang dari 1 menunjukkan *priority flag* tersebut bersifat protektif. Metrik ini penting karena memvalidasi kemampuan sistem dalam mengarahkan prioritas intervensi pada tingkat komunitas, terutama untuk kelas *Stunting High* yang seharusnya menampung kasus *stunting* paling pekat."));

children.push(H3("3.4.5   Analisis Trayektori Pertumbuhan"));

children.push(P("Sebagai bagian dari kerangka *monitoring*, evaluasi juga membandingkan label WHO HAZ saat lahir dengan *priority flag* sistem saat pengukuran untuk setiap balita. Perbandingan tersebut menghasilkan empat pola trayektori sebagaimana dijabarkan pada Tabel III.12 yang masing-masing merepresentasikan dinamika pertumbuhan berbeda. Analisis ini melengkapi metrik klasifikasi biner dengan memberikan gambaran kondisi pertumbuhan nyata balita dari momen lahir hingga momen pengukuran, sehingga sistem tidak hanya berfungsi sebagai pendeteksi pada satu titik waktu melainkan sebagai alat pemantauan perubahan kondisi balita."));

children.push(captionTab("Tabel III.12 Definisi Pola Trayektori Pertumbuhan Lahir ke Pengukuran"));
children.push(makeTable(
    ["No.", "Pola Trayektori", "Definisi", "Interpretasi"],
    [
        ["1", "Lahir Stunting → Sistem Stunting", "WHO saat lahir = Stunting DAN priority flag ≠ Normal", "Risiko persisten"],
        ["2", "Lahir Stunting → Sistem Normal", "WHO saat lahir = Stunting DAN priority flag = Normal", "Pemulihan positif"],
        ["3", "Lahir Normal → Sistem Stunting", "WHO saat lahir = Normal DAN priority flag ≠ Normal", "Pertumbuhan menurun"],
        ["4", "Lahir Normal → Sistem Normal", "WHO saat lahir = Normal DAN priority flag = Normal", "Konsisten sehat"],
    ],
    [800, 3000, 4000, 2360]
));
children.push(P("Hasil empirik untuk setiap pola trayektori beserta interpretasinya akan dilaporkan pada Bab IV sebagai bagian dari validasi terhadap kondisi pertumbuhan nyata balita di Kalimantan Tengah."));

// ============================================================
// 3.5 JADWAL PENELITIAN
// ============================================================
children.push(H2("3.5   Jadwal Penelitian"));
children.push(P("Rencana jadwal kegiatan penelitian dapat dilihat pada Tabel III.13 berikut. Penjadwalan disusun untuk memastikan seluruh tahapan dari studi literatur hingga penyusunan laporan akhir berjalan secara terurut dan terukur."));

children.push(captionTab("Tabel III.13 Jadwal Pelaksanaan Penelitian"));
children.push(makeTable(
    ["No.", "Kegiatan", "Bulan 1", "Bulan 2", "Bulan 3", "Bulan 4", "Bulan 5", "Bulan 6"],
    [
        ["1", "Studi Literatur", "■", "■", "", "", "", ""],
        ["2", "Pengumpulan Data", "", "■", "", "", "", ""],
        ["3", "Perancangan Sistem", "", "■", "■", "", "", ""],
        ["4", "Implementasi Sistem", "", "", "■", "■", "", ""],
        ["5", "Pengujian dan Evaluasi", "", "", "", "■", "■", ""],
        ["6", "Penyusunan Laporan", "", "", "", "", "■", "■"],
    ],
    [600, 2400, 860, 860, 860, 860, 860, 860]
));

const doc = new Document({
    creator: "Gerrard Sebastian",
    title: "BAB III - Metodologi Penelitian (FIX v4 - Gaya BUKU TA VER 01)",
    styles: {
        default: { document: { run: { font: FONT, size: SIZE } } },
        paragraphStyles: [
            {
                id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 28, bold: true, font: FONT },
                paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 }
            },
            {
                id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 26, bold: true, font: FONT },
                paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 }
            },
            {
                id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 24, bold: true, font: FONT },
                paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 2 }
            },
            {
                id: "Heading4", name: "Heading 4", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 24, bold: true, font: FONT },
                paragraph: { spacing: { before: 160, after: 100 }, outlineLevel: 3 }
            },
        ],
    },
    numbering: {
        config: [{
            reference: "bullets",
            levels: [{
                level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: { left: 720, hanging: 360 } } }
            }]
        }]
    },
    sections: [{
        properties: {
            page: {
                size: { width: 11906, height: 16838 },
                margin: { top: 1701, right: 1701, bottom: 1701, left: 2268 }
            }
        },
        children: children
    }]
});

Packer.toBuffer(doc).then(buf => {
    fs.writeFileSync("/Users/sinitygs/Projects2/train/BAB_III_FIX.docx", buf);
    console.log("OK file size: " + buf.length);
});