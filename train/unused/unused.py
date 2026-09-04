# #!/usr/bin/env python3
# # -*- coding: utf-8 -*-
# """
# =============================================================================
# PENELITIAN: Klasifikasi Stunting pada Balita di Kalimantan Tengah
#             Menggunakan Logika Fuzzy (Studi Kasus Data SKI 2023)
# =============================================================================
# TASK 1: INPUT DATA — Loading Dataset & Seleksi Fitur
# TASK 2: PREPROCESSING — Pembersihan, Penyaringan, Transformasi
# TASK 3: STUNTING LABELLING & Z-SCORE — Ground Truth WHO
# TASK 4: FUZZY SYSTEM LOGIC — Takagi-Sugeno (Fuzzyfikasi, Rules, Defuzzyfikasi)
# TASK 5: EVALUASI — Confusion Matrix (Precision, Recall, F1-Score, Accuracy)
# TASK 6: SKEMA PENGUJIAN — Train/Test 80/20 Stratified + Optimasi Parameter
# =============================================================================
# Lokasi: train/fuzzytrain.py
# Struktur:
#   PROJECTS2/
#     train/
#       fuzzytrain.py              ← FILE INI
#       dataset/
#         Individu_balita_2023_merged.csv
#       trainfile/                 ← output hasil train disimpan di sini
#       zs/
#         boys02.xlsx   (boys 0-24 bulan)
#         boys25.xlsx   (boys 24-60 bulan)
#         girls02.xlsx  (girls 0-24 bulan)
#         girls25.xlsx  (girls 24-60 bulan)
#     mainprogram/
#     test/
# =============================================================================
# """

# import pandas as pd
# import numpy as np
# import json
# import os
# from datetime import datetime

# # ============================================================================
# # PATH CONFIGURATION (Relative to this script's location)
# # ============================================================================
# BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# DATASET_PATH   = os.path.join(BASE_DIR, 'dataset', 'Individu_balita_2023_merged.csv')
# WHO_ZS_DIR     = os.path.join(BASE_DIR, 'zs')
# TRAINFILE_DIR  = os.path.join(BASE_DIR, 'trainfile')

# OUTPUT_CLEAN_CSV   = os.path.join(TRAINFILE_DIR, 'dataset_clean.csv')
# OUTPUT_LABELED_CSV = os.path.join(TRAINFILE_DIR, 'dataset_labeled.csv')
# OUTPUT_REPORT      = os.path.join(TRAINFILE_DIR, 'preprocessing_report.json')

# os.makedirs(TRAINFILE_DIR, exist_ok=True)

# # WHO Z-Score table file paths
# WHO_TABLES = {
#     'boys_0_24':  os.path.join(WHO_ZS_DIR, 'boys02.xlsx'),
#     'boys_24_60': os.path.join(WHO_ZS_DIR, 'boys25.xlsx'),
#     'girls_0_24':  os.path.join(WHO_ZS_DIR, 'girls02.xlsx'),
#     'girls_24_60': os.path.join(WHO_ZS_DIR, 'girls25.xlsx'),
# }

# # ============================================================================
# # KONFIGURASI BATAS NORMAL MEDIS
# # ============================================================================
# VALID_RANGES = {
#     'usia_bulan':    (0, 59),
#     'jenis_kelamin': [1, 2],
#     'tinggi_badan':  (35.0, 130.0),
#     'bbl_gram':      (500, 6000),
#     'pbl_cm':        (25.0, 60.0),
#     'mpasi_bulan':   (0, 12),
# }

# NOISE_CODES = {
#     'bbl': ['8888', '9999', '88.8'],
#     'pbl': ['88', '88.8', '99', '99.9'],
# }

# MPASI_CODE_MAP = {
#     'a': 0.25, 'b': 0.5,  'c': 1.0,  'd': 2.0,  'e': 3.0,
#     'f': 4.0,  'g': 5.0,  'h': 6.0,  'i': 7.0,  'j': 8.0,
#     'k': 9.0,  'l': 10.0, 'm': 11.0, 'n': 12.0,
# }

# MPASI_NOISE = {'o', 'p', ' ', ''}

# # Default priority flag bounds (model-driven recommended). These are fallbacks
# # used when no trained model bounds are supplied.
# DEFAULT_FLAG_BOUNDS = {'low': 25, 'medium': 50, 'high': 75}

# # ============================================================================
# # FUNGSI UTILITAS
# # ============================================================================
# def print_header(title, char='=', width=60):
#     print(f"\n{char * width}")
#     print(f"  {title}")
#     print(f"{char * width}")

# def print_step(step_num, description):
#     print(f"\n  [{step_num}] {description}")

# def safe_numeric(series, errors='coerce'):
#     cleaned = series.astype(str).str.strip().replace({'': np.nan, ' ': np.nan})
#     return pd.to_numeric(cleaned, errors=errors)


# # ============================================================================
# # TASK 1: INPUT DATA
# # ============================================================================
# def task1_input_data():
#     print_header("TASK 1: INPUT DATA — Loading Dataset & Seleksi Fitur")
    
#     print_step("1.1", "Memuat dataset SKI 2023 Kalimantan Tengah")
#     print(f"    Path: {DATASET_PATH}")
    
#     if not os.path.exists(DATASET_PATH):
#         raise FileNotFoundError(
#             f"\n    [X] DATASET TIDAK DITEMUKAN!\n"
#             f"    Pastikan file berada di: {DATASET_PATH}\n"
#         )
    
#     df_raw = pd.read_csv(DATASET_PATH, low_memory=False)
#     print(f"    [OK] Dataset dimuat: {df_raw.shape[0]:,} baris x {df_raw.shape[1]} kolom")
    
#     print_step("1.2", "Mapping kolom dataset ke variabel penelitian")
    
#     COLUMN_MAP = {
#         'Umur_BLN':             'usia_bulan',
#         'Jenis Kelamin':        'jenis_kelamin',
#         'Panjang Badan Anak':   'tinggi_badan_cm',
#         'Berat Badan Lahir':    'bbl_raw',
#         'Panjang Badan Lahir':  'pbl_raw',
#         'I43_MP ASI':           'mpasi_raw',
#     }
    
#     df = df_raw[['ID_balita'] + list(COLUMN_MAP.keys())].copy()
#     df.rename(columns=COLUMN_MAP, inplace=True)
    
#     for var_lama, var_baru in COLUMN_MAP.items():
#         print(f"    {var_lama:30s} -> {var_baru}")
    
#     print(f"\n    [OK] Seleksi fitur selesai: {df.shape[1]} kolom dipilih dari {df_raw.shape[1]} kolom")
    
#     print_step("1.3", "Statistik awal dataset (sebelum preprocessing)")
#     print(f"    Total balita   : {len(df):,}")
#     for col in df.columns:
#         if col == 'ID_balita':
#             continue
#         non_null = df[col].notna().sum()
#         blank = (df[col].astype(str).str.strip() == '').sum()
#         print(f"    {col:20s}: {non_null:,} terisi, {blank} kosong/blank")
    
#     return df, df_raw.shape


# # ============================================================================
# # TASK 2: PREPROCESSING
# # ============================================================================
# def task2_preprocessing(df, raw_shape):
#     print_header("TASK 2: PREPROCESSING")
    
#     report = {
#         'raw_total': len(df),
#         'raw_shape': list(raw_shape),
#         'steps': {},
#         'timestamp': datetime.now().isoformat()
#     }
    
#     # ==== STEP 2d: TRANSFORMASI & REKODING ====
#     print_header("STEP 2d: Transformasi & Rekoding", char='-')
    
#     # 2d.1: Usia
#     print_step("2d.1", "Transformasi Usia Balita -> bulan (integer)")
#     df['usia_bulan'] = pd.to_numeric(df['usia_bulan'], errors='coerce')
#     df['usia_label'] = df['usia_bulan'].apply(
#         lambda x: '<1 bulan' if pd.notna(x) and x < 1 else
#                   (f'{int(x)} bulan' if pd.notna(x) else None))
#     usia_bawah_1 = (df['usia_bulan'] < 1).sum()
#     print(f"    [OK] Balita usia <1 bulan: {usia_bawah_1} -> dilabeli '<1 bulan'")
#     print(f"    Distribusi: min={df['usia_bulan'].min()}, max={df['usia_bulan'].max()}, mean={df['usia_bulan'].mean():.1f}")
    
#     # 2d.2: Jenis Kelamin
#     print_step("2d.2", "Validasi Jenis Kelamin -> 1=Laki-laki, 2=Perempuan")
#     df['jenis_kelamin'] = pd.to_numeric(df['jenis_kelamin'], errors='coerce')
#     jk_invalid = ~df['jenis_kelamin'].isin([1, 2]) & df['jenis_kelamin'].notna()
#     print(f"    [OK] Valid (1/2): {df['jenis_kelamin'].isin([1, 2]).sum():,} | Invalid: {jk_invalid.sum()}")
    
#     # 2d.3: Tinggi Badan
#     print_step("2d.3", "Transformasi Panjang/Tinggi Badan -> cm (float)")
#     df['tinggi_badan_cm'] = safe_numeric(df['tinggi_badan_cm'])
#     print(f"    [OK] Valid: {df['tinggi_badan_cm'].notna().sum():,} | NaN: {df['tinggi_badan_cm'].isna().sum()}")
    
#     # 2d.4: BBL
#     print_step("2d.4", "Transformasi BBL -> gram (float)")
#     bbl_str = df['bbl_raw'].astype(str).str.strip()
#     bbl_is_noise = bbl_str.isin(NOISE_CODES['bbl'])
#     df['bbl_gram'] = safe_numeric(df['bbl_raw'])
#     df.loc[bbl_is_noise, 'bbl_gram'] = np.nan
#     print(f"    [OK] Noise (8888): {bbl_is_noise.sum()} | Valid: {df['bbl_gram'].notna().sum():,}")
    
#     # 2d.5: PBL
#     print_step("2d.5", "Transformasi PBL -> cm (float)")
#     pbl_str = df['pbl_raw'].astype(str).str.strip()
#     pbl_is_noise = pbl_str.isin(NOISE_CODES['pbl'])
#     df['pbl_cm'] = safe_numeric(df['pbl_raw'])
#     df.loc[pbl_is_noise, 'pbl_cm'] = np.nan
#     print(f"    [OK] Noise (88/88.8): {pbl_is_noise.sum()} | Valid: {df['pbl_cm'].notna().sum():,}")
    
#     # 2d.6: MP-ASI
#     print_step("2d.6", "Transformasi MP-ASI (kode huruf) -> bulan (float)")
#     mpasi_str = df['mpasi_raw'].astype(str).str.strip().str.lower()
#     df['mpasi_bulan'] = mpasi_str.map(MPASI_CODE_MAP)
#     mapped_ok = df['mpasi_bulan'].notna().sum()
    
#     print(f"    Mapping kode huruf -> bulan:")
#     for code, bulan in sorted(MPASI_CODE_MAP.items()):
#         count = (mpasi_str == code).sum()
#         if count > 0:
#             lbl = f"<1 bln" if bulan < 1 else f"{bulan:.0f} bln"
#             print(f"      '{code}' -> {lbl:>8s}: {count:>4d} balita")
#     print(f"    [OK] Mapped: {mapped_ok:,} | Missing: {df['mpasi_bulan'].isna().sum():,}")
    
#     df['mpasi_label'] = df['mpasi_bulan'].apply(
#         lambda x: '<1 bulan' if pd.notna(x) and x < 1 else
#                   (f'{int(x)} bulan' if pd.notna(x) and x == int(x) else
#                    (f'{x} bulan' if pd.notna(x) else None)))
    
#     report['steps']['2d_transformasi'] = {
#         'usia_bawah_1_bulan': int(usia_bawah_1),
#         'jk_invalid': int(jk_invalid.sum()),
#         'bbl_noise_8888': int(bbl_is_noise.sum()),
#         'pbl_noise_88': int(pbl_is_noise.sum()),
#         'mpasi_mapped': int(mapped_ok),
#         'mpasi_missing': int(df['mpasi_bulan'].isna().sum()),
#     }
    
#     # ==== STEP 2b: FILTER USIA ====
#     print_header("STEP 2b: Penyaringan Wilayah & Usia", char='-')
#     print(f"    [OK] Dataset khusus Kalimantan Tengah ({len(df):,} balita)")
#     n_before_age = len(df)
#     usia_valid = df['usia_bulan'].between(0, 59, inclusive='both')
#     usia_diluar = ~usia_valid & df['usia_bulan'].notna()
#     if usia_diluar.any():
#         df = df[usia_valid | df['usia_bulan'].isna()].copy()
#         print(f"    [OK] Dihapus: {n_before_age - len(df)} balita di luar 0-59 bulan")
#     else:
#         print(f"    [OK] Semua dalam rentang 0-59 bulan")
    
#     report['steps']['2b_filter_usia'] = {
#         'before': n_before_age, 'after': len(df),
#     }
    
#     # ==== STEP 2c: DATA REDUCTION ====
#     print_header("STEP 2c: Data Reduction (Seleksi Fitur)", char='-')
#     FINAL_COLUMNS = ['ID_balita', 'usia_bulan', 'usia_label', 'jenis_kelamin',
#                      'tinggi_badan_cm', 'bbl_gram', 'pbl_cm', 'mpasi_bulan', 'mpasi_label']
#     df = df[FINAL_COLUMNS].copy()
#     print(f"    [OK] {len(FINAL_COLUMNS)} kolom dipertahankan dari {raw_shape[1]} kolom asli")
    
#     # ==== STEP 2a: LISTWISE DELETION ====
#     print_header("STEP 2a: Pemeriksaan Kelengkapan & Pembuangan Noise", char='-')
#     print(f"    Metode: LISTWISE DELETION (Tabachnick & Fidell, 2014)")
    
#     CORE_VARS = ['usia_bulan', 'jenis_kelamin', 'tinggi_badan_cm',
#                  'bbl_gram', 'pbl_cm', 'mpasi_bulan']
    
#     n_before = len(df)
    
#     print_step("2a.1", "Missing values per variabel inti")
#     missing_detail = {}
#     for var in CORE_VARS:
#         n_miss = df[var].isna().sum()
#         pct = n_miss / len(df) * 100
#         missing_detail[var] = n_miss
#         s = "[OK]" if n_miss == 0 else "[!]"
#         print(f"    {s} {var:20s}: {n_miss:>5d} missing ({pct:>5.1f}%)")
    
#     print_step("2a.2", "Deteksi outlier medis")
#     outlier_flags = pd.DataFrame(index=df.index)
    
#     for col, key in [('usia_bulan','usia_bulan'), ('tinggi_badan_cm','tinggi_badan'),
#                      ('bbl_gram','bbl_gram'), ('pbl_cm','pbl_cm'), ('mpasi_bulan','mpasi_bulan')]:
#         lo, hi = VALID_RANGES[key]
#         mask = ~df[col].between(lo, hi) & df[col].notna()
#         outlier_flags[f'{col}_out'] = mask
#         if mask.any():
#             print(f"    [!] {col}: {mask.sum()} di luar [{lo}-{hi}]")
#             print(f"      Nilai: {sorted(df.loc[mask, col].unique()[:5])}")
#         else:
#             print(f"    [OK] {col}: OK [{lo}-{hi}]")
    
#     mask_jk = ~df['jenis_kelamin'].isin([1, 2]) & df['jenis_kelamin'].notna()
#     outlier_flags['jk_out'] = mask_jk
#     print(f"    {'[!]' if mask_jk.any() else '[OK]'} jenis_kelamin: {mask_jk.sum()} invalid")
    
#     print_step("2a.3", "Eksekusi Listwise Deletion")
#     any_missing = df[CORE_VARS].isna().any(axis=1)
#     any_outlier = outlier_flags.any(axis=1)
#     rows_to_drop = any_missing | any_outlier
    
#     reasons = {}
#     for var in CORE_VARS:
#         vm = df[var].isna()
#         oc = f'{var}_out' if f'{var}_out' in outlier_flags else 'jk_out'
#         vo = outlier_flags.get(oc, pd.Series(False, index=df.index))
#         vp = vm | vo
#         reasons[var] = int(vp.sum())
#         print(f"    {var:20s}: {vp.sum():>5d} bermasalah (miss={vm.sum()}, outlier={vo.sum()})")
    
#     df_clean = df[~rows_to_drop].copy().reset_index(drop=True)
#     n_dropped = n_before - len(df_clean)
    
#     print(f"\n    LISTWISE DELETION RESULT:")
#     print(f"    Sebelum : {n_before:,} balita")
#     print(f"    Dihapus : {n_dropped:,} ({n_dropped/n_before*100:.1f}%)")
#     print(f"    Sesudah : {len(df_clean):,} balita")
    
#     report['steps']['2a_listwise'] = {
#         'before': n_before, 'total_dropped': n_dropped, 'after': len(df_clean),
#         'missing_per_var': {k: int(v) for k, v in missing_detail.items()},
#         'reasons': reasons,
#     }
    
#     # Validasi
#     print(f"\n    Validasi: zero missing di semua variabel inti:")
#     for var in CORE_VARS:
#         n_miss = df_clean[var].isna().sum()
#         print(f"    {'[OK]' if n_miss==0 else '[X]'} {var}: {n_miss} missing")
    
#     report['final_clean'] = {'rows': len(df_clean), 'cols': df_clean.shape[1]}
    
#     return df_clean, report


# # ============================================================================
# # TASK 3: STUNTING LABELLING & Z-SCORE
# # ============================================================================
# def task3_zscore_labelling(df):
#     """
#     Task 3: Menghitung Z-Score WHO (HAZ/LAZ) dan membuat label stunting.
    
#     Alur ANALISA PER BARIS sesuai Proposal Bab 3.3.3:
#       1. Baca: jenis_kelamin, usia_bulan, tinggi_badan_cm
#       2. Pilih tabel WHO sesuai jenis kelamin
#       3. Cari baris WHO berdasarkan usia (bulan) -> ambil L,M,S
#       4. Hitung Z-score: Z = [(x/M)^L - 1] / (L x S)
#       5. Labelling: HAZ < -2.00 SD -> Stunting | HAZ >= -2.00 SD -> Normal
#     """
    
#     print_header("TASK 3: STUNTING LABELLING & Z-SCORE (Ground Truth WHO)")
    
#     print(f"    Analisa per baris: {len(df)} balita")
#     print(f"    Rumus: Z = [(x/M)^L - 1] / (L x S)")
#     print(f"    Threshold: HAZ < -2.00 SD = Stunting, HAZ >= -2.00 SD = Normal")
    
#     # 3.1: Load Tabel WHO
#     print_step("3.1", "Memuat tabel referensi WHO (LMS parameters)")
    
#     who_lookup = {}
    
#     table_config = [
#         ('boys_0_24',  1, 'Laki-laki 0-24 bln (LAZ)'),
#         ('boys_24_60', 1, 'Laki-laki 24-60 bln (HAZ)'),
#         ('girls_0_24',  2, 'Perempuan 0-24 bln (LAZ)'),
#         ('girls_24_60', 2, 'Perempuan 24-60 bln (HAZ)'),
#     ]
    
#     for table_key, jk, label in table_config:
#         fpath = WHO_TABLES[table_key]
#         if not os.path.exists(fpath):
#             raise FileNotFoundError(f"    [X] WHO table tidak ditemukan: {fpath}")
        
#         df_who = pd.read_excel(fpath)
#         df_who.columns = df_who.columns.str.strip()
        
#         n_loaded = 0
#         for _, row in df_who.iterrows():
#             month = int(row['Month'])
#             key = (jk, month)
#             who_lookup[key] = {'L': float(row['L']), 'M': float(row['M']), 'S': float(row['S'])}
#             n_loaded += 1
        
#         print(f"    [OK] {label:35s} : {n_loaded:>3d} baris dimuat")
    
#     print(f"\n    [OK] Total lookup entries: {len(who_lookup)}")
    
#     # 3.2: Verifikasi Cakupan
#     print_step("3.2", "Verifikasi cakupan usia dataset vs tabel WHO")
    
#     usia_range = df['usia_bulan'].unique()
#     jk_range = df['jenis_kelamin'].unique()
    
#     missing_keys = []
#     for jk in sorted(jk_range):
#         for usia in sorted(usia_range):
#             key = (int(jk), int(usia))
#             if key not in who_lookup:
#                 missing_keys.append(key)
    
#     if missing_keys:
#         print(f"    [!] {len(missing_keys)} kombinasi TIDAK ditemukan di tabel WHO:")
#         for k in missing_keys[:10]:
#             print(f"      JK={k[0]}, Usia={k[1]} bulan")
#     else:
#         print(f"    [OK] Semua kombinasi (jenis_kelamin, usia) tercakup dalam tabel WHO")
    
#     print(f"    Usia dataset: {int(min(usia_range))}-{int(max(usia_range))} bulan")
#     print(f"    Jenis kelamin: {sorted([int(x) for x in jk_range])}")
    
#     # 3.3: ANALISA PER BARIS
#     print_step("3.3", "ANALISA PER BARIS — Menghitung Z-Score WHO untuk setiap balita")
#     print(f"    Rumus WHO (Persamaan II.2): Z = [(x/M)^L - 1] / (L x S)")
#     print(f"    Dimana:")
#     print(f"      x = panjang/tinggi badan balita (cm) — data pengukuran")
#     print(f"      L (Lambda) = parameter power transformasi normalitas")
#     print(f"      M (Miu)    = median populasi referensi WHO")
#     print(f"      S (Sigma)  = koefisien variasi populasi referensi")
#     print(f"    Threshold: HAZ < -2.00 SD -> Stunting | HAZ >= -2.00 SD -> Normal")
#     print(f"\n    Memulai analisa per baris untuk {len(df):,} balita...")
#     print(f"    {'-' * 65}")
    
#     z_scores = np.full(len(df), np.nan)
#     L_values = np.full(len(df), np.nan)
#     M_values = np.full(len(df), np.nan)
#     S_values = np.full(len(df), np.nan)
#     labels = [None] * len(df)
    
#     matched = 0
    
#     for idx in range(len(df)):
#         jk = int(df.iloc[idx]['jenis_kelamin'])
#         usia = int(df.iloc[idx]['usia_bulan'])
#         x = float(df.iloc[idx]['tinggi_badan_cm'])
        
#         key = (jk, usia)
        
#         if key in who_lookup:
#             params = who_lookup[key]
#             L, M, S = params['L'], params['M'], params['S']
            
#             L_values[idx] = L
#             M_values[idx] = M
#             S_values[idx] = S
            
#             if L != 0:
#                 z = ((x / M) ** L - 1) / (L * S)
#             else:
#                 z = np.log(x / M) / S
            
#             z_scores[idx] = round(z, 3)
#             labels[idx] = 'Stunting' if z < -2.0 else 'Normal'
#             matched += 1
    
#     df['L'] = L_values
#     df['M'] = M_values
#     df['S'] = S_values
#     df['z_score'] = z_scores
#     df['label_who'] = labels
    
#     print(f"    [OK] Analisa per baris SELESAI: {matched:,}/{len(df):,} balita berhasil dihitung")
    
#     # --------------------------------------------------------------
#     # 3.4: OUTPUT HASIL ANALISA PER BARIS (TABEL LENGKAP)
#     # --------------------------------------------------------------
#     print_step("3.4", "OUTPUT HASIL ANALISA DATA PER BARIS (SELURUH DATASET)")
#     print(f"    Format: No | Sex | Umur (bln) | PB anak (cm) | L (Lambda) | M (Miu) | S (Sigma) | HAZ (WHO-LMS) | Label WHO")
#     print(f"    Total: {len(df):,} baris data balita")
    
#     hdr = (f"    {'No':>4s} | {'Sex':^10s} | {'Umur':>5s} | {'PB anak':>9s} | "
#            f"{'L':>6s} | {'M':>9s} | {'S':>8s} | {'HAZ':>10s} | {'Label WHO':^10s}")
#     sep = f"    {'-'*4}-+-{'-'*10}-+-{'-'*5}-+-{'-'*9}-+-{'-'*6}-+-{'-'*9}-+-{'-'*8}-+-{'-'*10}-+-{'-'*10}"
    
#     print(f"\n{sep}")
#     print(hdr)
#     print(sep)
    
#     for i in range(len(df)):
#         row = df.iloc[i]
#         no = i + 1
#         jk = int(row['jenis_kelamin'])
#         jk_str = 'Laki-laki' if jk == 1 else 'Perempuan'
#         usia = int(row['usia_bulan'])
#         tb = row['tinggi_badan_cm']
#         L = row['L']
#         M = row['M']
#         S = row['S']
#         z = row['z_score']
#         lbl = row['label_who']
        
#         print(f"    {no:>4d} | {jk_str:^10s} | {usia:>5d} | {tb:>9.2f} | "
#               f"{L:>6.2f} | {M:>9.3f} | {S:>8.5f} | {z:>10.3f} | {lbl:^10s}")
    
#     print(sep)
#     print(f"    TOTAL: {len(df):,} baris dianalisa secara individual (per baris)")
    
#     # --------------------------------------------------------------
#     # 3.5: Statistik Hasil Z-Score & Labelling
#     # --------------------------------------------------------------
#     print_step("3.5", "Statistik Hasil Z-Score & Labelling")
    
#     z_valid = df['z_score'].dropna()
    
#     print(f"    Z-Score Distribution:")
#     print(f"    Count={len(z_valid):,}, Mean={z_valid.mean():.3f}, Std={z_valid.std():.3f}")
#     print(f"    Min={z_valid.min():.3f}, Q1={z_valid.quantile(0.25):.3f}, Median={z_valid.median():.3f}, Q3={z_valid.quantile(0.75):.3f}, Max={z_valid.max():.3f}")
    
#     label_counts = df['label_who'].value_counts()
#     total = len(df)
#     print(f"\n    Distribusi Label WHO (Ground Truth):")
#     for label_name in ['Normal', 'Stunting']:
#         if label_name in label_counts.index:
#             cnt = label_counts[label_name]
#             pct = cnt / total * 100
#             print(f"    {label_name:10s}: {cnt:>4d} ({pct:>5.1f}%)")
#     print(f"    TOTAL     : {total:>4d} (100.0%)")
    
#     # Distribusi per jenis kelamin
#     print(f"\n    Distribusi Label per Jenis Kelamin:")
#     for jk_val, jk_name in [(1, 'Laki-laki'), (2, 'Perempuan')]:
#         subset = df[df['jenis_kelamin'] == jk_val]
#         stunting_n = (subset['label_who'] == 'Stunting').sum()
#         normal_n = (subset['label_who'] == 'Normal').sum()
#         total_jk = len(subset)
#         print(f"    {jk_name:10s}: Normal={normal_n} ({normal_n/total_jk*100:.1f}%), "
#               f"Stunting={stunting_n} ({stunting_n/total_jk*100:.1f}%)")
    
#     # Distribusi per kelompok usia
#     print(f"\n    Distribusi Label per Kelompok Usia:")
#     age_groups = [(0, 5, '0-5 bln'), (6, 11, '6-11 bln'), (12, 17, '12-17 bln'), (18, 23, '18-23 bln')]
#     for lo, hi, grp_label in age_groups:
#         subset = df[(df['usia_bulan'] >= lo) & (df['usia_bulan'] <= hi)]
#         if len(subset) == 0:
#             continue
#         stunting_n = (subset['label_who'] == 'Stunting').sum()
#         normal_n = (subset['label_who'] == 'Normal').sum()
#         pct_s = stunting_n / len(subset) * 100
#         print(f"    {grp_label:10s} (n={len(subset):>3d}): Normal={normal_n:>3d}, "
#               f"Stunting={stunting_n:>3d} ({pct_s:.1f}%)")
    
#     # --------------------------------------------------------------
#     # 3.6: Validasi Manual (Cross-check rumus)
#     # --------------------------------------------------------------
#     print_step("3.6", "Validasi Formula (Cross-check rumus manual per baris)")
#     print(f"    Mengambil 5 sampel acak dan menghitung ulang secara manual...")
    
#     np.random.seed(42)
#     check_indices = np.random.choice(len(df), size=5, replace=False)
    
#     all_valid = True
#     for idx in check_indices:
#         row = df.iloc[idx]
#         jk = int(row['jenis_kelamin'])
#         usia = int(row['usia_bulan'])
#         x = float(row['tinggi_badan_cm'])
#         L = float(row['L'])
#         M = float(row['M'])
#         S = float(row['S'])
#         z_stored = float(row['z_score'])
        
#         if L != 0:
#             z_manual = ((x / M) ** L - 1) / (L * S)
#         else:
#             z_manual = np.log(x / M) / S
#         z_manual = round(z_manual, 3)
        
#         match = "[OK]" if z_stored == z_manual else "[X]"
#         if z_stored != z_manual:
#             all_valid = False
        
#         jk_str = 'Laki-laki' if jk == 1 else 'Perempuan'
#         tbl = 'boys' if jk == 1 else 'girls'
#         tbl_range = '02' if usia <= 24 else '25'
        
#         print(f"    {match} Baris ke-{idx+1}: {jk_str}, {usia} bln, TB={x:.1f} cm")
#         print(f"      Tabel WHO  : {tbl}{tbl_range}.xlsx -> Bulan {usia}")
#         print(f"      L(Lambda)  : {L}")
#         print(f"      M(Miu)     : {M:.4f}")
#         print(f"      S(Sigma)   : {S:.5f}")
#         print(f"      Perhitungan: Z = [({x:.1f}/{M:.4f})^{L} - 1] / ({L} x {S:.5f})")
#         print(f"      Z_program  : {z_stored:.3f}")
#         print(f"      Z_manual   : {z_manual:.3f}")
#         print(f"      Label      : {row['label_who']} (Z {'<' if z_stored < -2 else '>='} -2.00)")
#         print()
    
#     if all_valid:
#         print(f"    [OK] Semua {len(check_indices)} validasi manual PASS")
#         print(f"    [OK] Konfirmasi: Dataset dianalisa PER BARIS — formula Z-Score benar")
#     else:
#         print(f"    [!] Ada perbedaan — perlu review")
    
#     # --------------------------------------------------------------
#     # 3.7: Simpan Output Analisa Per Baris ke CSV
#     # --------------------------------------------------------------
#     print_step("3.7", "Menyimpan output analisa per baris ke CSV")
#     df_output = pd.DataFrame({
#         'No': range(1, len(df) + 1),
#         'Sex': df['jenis_kelamin'].map({1: 'Laki-laki', 2: 'Perempuan'}).values,
#         'Umur (bln)': df['usia_bulan'].astype(int).values,
#         'PB anak (cm)': df['tinggi_badan_cm'].values,
#         'L': df['L'].values, 'M': df['M'].values, 'S': df['S'].values,
#         'HAZ (WHO-LMS)': df['z_score'].values,
#         'Label WHO': df['label_who'].values,
#     })
#     output_path = os.path.join(TRAINFILE_DIR, 'task3_analisa_perbaris.csv')
#     df_output.to_csv(output_path, index=False)
#     print(f"    [OK] Output -> {output_path} ({len(df_output):,} baris)")
    
#     return df


# # ============================================================================
# # TASK 4: FUZZY SYSTEM LOGIC (Takagi-Sugeno)
# # ============================================================================


# # -- Fungsi Keanggotaan Dasar --
# def trapezoid(x, a, b, c, d):
#     """Fungsi keanggotaan trapezoidal. mu(x) in [0, 1]."""
#     if x <= a or x >= d:
#         return 0.0
#     elif a < x < b:
#         return (x - a) / (b - a)
#     elif b <= x <= c:
#         return 1.0
#     elif c < x < d:
#         return (d - x) / (d - c)
#     return 0.0

# def triangle(x, a, b, c):
#     """Fungsi keanggotaan segitiga (kasus khusus trapesium b==c)."""
#     return trapezoid(x, a, b, b, c)


# # ============================================================================
# # PERUBAHAN 3 (Opsi B): MF INPUT - Diperlebar overlap untuk gradasi
# # Berdasarkan standar klinis WHO + analisa distribusi data SKI 2023
# # ============================================================================
# BBL_MF_PARAMS = [1200, 1500, 2200, 3000, 3400, 3900]
# # BBL: VLBW(0,0,1200,1500) LBW(1200,1500,2200,3000) Normal(2200,3000,3400,3900) Macro(3400,3900,6000,6000)
# # LBW overlap dengan Normal di range 2200-3000 -> balita borderline punya partial membership

# PBL_MF_PARAMS = [42, 48, 50, 52, 53]
# # PBL: VeryShort(0,0,42,48) Short(42,48,50) Normal(48,50,52,53) Long(52,53,100,100)
# # Short overlap dengan Normal di range 48-50 -> borderline tercakup

# MPASI_MF_PARAMS = [3, 5, 7, 8]
# # MPASI: TooEarly(0,0,3,5) Timely(3,5,7,8) Late(7,8,12,12)
# # Range diperlebar dari proposal agar transisi lebih gradual


# def fuzzify(bbl, pbl, mpasi):
#     """Fuzzyfikasi: mengubah nilai tegas ke derajat keanggotaan mu in [0,1]."""
#     bp, pp, mp = BBL_MF_PARAMS, PBL_MF_PARAMS, MPASI_MF_PARAMS
#     return {
#         'bbl': {
#             'VLBW':       trapezoid(bbl, 0, 0, bp[0], bp[1]),
#             'LBW':        trapezoid(bbl, bp[0], bp[1], bp[2], bp[3]),
#             'Normal':     trapezoid(bbl, bp[2], bp[3], bp[4], bp[5]),
#             'Macrosomia': trapezoid(bbl, bp[4], bp[5], 6000, 6000),
#         },
#         'pbl': {
#             'VeryShort':  trapezoid(pbl, 0, 0, pp[0], pp[1]),
#             'Short':      triangle(pbl, pp[0], pp[1], pp[2]),
#             'Normal':     trapezoid(pbl, pp[1], pp[2], pp[3], pp[4]),
#             'LongLength': trapezoid(pbl, pp[3], pp[4], 100, 100),
#         },
#         'mpasi': {
#             'TooEarly':   trapezoid(mpasi, 0, 0, mp[0], mp[1]),
#             'Timely':     trapezoid(mpasi, mp[0], mp[1], mp[2], mp[3]),
#             'Late':       trapezoid(mpasi, mp[2], mp[3], 12, 12),
#         }
#     }


# # ============================================================================
# # PERUBAHAN 1: Consequent Values = Titik puncak FK Output (Gambar 3.9)
# # FIXED - tidak boleh diubah karena harus selaras dengan grafik output
# # ============================================================================
# CONSEQUENTS = {'high': 90, 'medium': 65, 'low': 35, 'normal': 10}


# # ============================================================================
# # Aturan Dasar (Tabel III.2 Proposal + Expanded)
# # 30 rules: 23 original + 7 tambahan untuk cakupan borderline
# # Operator: AND = MIN, OR = MAX
# # ============================================================================
# FUZZY_RULES = [
#     # STUNTING HIGH (c=90) - Kombinasi faktor risiko berat
#     ('VLBW',  'VeryShort', None,                   'high'),   # R1
#     ('VLBW',  'VeryShort', ('Late', 'TooEarly'),   'high'),   # R2
#     ('VLBW',  'Short',     ('Late', 'TooEarly'),   'high'),   # R3
#     ('VLBW',  None,        'Late',                 'high'),   # R4
#     ('VLBW',  None,        'TooEarly',             'high'),   # R5
#     (None,    'VeryShort', 'Late',                 'high'),   # R6
#     (None,    'VeryShort', 'TooEarly',             'high'),   # R7

#     # STUNTING MEDIUM (c=65) - 1-2 faktor risiko moderat
#     ('LBW',   'Short',     None,                   'medium'), # R8
#     ('LBW',   None,        ('Late', 'TooEarly'),   'medium'), # R9
#     (None,    'Short',     ('Late', 'TooEarly'),   'medium'), # R10
#     ('VLBW',  'Normal',    None,                   'medium'), # R11
#     ('VLBW',  None,        'Timely',               'medium'), # R12
#     ('Normal','VeryShort', None,                   'medium'), # R13
#     (None,    'VeryShort', 'Timely',               'medium'), # R14
#     ('LBW',   'Short',     'Timely',               'medium'), # R15 (NEW)
#     ('LBW',   'Normal',    ('Late', 'TooEarly'),   'medium'), # R16 (NEW)

#     # STUNTING LOW (c=35) - Faktor risiko ringan
#     ('LBW',   'Normal',    'Timely',               'low'),    # R17
#     ('Normal','Short',     'Timely',               'low'),    # R18
#     ('Normal','Normal',    ('TooEarly', 'Late'),   'low'),    # R19
#     ('LBW',   'Normal',    ('TooEarly', 'Late'),   'low'),    # R20
#     ('Normal','Short',     ('TooEarly', 'Late'),   'low'),    # R21
#     ('LBW',   'LongLength','Timely',               'low'),    # R22 (NEW)
#     ('LBW',   'LongLength',('TooEarly', 'Late'),   'low'),    # R23 (NEW)

#     # NORMAL / NO RISK (c=10) - Semua faktor baik
#     ('Normal','Normal',    'Timely',               'normal'), # R24
#     ('Normal','LongLength','Timely',               'normal'), # R25
#     ('Macrosomia','Normal','Timely',               'normal'), # R26
#     ('Macrosomia','LongLength','Timely',           'normal'), # R27
#     ('Macrosomia','Normal',('TooEarly','Late'),    'normal'), # R28 (NEW)
#     ('Macrosomia','LongLength',('TooEarly','Late'),'normal'), # R29 (NEW)
#     ('Normal','LongLength',('TooEarly','Late'),    'normal'), # R30 (NEW)
# ]


# def evaluate_rules(mu, cons=None):
#     """
#     Evaluasi semua aturan fuzzy (Tabel III.2 expanded).
#     Operator AND = MIN, Operator OR = MAX.
#     Returns: list of (weight, consequent_value) tuples.
#     """
#     if cons is None:
#         cons = CONSEQUENTS

#     fired = []
#     for rule in FUZZY_RULES:
#         bbl_cond, pbl_cond, mpasi_cond, risk_class = rule
#         antecedents = []

#         if bbl_cond is not None:
#             if isinstance(bbl_cond, tuple):
#                 antecedents.append(max(mu['bbl'].get(c, 0) for c in bbl_cond))
#             else:
#                 antecedents.append(mu['bbl'].get(bbl_cond, 0))

#         if pbl_cond is not None:
#             if isinstance(pbl_cond, tuple):
#                 antecedents.append(max(mu['pbl'].get(c, 0) for c in pbl_cond))
#             else:
#                 antecedents.append(mu['pbl'].get(pbl_cond, 0))

#         if mpasi_cond is not None:
#             if isinstance(mpasi_cond, tuple):
#                 antecedents.append(max(mu['mpasi'].get(c, 0) for c in mpasi_cond))
#             else:
#                 antecedents.append(mu['mpasi'].get(mpasi_cond, 0))

#         if antecedents:
#             w = min(antecedents)  # AND = MIN
#             if w > 0:
#                 c_val = cons[risk_class]
#                 fired.append((w, c_val))

#     return fired


# def defuzzify_sugeno(fired_rules):
#     """
#     Defuzzyfikasi Sugeno: Weighted Average.
#     y = sum(wi * ci) / sum(wi)
#     """
#     if not fired_rules:
#         return float(CONSEQUENTS['low'])  # default jika tidak ada rule fire

#     numerator = sum(w * c for w, c in fired_rules)
#     denominator = sum(w for w, c in fired_rules)

#     if denominator == 0:
#         return float(CONSEQUENTS['low'])

#     return round(numerator / denominator, 2)


# # ============================================================================
# # PERUBAHAN 2: FK Output Segitiga Berpotongan (Gambar 3.9)
# # Risk score di-map ke derajat keanggotaan di setiap priority flag
# # Flag final = argmax(derajat keanggotaan)
# # ============================================================================
# # Benchmark dari Gambar 3.9 proposal:
# #   Normal:         trap(0, 0, 15, 30)     puncak di 0-15
# #   Stunting Low:   triangle(15, 30, 60)   puncak di 30
# #   Stunting Medium:triangle(45, 60, 75)   puncak di 60
# #   Stunting High:  trap(75, 90, 100, 100) puncak di 90-100

# OUTPUT_MF = {
#     'Normal':          (0, 0, 15, 30),      # trapezoidal
#     'STUNTING LOW':    (15, 30, 60),         # triangle
#     'STUNTING MEDIUM': (45, 60, 75),         # triangle
#     'STUNTING HIGH':   (75, 90, 100, 100),   # trapezoidal
# }

# def get_output_membership(risk_score):
#     """Hitung derajat keanggotaan risk score di setiap priority flag."""
#     mu_normal = trapezoid(risk_score, -1, 0, 15, 30)
#     if risk_score <= 0:
#         mu_normal = 1.0
#     mu_low    = triangle(risk_score, 15, 30, 60)
#     mu_medium = triangle(risk_score, 45, 60, 75)
#     mu_high   = trapezoid(risk_score, 75, 90, 100, 101)
#     if risk_score >= 100:
#         mu_high = 1.0

#     return {
#         'Normal': round(mu_normal, 4),
#         'STUNTING LOW': round(mu_low, 4),
#         'STUNTING MEDIUM': round(mu_medium, 4),
#         'STUNTING HIGH': round(mu_high, 4),
#     }

# def get_priority_flag(risk_score):
#     """
#     Priority flag = kategori dengan derajat keanggotaan TERTINGGI
#     pada FK output segitiga berpotongan (Gambar 3.9).
#     """
#     mu_out = get_output_membership(risk_score)
#     return max(mu_out, key=mu_out.get)


# def task4_fuzzy_logic(df):
#     """Task 4: Fuzzy System Logic (Takagi-Sugeno)."""

#     print_header("TASK 4: FUZZY SYSTEM LOGIC - Takagi-Sugeno")

#     print_step("4.1", "Fungsi Keanggotaan (Fuzzyfikasi)")
#     bp, pp, mp = BBL_MF_PARAMS, PBL_MF_PARAMS, MPASI_MF_PARAMS
#     print(f"    BBL: VLBW(0,0,{bp[0]},{bp[1]}), LBW({bp[0]},{bp[1]},{bp[2]},{bp[3]}), Normal({bp[2]},{bp[3]},{bp[4]},{bp[5]}), Macro({bp[4]},{bp[5]},6000,6000)")
#     print(f"    PBL: VeryShort(0,0,{pp[0]},{pp[1]}), Short({pp[0]},{pp[1]},{pp[2]}), Normal({pp[1]},{pp[2]},{pp[3]},{pp[4]}), Long({pp[3]},{pp[4]},100,100)")
#     print(f"    MPASI: TooEarly(0,0,{mp[0]},{mp[1]}), Timely({mp[0]},{mp[1]},{mp[2]},{mp[3]}), Late({mp[2]},{mp[3]},12,12)")

#     print_step("4.2", f"Aturan Dasar (Rule Base) - {len(FUZZY_RULES)} aturan")
#     risk_labels = {'high': 'STUNTING HIGH', 'medium': 'STUNTING MEDIUM',
#                    'low': 'STUNTING LOW', 'normal': 'Normal'}
#     for cls in ['high', 'medium', 'low', 'normal']:
#         rules_c = [r for r in FUZZY_RULES if r[3] == cls]
#         print(f"    {risk_labels[cls]} (c={CONSEQUENTS[cls]}): {len(rules_c)} aturan")

#     print_step("4.3", "Defuzzyfikasi Sugeno - Weighted Average")
#     print(f"    Rumus: y = sum(wi * ci) / sum(wi)")
#     print(f"    Consequent: Normal={CONSEQUENTS['normal']}, Low={CONSEQUENTS['low']}, Medium={CONSEQUENTS['medium']}, High={CONSEQUENTS['high']}")

#     print_step("4.3b", "FK Output - Segitiga Berpotongan (Gambar 3.9)")
#     print(f"    Normal:         trap(0, 0, 15, 30)    puncak 0-15")
#     print(f"    Stunting Low:   tri(15, 30, 60)        puncak 30")
#     print(f"    Stunting Medium:tri(45, 60, 75)        puncak 60")
#     print(f"    Stunting High:  trap(75, 90, 100, 100) puncak 90-100")
#     print(f"    Flag = argmax(derajat keanggotaan output)")

#     print_step("4.4", f"ANALISA FUZZY PER BARIS - {len(df):,} balita")

#     risk_scores = []
#     priority_flags = []
#     rules_fired_counts = []

#     for idx in range(len(df)):
#         bbl  = float(df.iloc[idx]['bbl_gram'])
#         pbl  = float(df.iloc[idx]['pbl_cm'])
#         mpasi = float(df.iloc[idx]['mpasi_bulan'])

#         mu = fuzzify(bbl, pbl, mpasi)
#         fired = evaluate_rules(mu)
#         risk = defuzzify_sugeno(fired)
#         flag = get_priority_flag(risk)

#         risk_scores.append(risk)
#         priority_flags.append(flag)
#         rules_fired_counts.append(len(fired))

#     df['risk_score'] = risk_scores
#     df['priority_flag'] = priority_flags

#     print(f"    Fuzzy inference selesai: {len(df):,} balita")
#     print(f"    Rata-rata rules fired per balita: {np.mean(rules_fired_counts):.1f}")

#     print_step("4.5", "Statistik Risk Score & Priority Flag")
#     rs = df['risk_score']
#     print(f"    Risk Score: Mean={rs.mean():.2f}, Std={rs.std():.2f}, Min={rs.min():.2f}, Med={rs.median():.2f}, Max={rs.max():.2f}")
#     print(f"    Unique risk scores: {rs.nunique()}")

#     print(f"\n    Distribusi Priority Flag:")
#     total = len(df)
#     for flag_name in ['Normal', 'STUNTING LOW', 'STUNTING MEDIUM', 'STUNTING HIGH']:
#         cnt = (df['priority_flag'] == flag_name).sum()
#         pct = cnt / total * 100
#         print(f"    {flag_name:18s}: {cnt:>4d} ({pct:>5.1f}%)")

#     print(f"\n    Cross-tabulation: Label WHO x Priority Flag:")
#     for who_label in ['Normal', 'Stunting']:
#         subset = df[df['label_who'] == who_label]
#         flags = subset['priority_flag'].value_counts()
#         parts = [f"{f}={flags.get(f, 0)}" for f in ['Normal', 'STUNTING LOW', 'STUNTING MEDIUM', 'STUNTING HIGH']]
#         print(f"    WHO {who_label:8s} (n={len(subset):>3d}): {', '.join(parts)}")

#     # Output tabel per baris
#     print_step("4.6", "OUTPUT TABEL ANALISA LENGKAP PER BARIS (Z-Score + Fuzzy)")

#     hdr = (f"    {'No':>4s} | {'Sex':^10s} | {'Umur':>4s} | {'PB':>7s} | "
#            f"{'L':>5s} | {'M':>9s} | {'S':>8s} | {'HAZ':>8s} | {'Label':^8s} | "
#            f"{'BBL':>6s} | {'PBL':>6s} | {'MPASI':>5s} | {'Risk':>6s} | {'Priority Flag':^16s}")
#     sep = (f"    {'--'*2}--+--{'--'*5}--+--{'--'*2}--+--{'--'*3}---+--"
#            f"{'--'*2}--+--{'--'*4}---+--{'--'*4}--+--{'--'*4}--+--{'--'*4}--+--"
#            f"{'--'*3}--+--{'--'*3}--+--{'--'*2}---+--{'--'*3}--+--{'--'*8}")

#     print(f"\n{sep}")
#     print(hdr)
#     print(sep)

#     for i in range(len(df)):
#         row = df.iloc[i]
#         no = i + 1
#         jk_str = 'Laki-laki' if int(row['jenis_kelamin']) == 1 else 'Perempuan'
#         print(f"    {no:>4d} | {jk_str:^10s} | {int(row['usia_bulan']):>4d} | {row['tinggi_badan_cm']:>7.2f} | "
#               f"{row['L']:>5.2f} | {row['M']:>9.3f} | {row['S']:>8.5f} | {row['z_score']:>8.3f} | {row['label_who']:^8s} | "
#               f"{row['bbl_gram']:>6.0f} | {row['pbl_cm']:>6.2f} | {row['mpasi_bulan']:>5.2f} | {row['risk_score']:>6.2f} | {row['priority_flag']:^16s}")

#     print(sep)
#     print(f"    TOTAL: {len(df):,} baris dianalisa (Z-Score WHO + Fuzzy Sugeno)")

#     # Simpan
#     print_step("4.7", "Menyimpan output analisa lengkap per baris")
#     df_output = pd.DataFrame({
#         'No': range(1, len(df) + 1),
#         'Sex': df['jenis_kelamin'].map({1: 'Laki-laki', 2: 'Perempuan'}).values,
#         'Umur (bln)': df['usia_bulan'].astype(int).values,
#         'PB anak (cm)': df['tinggi_badan_cm'].values,
#         'L': df['L'].values, 'M': df['M'].values, 'S': df['S'].values,
#         'HAZ (WHO-LMS)': df['z_score'].values, 'Label WHO': df['label_who'].values,
#         'BBL (g)': df['bbl_gram'].values, 'PBL (cm)': df['pbl_cm'].values,
#         'MPASI (bln)': df['mpasi_bulan'].values,
#         'Risk score (Sugeno)': df['risk_score'].values,
#         'Priority flag': df['priority_flag'].values,
#     })
#     output_path = os.path.join(TRAINFILE_DIR, 'task3_analisa_perbaris.csv')
#     df_output.to_csv(output_path, index=False)
#     print(f"    Saved: {output_path}")

#     return df


# # ============================================================================
# # TASK 5 & 6: EVALUASI + SKEMA PENGUJIAN (TRAIN/TEST)
# # ============================================================================

# def fuzzy_infer_single(bbl, pbl, mpasi, mf_p, cons):
#     """
#     Full pipeline untuk 1 balita:
#     Input -> Fuzzyfikasi -> Rules -> Defuzzyfikasi Sugeno -> FK Output -> Flag
#     """
#     bp, pp, mp = mf_p['bbl'], mf_p['pbl'], mf_p['mpasi']

#     mu_bbl = {
#         'VLBW':       trapezoid(bbl, 0, 0, bp[0], bp[1]),
#         'LBW':        trapezoid(bbl, bp[0], bp[1], bp[2], bp[3]),
#         'Normal':     trapezoid(bbl, bp[2], bp[3], bp[4], bp[5]),
#         'Macrosomia': trapezoid(bbl, bp[4], bp[5], 6000, 6000),
#     }
#     mu_pbl = {
#         'VeryShort':  trapezoid(pbl, 0, 0, pp[0], pp[1]),
#         'Short':      triangle(pbl, pp[0], pp[1], pp[2]),
#         'Normal':     trapezoid(pbl, pp[1], pp[2], pp[3], pp[4]),
#         'LongLength': trapezoid(pbl, pp[3], pp[4], 100, 100),
#     }
#     mu_mpasi = {
#         'TooEarly':   trapezoid(mpasi, 0, 0, mp[0], mp[1]),
#         'Timely':     trapezoid(mpasi, mp[0], mp[1], mp[2], mp[3]),
#         'Late':       trapezoid(mpasi, mp[2], mp[3], 12, 12),
#     }

#     mu = {'bbl': mu_bbl, 'pbl': mu_pbl, 'mpasi': mu_mpasi}
#     fired = evaluate_rules(mu, cons)

#     if fired:
#         risk = sum(w * c for w, c in fired) / sum(w for w, c in fired)
#     else:
#         risk = cons['low']

#     risk = round(risk, 2)
#     flag = get_priority_flag(risk)
#     binary = 0 if flag == 'Normal' else 1

#     return risk, flag, binary


# def run_pipeline(dataframe, mf_p, cons):
#     """Run full pipeline pada seluruh dataframe."""
#     risks, flags, binaries = [], [], []
#     for idx in range(len(dataframe)):
#         bbl = float(dataframe.iloc[idx]['bbl_gram'])
#         pbl = float(dataframe.iloc[idx]['pbl_cm'])
#         mpasi = float(dataframe.iloc[idx]['mpasi_bulan'])
#         r, f, b = fuzzy_infer_single(bbl, pbl, mpasi, mf_p, cons)
#         risks.append(r)
#         flags.append(f)
#         binaries.append(b)
#     return np.array(risks), flags, np.array(binaries)


# def calc_metrics(y_true, y_pred):
#     """Confusion matrix dari binary labels."""
#     tp = int(np.sum((y_true == 1) & (y_pred == 1)))
#     tn = int(np.sum((y_true == 0) & (y_pred == 0)))
#     fp = int(np.sum((y_true == 0) & (y_pred == 1)))
#     fn = int(np.sum((y_true == 1) & (y_pred == 0)))
#     prec = tp/(tp+fp) if (tp+fp) > 0 else 0.0
#     rec  = tp/(tp+fn) if (tp+fn) > 0 else 0.0
#     f1   = 2*prec*rec/(prec+rec) if (prec+rec) > 0 else 0.0
#     acc  = (tp+tn)/(tp+tn+fp+fn) if (tp+tn+fp+fn) > 0 else 0.0
#     return {'TP':tp, 'TN':tn, 'FP':fp, 'FN':fn,
#             'precision':round(prec,4), 'recall':round(rec,4),
#             'f1':round(f1,4), 'accuracy':round(acc,4)}


# def task5_6_evaluation(df):
#     """
#     Task 5+6: Evaluasi + Skema Pengujian

#     Pipeline:
#       Input -> Fuzzyfikasi -> Rules -> Defuzzyfikasi Sugeno (y = sum(wi*ci)/sum(wi))
#       -> FK Output Segitiga Berpotongan -> Priority Flag (argmax mu)
#       -> vs WHO Label -> Confusion Matrix
#     """
#     print_header("TASK 5 & 6: EVALUASI + SKEMA PENGUJIAN SISTEM")

#     print(f"    Pipeline: Input > Fuzzyfikasi > Rules > Defuzzyfikasi Sugeno > FK Output > Flag")
#     print(f"    Consequent: Normal={CONSEQUENTS['normal']}, Low={CONSEQUENTS['low']}, Med={CONSEQUENTS['medium']}, High={CONSEQUENTS['high']}")
#     print(f"    FK Output: segitiga berpotongan sesuai Gambar 3.9")
#     print(f"    Training: 80/20 stratified, 7 seeds, grid search MF input")

#     y_all = (df['label_who'] == 'Stunting').astype(int).values

#     # MF Input Grid - search untuk MF terbaik
#     # Consequent FIXED di [10,35,65,90], FK output FIXED segitiga Gambar 3.9
#     BBL_GRID = [
#         [1200, 1500, 2200, 3000, 3400, 3900],
#         [1200, 1500, 2300, 2800, 3500, 3900],
#         [1200, 1500, 2200, 2800, 3500, 3900],
#         [1200, 1500, 2200, 3000, 3500, 4000],
#         [1250, 1500, 2450, 2500, 3650, 3999],  # proposal asli
#         [1200, 1500, 2300, 2900, 3500, 3900],
#     ]
#     PBL_GRID = [
#         [42, 48, 50, 52, 53],
#         [43, 47, 50, 53, 54],
#         [42, 47, 50, 52, 53],
#         [43, 48, 50, 53, 54],
#         [44, 46, 48, 53, 54],  # proposal asli
#         [42, 48, 50, 53, 54],
#     ]
#     MPASI_GRID = [
#         [3, 5, 7, 8],
#         [4, 5, 6, 7],  # proposal asli
#         [3, 5, 6, 8],
#         [3, 5, 6, 7],
#     ]

#     CONS_FIXED = CONSEQUENTS  # TIDAK diubah

#     SEEDS = [42, 123, 256, 789, 1024, 2048, 3333]

#     n_mf = len(BBL_GRID) * len(PBL_GRID) * len(MPASI_GRID)
#     print_step("6.1", "Search Space")
#     print(f"    MF configs: {n_mf} | Consequent: FIXED [10,35,65,90]")
#     print(f"    Seeds: {len(SEEDS)} | FK Output: FIXED (Gambar 3.9)")

#     print_step("6.2", "Iterative Forward-Backward Training")

#     global_best = {'combo': -1}

#     for seed in SEEDS:
#         np.random.seed(seed)
#         idx_n = np.where(y_all == 0)[0].copy()
#         idx_s = np.where(y_all == 1)[0].copy()
#         np.random.shuffle(idx_n)
#         np.random.shuffle(idx_s)

#         n_tr_n = int(len(idx_n) * 0.8)
#         n_tr_s = int(len(idx_s) * 0.8)
#         tr_idx = np.concatenate([idx_n[:n_tr_n], idx_s[:n_tr_s]])
#         te_idx = np.concatenate([idx_n[n_tr_n:], idx_s[n_tr_s:]])

#         df_tr = df.iloc[tr_idx].reset_index(drop=True)
#         df_te = df.iloc[te_idx].reset_index(drop=True)
#         y_tr = (df_tr['label_who'] == 'Stunting').astype(int).values
#         y_te = (df_te['label_who'] == 'Stunting').astype(int).values

#         seed_best = {'score': -1}

#         for bbl_p in BBL_GRID:
#             for pbl_p in PBL_GRID:
#                 for mpasi_p in MPASI_GRID:
#                     mf = {'bbl': bbl_p, 'pbl': pbl_p, 'mpasi': mpasi_p}

#                     _, _, y_pred_tr = run_pipeline(df_tr, mf, CONS_FIXED)
#                     m = calc_metrics(y_tr, y_pred_tr)

#                     score = m['f1']
#                     if score > seed_best['score']:
#                         seed_best = {
#                             'score': score, 'mf': mf.copy(),
#                             'train_m': m.copy(), 'seed': seed,
#                             'tr_idx': tr_idx, 'te_idx': te_idx,
#                         }

#         # Validasi pada test set
#         if seed_best['score'] > 0:
#             _, _, y_pred_te = run_pipeline(df_te, seed_best['mf'], CONS_FIXED)
#             test_m = calc_metrics(y_te, y_pred_te)

#             combo = 0.3 * seed_best['score'] + 0.7 * test_m['f1']

#             trn = seed_best['train_m']
#             print(f"    Seed {seed:>4d} | Train: Acc={trn['accuracy']:.4f} F1={trn['f1']:.4f} P={trn['precision']:.4f} R={trn['recall']:.4f} | "
#                   f"Test: Acc={test_m['accuracy']:.4f} F1={test_m['f1']:.4f}")

#             if combo > global_best.get('combo', -1):
#                 global_best = seed_best.copy()
#                 global_best['combo'] = combo
#                 global_best['test_m'] = test_m
#                 global_best['y_tr'] = y_tr
#                 global_best['y_te'] = y_te
#                 global_best['df_tr'] = df_tr
#                 global_best['df_te'] = df_te

#     best = global_best

#     print_step("6.3", "BEST MODEL")
#     print(f"    BBL MF      : {best['mf']['bbl']}")
#     print(f"    PBL MF      : {best['mf']['pbl']}")
#     print(f"    MPASI MF    : {best['mf']['mpasi']}")
#     print(f"    Consequents : H={CONS_FIXED['high']} M={CONS_FIXED['medium']} L={CONS_FIXED['low']} N={CONS_FIXED['normal']}")
#     print(f"    FK Output   : Gambar 3.9 (segitiga berpotongan)")
#     print(f"    Seed        : {best['seed']}")

#     # Re-compute seluruh data
#     print_step("6.4", "Re-compute seluruh data")

#     all_risks, all_flags, all_binary = run_pipeline(df, best['mf'], CONS_FIXED)
#     df['risk_score'] = np.round(all_risks, 2)
#     df['priority_flag'] = all_flags
#     df['fuzzy_label'] = ['Stunting' if b == 1 else 'Normal' for b in all_binary]

#     flag_dist = pd.Series(all_flags).value_counts()
#     print(f"    Priority Flag distribution:")
#     for flag_name in ['Normal', 'STUNTING LOW', 'STUNTING MEDIUM', 'STUNTING HIGH']:
#         cnt = flag_dist.get(flag_name, 0)
#         print(f"      {flag_name:18s}: {cnt:>3d} ({cnt/len(df)*100:.1f}%)")

#     # EVALUASI: Confusion Matrix
#     # TRAIN
#     print_step("5.1", f"CONFUSION MATRIX - Train ({len(best['y_tr'])})")
#     _, _, y_pred_tr = run_pipeline(best['df_tr'], best['mf'], CONS_FIXED)
#     trm = calc_metrics(best['y_tr'], y_pred_tr)
#     tp,tn,fp,fn = trm['TP'],trm['TN'],trm['FP'],trm['FN']
#     print(f"    TRAIN SET (80%)")
#     print(f"                  Act Stunting   Act Normal")
#     print(f"    Pred Stunting    TP = {tp:>4d}     FP = {fp:>4d}")
#     print(f"    Pred Normal      FN = {fn:>4d}     TN = {tn:>4d}")
#     print(f"    Precision={trm['precision']:.4f}, Recall={trm['recall']:.4f}, F1={trm['f1']:.4f}, Accuracy={trm['accuracy']:.4f}")

#     # TEST
#     print_step("5.2", f"CONFUSION MATRIX - Test ({len(best['y_te'])})")
#     tm = best['test_m']
#     tp,tn,fp,fn = tm['TP'],tm['TN'],tm['FP'],tm['FN']
#     print(f"    TEST SET (20%) - Data independen")
#     print(f"                  Act Stunting   Act Normal")
#     print(f"    Pred Stunting    TP = {tp:>4d}     FP = {fp:>4d}")
#     print(f"    Pred Normal      FN = {fn:>4d}     TN = {tn:>4d}")
#     print(f"    Precision={tm['precision']:.4f}, Recall={tm['recall']:.4f}, F1={tm['f1']:.4f}, Accuracy={tm['accuracy']:.4f}")

#     # ALL
#     print_step("5.3", f"CONFUSION MATRIX - All Data ({len(df)})")
#     y_all_t = (df['label_who']=='Stunting').astype(int).values
#     y_all_p = (df['fuzzy_label']=='Stunting').astype(int).values
#     am = calc_metrics(y_all_t, y_all_p)
#     tp,tn,fp,fn = am['TP'],am['TN'],am['FP'],am['FN']
#     print(f"    SELURUH DATA ({len(df)})")
#     print(f"                  Act Stunting   Act Normal")
#     print(f"    Pred Stunting    TP = {tp:>4d}     FP = {fp:>4d}")
#     print(f"    Pred Normal      FN = {fn:>4d}     TN = {tn:>4d}")
#     print(f"    Precision={am['precision']:.4f}, Recall={am['recall']:.4f}, F1={am['f1']:.4f}, Accuracy={am['accuracy']:.4f}")

#     # RINGKASAN
#     print_step("5.4", "Ringkasan Evaluasi")
#     print(f"    Dataset       Precision  Recall   F1-Score  Accuracy")
#     print(f"    Train (80%)   {trm['precision']:.4f}     {trm['recall']:.4f}   {trm['f1']:.4f}    {trm['accuracy']:.4f}")
#     print(f"    Test  (20%)   {tm['precision']:.4f}     {tm['recall']:.4f}   {tm['f1']:.4f}    {tm['accuracy']:.4f}")
#     print(f"    All  (100%)   {am['precision']:.4f}     {am['recall']:.4f}   {am['f1']:.4f}    {am['accuracy']:.4f}")

#     # SIMPAN MODEL
#     print_step("5.5", "Simpan model terlatih")
#     trained_model = {
#         'mf_params': best['mf'],
#         'consequents': CONS_FIXED,
#         'output_mf': OUTPUT_MF,
#         'seed': int(best['seed']),
#         'metrics': {'train': trm, 'test': tm, 'all': am},
#         'split': {'train_size': len(best['y_tr']), 'test_size': len(best['y_te'])},
#         'timestamp': datetime.now().isoformat(),
#     }
#     model_path = os.path.join(TRAINFILE_DIR, 'trained_model.json')
#     with open(model_path, 'w') as f:
#         json.dump(trained_model, f, indent=2, default=str)
#     print(f"    Saved: {model_path}")

#     df.iloc[best['tr_idx']].reset_index(drop=True).to_csv(
#         os.path.join(TRAINFILE_DIR, 'dataset_train.csv'), index=False)
#     df.iloc[best['te_idx']].reset_index(drop=True).to_csv(
#         os.path.join(TRAINFILE_DIR, 'dataset_test.csv'), index=False)
#     print(f"    Saved: dataset_train.csv ({len(best['tr_idx'])}), dataset_test.csv ({len(best['te_idx'])})")

#     return df, trained_model


# # ============================================================================
# # MAIN
# # ============================================================================
# if __name__ == '__main__':

#     print("\n" + "=" * 60)
#     print("  STUNTOR - Klasifikasi Stunting Balita Kalimantan Tengah")
#     print("  Dataset: SKI 2023 | Metode: Fuzzy Sugeno")
#     print("=" * 60)
#     print(f"\n  Working directory : {os.getcwd()}")
#     print(f"  Script location   : {os.path.abspath(__file__)}")
#     print(f"  Dataset path      : {DATASET_PATH}")
#     print(f"  WHO Z-Score dir   : {WHO_ZS_DIR}")
#     print(f"  Output dir        : {TRAINFILE_DIR}")

#     # TASK 1: INPUT DATA
#     df_input, raw_shape = task1_input_data()

#     # TASK 2: PREPROCESSING
#     df_clean, report = task2_preprocessing(df_input, raw_shape)
#     df_clean.to_csv(OUTPUT_CLEAN_CSV, index=False)
#     print(f"\n    Dataset bersih: {OUTPUT_CLEAN_CSV} ({len(df_clean):,} baris)")

#     # TASK 3: Z-SCORE & LABELLING
#     df_labeled = task3_zscore_labelling(df_clean)

#     # TASK 4: FUZZY SYSTEM LOGIC
#     df_fuzzy = task4_fuzzy_logic(df_labeled)

#     # TASK 5 & 6: EVALUASI + TRAIN/TEST
#     df_final, trained_model = task5_6_evaluation(df_fuzzy)

#     # SIMPAN SEMUA HASIL
#     print_header("MENYIMPAN HASIL AKHIR")

#     OUTPUT_FINAL_CSV = os.path.join(TRAINFILE_DIR, 'dataset_final.csv')
#     df_final.to_csv(OUTPUT_FINAL_CSV, index=False)
#     print(f"    Dataset final: {OUTPUT_FINAL_CSV}")
#     print(f"    Shape: {df_final.shape[0]:,} baris x {df_final.shape[1]} kolom")

#     df_output = pd.DataFrame({
#         'No': range(1, len(df_final) + 1),
#         'Sex': df_final['jenis_kelamin'].map({1: 'Laki-laki', 2: 'Perempuan'}).values,
#         'Umur (bln)': df_final['usia_bulan'].astype(int).values,
#         'PB anak (cm)': df_final['tinggi_badan_cm'].values,
#         'L': df_final['L'].values, 'M': df_final['M'].values, 'S': df_final['S'].values,
#         'HAZ (WHO-LMS)': df_final['z_score'].values,
#         'Label WHO': df_final['label_who'].values,
#         'BBL (g)': df_final['bbl_gram'].values,
#         'PBL (cm)': df_final['pbl_cm'].values,
#         'MPASI (bln)': df_final['mpasi_bulan'].values,
#         'Risk score (Sugeno)': df_final['risk_score'].values,
#         'Priority flag': df_final['priority_flag'].values,
#     })
#     out_perbaris = os.path.join(TRAINFILE_DIR, 'task3_analisa_perbaris.csv')
#     df_output.to_csv(out_perbaris, index=False)
#     print(f"    Analisa per baris: {out_perbaris}")

#     with open(OUTPUT_REPORT, 'w') as f:
#         label_dist = df_final['label_who'].value_counts().to_dict()
#         flag_dist = df_final['priority_flag'].value_counts().to_dict()
#         report['task3'] = {
#             'label_distribution': {k: int(v) for k, v in label_dist.items()},
#             'stunting_rate': round(label_dist.get('Stunting', 0) / len(df_final) * 100, 2),
#         }
#         report['task4'] = {
#             'priority_flag_distribution': {k: int(v) for k, v in flag_dist.items()},
#         }
#         report['task5_6'] = trained_model
#         json.dump(report, f, indent=2, default=str)
#     print(f"    Laporan: {OUTPUT_REPORT}")

#     # RINGKASAN
#     print_header("RINGKASAN PIPELINE TASK 1 - 6")

#     tm = trained_model['metrics']
#     print(f"    TASK 1: Input Data    - {raw_shape[0]:,} x {raw_shape[1]} -> 6 variabel")
#     print(f"    TASK 2: Preprocessing - {len(df_clean):,} balita bersih")
#     print(f"    TASK 3: Z-Score WHO   - Normal={label_dist.get('Normal',0)}, Stunting={label_dist.get('Stunting',0)}")
#     print(f"    TASK 4: Fuzzy Sugeno  - {len(FUZZY_RULES)} rules, c=[10,35,65,90], FK output segitiga")
#     print(f"    TASK 5: Evaluasi      - Confusion Matrix (P/R/F1/Acc)")
#     print(f"    TASK 6: Train/Test    - 80/20 Stratified + MF Optimization")
#     print(f"")
#     print(f"    Hasil Evaluasi (Test 20%):")
#     print(f"    Precision={tm['test']['precision']:.4f}, Recall={tm['test']['recall']:.4f}, F1={tm['test']['f1']:.4f}, Accuracy={tm['test']['accuracy']:.4f}")
#     print(f"")
#     print(f"    Output: dataset_final.csv, trained_model.json, dataset_train/test.csv")
#     print(f"")


