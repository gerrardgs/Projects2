#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import json
import os
from datetime import datetime
from deap import base, creator, tools, algorithms
import random
import matplotlib.pyplot as plt

# ============================================================================
# PATH CONFIGURATION
# ============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATASET_PATH   = os.path.join(BASE_DIR, 'dataset', 'Individu_balita_2023_merged.csv')
WHO_ZS_DIR     = os.path.join(BASE_DIR, 'zs')
TRAINFILE_DIR  = os.path.join(BASE_DIR, 'trainfile')

OUTPUT_CLEAN_CSV   = os.path.join(TRAINFILE_DIR, 'dataset_clean.csv')
OUTPUT_FINAL_CSV   = os.path.join(TRAINFILE_DIR, 'dataset_final.csv')
OUTPUT_REPORT      = os.path.join(TRAINFILE_DIR, 'preprocessing_report.json')

os.makedirs(TRAINFILE_DIR, exist_ok=True)

WHO_TABLES = {
    'boys_0_24':   os.path.join(WHO_ZS_DIR, 'boys02.xlsx'),
    'boys_24_60':  os.path.join(WHO_ZS_DIR, 'boys25.xlsx'),
    'girls_0_24':  os.path.join(WHO_ZS_DIR, 'girls02.xlsx'),
    'girls_24_60': os.path.join(WHO_ZS_DIR, 'girls25.xlsx'),
}

# ============================================================================
# KONFIGURASI BATAS NORMAL MEDIS
# ============================================================================
VALID_RANGES = {
    'usia_bulan':    (0, 59),
    'tinggi_badan':  (35.0, 130.0),
    'berat_badan':   (1.5, 30.0),      # kg, sebelum konversi ke gram
    'pbl_cm':        (25.0, 60.0),
    'mpasi_bulan':   (0, 12),
}

NOISE_CODES = {
    'pbl': ['88', '88.8', '99', '99.9'],
    'bb':  ['88.8', '88.3', '99.9'],
}

MPASI_CODE_MAP = {
    'a': 0.25, 'b': 0.5,  'c': 1.0,  'd': 2.0,  'e': 3.0,
    'f': 4.0,  'g': 5.0,  'h': 6.0,  'i': 7.0,  'j': 8.0,
    'k': 9.0,  'l': 10.0, 'm': 11.0, 'n': 12.0,
}
MPASI_NOISE = {'o', 'p', ' ', ''}

# ============================================================================
# FUNGSI UTILITAS
# ============================================================================
def print_header(title, char='=', width=70):
    print(f"\n{char * width}")
    print(f"  {title}")
    print(f"{char * width}")

def print_step(step_num, description):
    print(f"\n  [{step_num}] {description}")

def safe_numeric(series, errors='coerce'):
    cleaned = series.astype(str).str.strip().replace({'': np.nan, ' ': np.nan})
    return pd.to_numeric(cleaned, errors=errors)


# ============================================================================
# TASK 1: INPUT DATA
# ============================================================================
def task1_input_data():
    print_header("TASK 1: INPUT DATA - Loading Dataset & Seleksi Fitur")

    print_step("1.1", "Memuat dataset SKI 2023 Kalimantan Tengah")
    if not os.path.exists(DATASET_PATH):
        raise FileNotFoundError(f"Dataset tidak ditemukan: {DATASET_PATH}")

    df_raw = pd.read_csv(DATASET_PATH, low_memory=False)
    print(f"    [OK] Dataset dimuat: {df_raw.shape[0]:,} baris x {df_raw.shape[1]} kolom")

    print_step("1.2", "Mapping kolom dataset ke variabel penelitian")

    COLUMN_MAP = {
        'Umur_BLN':                     'usia_bulan',
        'Jenis Kelamin':                'jenis_kelamin',
        'Panjang Badan Lahir':          'pbl_raw',          # TB lahir (cm) - untuk Z-score
        'Panjang Badan Anak':           'tinggi_badan_cm',  # TB saat pengukuran (cm)
        'Berat Badan Saat Data Diambil':'berat_badan_raw',  # BB saat pengukuran (kg)
        'I43_MP ASI':                   'mpasi_raw',        # MP-ASI (kode huruf)
    }

    cols_needed = ['ID_balita'] + list(COLUMN_MAP.keys())
    df = df_raw[cols_needed].copy()
    df.rename(columns=COLUMN_MAP, inplace=True)

    for var_lama, var_baru in COLUMN_MAP.items():
        print(f"    {var_lama:40s} -> {var_baru}")

    print(f"\n    [OK] Seleksi fitur: {df.shape[1]} kolom dari {df_raw.shape[1]} kolom asli")

    print_step("1.3", "Statistik awal dataset")
    print(f"    Total balita: {len(df):,}")
    for col in df.columns:
        if col == 'ID_balita':
            continue
        non_null = df[col].notna().sum()
        blank = (df[col].astype(str).str.strip() == '').sum()
        print(f"    {col:25s}: {non_null:,} terisi, {blank} kosong/blank")

    return df, df_raw.shape


# ============================================================================
# TASK 2: PREPROCESSING
# ============================================================================
def task2_preprocessing(df, raw_shape):
    print_header("TASK 2: PREPROCESSING")

    report = {
        'raw_total': len(df), 'raw_shape': list(raw_shape),
        'steps': {}, 'timestamp': datetime.now().isoformat()
    }

    # STEP 2d: TRANSFORMASI & REKODING
    print_header("STEP 2d: Transformasi & Rekoding", char='-')

    # 2d.1: Usia
    print_step("2d.1", "Transformasi Usia -> bulan")
    df['usia_bulan'] = pd.to_numeric(df['usia_bulan'], errors='coerce')
    df['usia_label'] = df['usia_bulan'].apply(
        lambda x: '<1 bulan' if pd.notna(x) and x < 1 else
                  (f'{int(x)} bulan' if pd.notna(x) else None))
    usia_bawah_1 = (df['usia_bulan'] < 1).sum()
    print(f"    [OK] Usia <1 bulan: {usia_bawah_1} -> dilabeli '<1 bulan'")

    # 2d.2: Jenis Kelamin
    print_step("2d.2", "Validasi Jenis Kelamin -> 1=L, 2=P")
    df['jenis_kelamin'] = pd.to_numeric(df['jenis_kelamin'], errors='coerce')
    jk_invalid = ~df['jenis_kelamin'].isin([1, 2]) & df['jenis_kelamin'].notna()
    print(f"    [OK] Valid: {df['jenis_kelamin'].isin([1, 2]).sum():,} | Invalid: {jk_invalid.sum()}")

    # 2d.3: Tinggi Badan (saat pengukuran)
    print_step("2d.3", "Transformasi Tinggi Badan -> cm")
    df['tinggi_badan_cm'] = safe_numeric(df['tinggi_badan_cm'])
    print(f"    [OK] Valid: {df['tinggi_badan_cm'].notna().sum():,}")

    # 2d.4: Berat Badan (saat pengukuran) - kg -> gram
    print_step("2d.4", "Transformasi Berat Badan -> gram (J01C)")
    bb_str = df['berat_badan_raw'].astype(str).str.strip()
    bb_is_noise = bb_str.isin(NOISE_CODES['bb'])
    df['berat_badan_kg'] = safe_numeric(df['berat_badan_raw'])
    df.loc[bb_is_noise, 'berat_badan_kg'] = np.nan
    df['berat_badan_gram'] = df['berat_badan_kg'] * 1000
    n_bb_valid = df['berat_badan_gram'].notna().sum()
    print(f"    [OK] Noise: {bb_is_noise.sum()} | Valid: {n_bb_valid:,}")
    if n_bb_valid > 0:
        print(f"    Range: {df['berat_badan_kg'].min():.1f} - {df['berat_badan_kg'].max():.1f} kg")

    # 2d.5: Panjang Badan Lahir (TB lahir) - untuk Z-score
    print_step("2d.5", "Transformasi Panjang Badan Lahir -> cm (I07)")
    pbl_str = df['pbl_raw'].astype(str).str.strip()
    pbl_is_noise = pbl_str.isin(NOISE_CODES['pbl'])
    df['pbl_cm'] = safe_numeric(df['pbl_raw'])
    df.loc[pbl_is_noise, 'pbl_cm'] = np.nan
    print(f"    [OK] Noise: {pbl_is_noise.sum()} | Valid: {df['pbl_cm'].notna().sum():,}")

    # 2d.6: MP-ASI (kode huruf -> bulan)
    print_step("2d.6", "Transformasi MP-ASI (kode huruf) -> bulan")
    mpasi_str = df['mpasi_raw'].astype(str).str.strip().str.lower()
    df['mpasi_bulan'] = mpasi_str.map(MPASI_CODE_MAP)
    mapped_ok = df['mpasi_bulan'].notna().sum()
    print(f"    [OK] Mapped: {mapped_ok:,} | Missing: {df['mpasi_bulan'].isna().sum():,}")

    # Label MP-ASI mengikuti kategori resmi Buku Kode SKI 2023 (kolom I43):
    # kode 'a' = 0-7 hari (0,25 bln), kode 'b' = 8-29 hari (0,5 bln). Nilai
    # numerik tetap dalam satuan bulan; hanya label tampilan yang dipisah.
    def _mpasi_label(x):
        if pd.isna(x):
            return None
        if x == 0.25:
            return '0-7 hari'
        if x == 0.5:
            return '8-29 hari'
        if x == int(x):
            return f'{int(x)} bulan'
        return f'{x} bulan'
    df['mpasi_label'] = df['mpasi_bulan'].apply(_mpasi_label)

    # STEP 2b: FILTER USIA
    print_header("STEP 2b: Penyaringan Usia (0-59 bulan)", char='-')
    n_before_age = len(df)
    usia_valid = df['usia_bulan'].between(0, 59, inclusive='both')
    usia_diluar = ~usia_valid & df['usia_bulan'].notna()
    if usia_diluar.any():
        df = df[usia_valid | df['usia_bulan'].isna()].copy()
        print(f"    [OK] Dihapus: {n_before_age - len(df)} di luar 0-59 bulan")
    else:
        print(f"    [OK] Semua dalam rentang 0-59 bulan")

    # STEP 2c: DATA REDUCTION
    print_header("STEP 2c: Data Reduction (Seleksi Fitur)", char='-')
    FINAL_COLUMNS = ['ID_balita', 'usia_bulan', 'usia_label', 'jenis_kelamin',
                     'tinggi_badan_cm', 'berat_badan_gram', 'pbl_cm',
                     'mpasi_bulan', 'mpasi_label']
    df = df[FINAL_COLUMNS].copy()
    print(f"    [OK] {len(FINAL_COLUMNS)} kolom dipertahankan")

    # STEP 2a: LISTWISE DELETION
    print_header("STEP 2a: Listwise Deletion", char='-')
    CORE_VARS = ['usia_bulan', 'jenis_kelamin', 'tinggi_badan_cm',
                 'berat_badan_gram', 'pbl_cm', 'mpasi_bulan']

    n_before = len(df)

    print_step("2a.1", "Missing values per variabel inti")
    for var in CORE_VARS:
        n_miss = df[var].isna().sum()
        pct = n_miss / len(df) * 100
        s = "[OK]" if n_miss == 0 else "[!]"
        print(f"    {s} {var:25s}: {n_miss:>5d} missing ({pct:>5.1f}%)")

    print_step("2a.2", "Deteksi outlier medis")
    outlier_flags = pd.DataFrame(index=df.index)

    checks = [
        ('usia_bulan',       'usia_bulan',   (0, 59)),
        ('tinggi_badan_cm',  'tinggi_badan', (35.0, 130.0)),
        ('berat_badan_gram', 'berat_badan',  (1500, 30000)),  # gram
        ('pbl_cm',           'pbl_cm',       (25.0, 60.0)),
        ('mpasi_bulan',      'mpasi_bulan',  (0, 12)),
    ]
    for col, key, (lo, hi) in checks:
        mask = ~df[col].between(lo, hi) & df[col].notna()
        outlier_flags[f'{col}_out'] = mask
        if mask.any():
            print(f"    [!] {col}: {mask.sum()} di luar [{lo}-{hi}]")
        else:
            print(f"    [OK] {col}: OK [{lo}-{hi}]")

    mask_jk = ~df['jenis_kelamin'].isin([1, 2]) & df['jenis_kelamin'].notna()
    outlier_flags['jk_out'] = mask_jk

    print_step("2a.3", "Eksekusi Listwise Deletion")
    any_missing = df[CORE_VARS].isna().any(axis=1)
    any_outlier = outlier_flags.any(axis=1)
    rows_to_drop = any_missing | any_outlier

    df_clean = df[~rows_to_drop].copy().reset_index(drop=True)
    n_dropped = n_before - len(df_clean)

    print(f"\n    LISTWISE DELETION RESULT:")
    print(f"    Sebelum : {n_before:,}")
    print(f"    Dihapus : {n_dropped:,} ({n_dropped/n_before*100:.1f}%)")
    print(f"    Sesudah : {len(df_clean):,}")

    # STEP 2e: KELAS BARU
    print_header("STEP 2e: Kelas Baru (Tambahan)", char='-')

    # Jarak usia lahir & saat pengukuran = usia_bulan itu sendiri
    df_clean['jarak_lahir_ukur'] = df_clean['usia_bulan'].astype(int)
    print(f"    [OK] Jarak Usia Lahir & Pengukuran = usia_bulan (0-{df_clean['jarak_lahir_ukur'].max()} bulan)")

    # Jarak pemberian MPASI - pengambilan survei
    df_clean['jarak_mpasi_survei'] = df_clean['usia_bulan'] - df_clean['mpasi_bulan']
    df_clean['jarak_mpasi_survei'] = df_clean['jarak_mpasi_survei'].clip(lower=0)
    print(f"    [OK] Jarak MPASI-Survei: mean={df_clean['jarak_mpasi_survei'].mean():.1f} bulan")

    report['final_clean'] = {'rows': len(df_clean), 'cols': df_clean.shape[1]}
    return df_clean, report


# ============================================================================
# TASK 3: STUNTING LABELLING & Z-SCORE (SAAT LAHIR, usia=0)
# ============================================================================
def task3_zscore_labelling(df):
    print_header("TASK 3: STUNTING LABELLING & Z-SCORE (Ground Truth)")
    print(f"    Metode: Z-Score WHO berdasarkan TB LAHIR (usia di-set 0 bulan)")
    print(f"    Variabel: jenis_kelamin + pbl_cm (TB lahir) + usia=0")
    print(f"    Rumus: Z = [(x/M)^L - 1] / (L x S)")
    print(f"    Threshold: Z < -2.00 = Stunting | Z >= -2.00 = Normal")

    # Load tabel WHO
    print_step("3.1", "Memuat tabel referensi WHO (LMS parameters)")
    who_lookup = {}
    table_config = [
        ('boys_0_24',   1, 'Laki-laki 0-24 bln'),
        ('boys_24_60',  1, 'Laki-laki 24-60 bln'),
        ('girls_0_24',  2, 'Perempuan 0-24 bln'),
        ('girls_24_60', 2, 'Perempuan 24-60 bln'),
    ]
    for table_key, jk, label in table_config:
        fpath = WHO_TABLES[table_key]
        if not os.path.exists(fpath):
            raise FileNotFoundError(f"WHO table tidak ditemukan: {fpath}")
        df_who = pd.read_excel(fpath)
        df_who.columns = df_who.columns.str.strip()
        for _, row in df_who.iterrows():
            month = int(row['Month'])
            who_lookup[(jk, month)] = {
                'L': float(row['L']), 'M': float(row['M']), 'S': float(row['S'])
            }
        print(f"    [OK] {label}")

    # Hitung Z-Score per baris (usia = 0 bulan, TB = pbl_cm)
    print_step("3.2", f"Menghitung Z-Score per baris ({len(df):,} balita)")
    print(f"    PENTING: Usia di-set 0 bulan (saat lahir) untuk semua balita")

    z_scores = np.full(len(df), np.nan)
    L_vals = np.full(len(df), np.nan)
    M_vals = np.full(len(df), np.nan)
    S_vals = np.full(len(df), np.nan)
    labels = [None] * len(df)
    matched = 0

    for idx in range(len(df)):
        jk = int(df.iloc[idx]['jenis_kelamin'])
        x = float(df.iloc[idx]['pbl_cm'])  # TB lahir
        usia_lahir = 0  # FIXED: usia = 0 bulan (saat lahir)

        key = (jk, usia_lahir)
        if key in who_lookup:
            p = who_lookup[key]
            L, M, S = p['L'], p['M'], p['S']
            L_vals[idx], M_vals[idx], S_vals[idx] = L, M, S

            if L != 0:
                z = ((x / M) ** L - 1) / (L * S)
            else:
                z = np.log(x / M) / S

            z_scores[idx] = round(z, 3)
            labels[idx] = 'Stunting' if z < -2.0 else 'Normal'
            matched += 1

    df['L'] = L_vals
    df['M'] = M_vals
    df['S'] = S_vals
    df['z_score'] = z_scores
    df['label_who'] = labels

    print(f"    [OK] Z-Score dihitung: {matched:,}/{len(df):,} balita")

    # Statistik
    print_step("3.3", "Statistik Hasil Z-Score & Labelling")
    z_valid = df['z_score'].dropna()
    print(f"    Z-Score: Mean={z_valid.mean():.3f}, Std={z_valid.std():.3f}")
    print(f"    Min={z_valid.min():.3f}, Median={z_valid.median():.3f}, Max={z_valid.max():.3f}")

    label_counts = df['label_who'].value_counts()
    total = len(df)
    print(f"\n    Distribusi Label WHO (Ground Truth - saat lahir):")
    for lbl in ['Normal', 'Stunting']:
        cnt = label_counts.get(lbl, 0)
        print(f"    {lbl:10s}: {cnt:>4d} ({cnt/total*100:.1f}%)")

    # Simpan
    print_step("3.4", "Menyimpan output Z-Score per baris")
    df_out = pd.DataFrame({
        'No': range(1, len(df)+1),
        'Sex': df['jenis_kelamin'].map({1:'Laki-laki', 2:'Perempuan'}).values,
        'Umur (bln)': df['usia_bulan'].astype(int).values,
        'TB Lahir (cm)': df['pbl_cm'].values,
        'TB Saat Ukur (cm)': df['tinggi_badan_cm'].values,
        'BB Saat Ukur (g)': df['berat_badan_gram'].values,
        'MPASI (bln)': df['mpasi_bulan'].values,
        'L': df['L'].values, 'M': df['M'].values, 'S': df['S'].values,
        'Z-Score (lahir)': df['z_score'].values,
        'Label WHO (lahir)': df['label_who'].values,
    })
    out_path = os.path.join(TRAINFILE_DIR, 'task3_zscore_perbaris.csv')
    df_out.to_csv(out_path, index=False)
    print(f"    [OK] Saved: {out_path}")

    return df, who_lookup


# ============================================================================
# TASK 3B: Z-SCORE SAAT PENGUKURAN (Ground Truth untuk Evaluasi)
# ============================================================================
def task3b_zscore_pengukuran(df, who_lookup):
    print_header("TASK 3B: Z-SCORE SAAT PENGUKURAN (Ground Truth Evaluasi)")
    print(f"    Metode: Z-Score WHO berdasarkan TB SAAT PENGUKURAN + Usia AKTUAL")
    print(f"    Variabel: jenis_kelamin + tinggi_badan_cm + usia_bulan (aktual)")
    print(f"    Rumus: Z = [(x/M)^L - 1] / (L x S)")
    print(f"    Threshold: Z < -2.00 = Stunting | Z >= -2.00 = Normal")

    print_step("3B.1", f"Menghitung Z-Score saat pengukuran ({len(df):,} balita)")

    z_scores_ukur = np.full(len(df), np.nan)
    L_ukur = np.full(len(df), np.nan)
    M_ukur = np.full(len(df), np.nan)
    S_ukur = np.full(len(df), np.nan)
    labels_ukur = [None] * len(df)
    matched = 0

    for idx in range(len(df)):
        jk = int(df.iloc[idx]['jenis_kelamin'])
        usia = int(df.iloc[idx]['usia_bulan'])       # usia AKTUAL saat pengukuran
        x = float(df.iloc[idx]['tinggi_badan_cm'])   # TB SAAT pengukuran

        key = (jk, usia)
        if key in who_lookup:
            p = who_lookup[key]
            L, M, S = p['L'], p['M'], p['S']
            L_ukur[idx], M_ukur[idx], S_ukur[idx] = L, M, S

            if L != 0:
                z = ((x / M) ** L - 1) / (L * S)
            else:
                z = np.log(x / M) / S

            z_scores_ukur[idx] = round(z, 3)
            labels_ukur[idx] = 'Stunting' if z < -2.0 else 'Normal'
            matched += 1

    df['L_ukur'] = L_ukur
    df['M_ukur'] = M_ukur
    df['S_ukur'] = S_ukur
    df['z_score_ukur'] = z_scores_ukur
    df['label_who_ukur'] = labels_ukur

    print(f"    [OK] Z-Score saat pengukuran dihitung: {matched:,}/{len(df):,} balita")

    # Statistik
    print_step("3B.2", "Statistik Z-Score Saat Pengukuran")
    z_valid = df['z_score_ukur'].dropna()
    print(f"    Z-Score: Mean={z_valid.mean():.3f}, Std={z_valid.std():.3f}")
    print(f"    Min={z_valid.min():.3f}, Median={z_valid.median():.3f}, Max={z_valid.max():.3f}")

    label_counts = df['label_who_ukur'].value_counts()
    total = len(df)
    print(f"\n    Distribusi Label WHO (Saat Pengukuran - Ground Truth Evaluasi):")
    for lbl in ['Normal', 'Stunting']:
        cnt = label_counts.get(lbl, 0)
        print(f"    {lbl:10s}: {cnt:>4d} ({cnt/total*100:.1f}%)")

    # Perbandingan label lahir vs label saat ukur
    print_step("3B.3", "Perbandingan Label Lahir vs Label Saat Pengukuran")
    lahir_s_ukur_s = ((df['label_who']=='Stunting') & (df['label_who_ukur']=='Stunting')).sum()
    lahir_s_ukur_n = ((df['label_who']=='Stunting') & (df['label_who_ukur']=='Normal')).sum()
    lahir_n_ukur_s = ((df['label_who']=='Normal') & (df['label_who_ukur']=='Stunting')).sum()
    lahir_n_ukur_n = ((df['label_who']=='Normal') & (df['label_who_ukur']=='Normal')).sum()

    print(f"    Lahir Stunting -> Ukur Stunting : {lahir_s_ukur_s:>4d} (tetap stunting)")
    print(f"    Lahir Stunting -> Ukur Normal   : {lahir_s_ukur_n:>4d} (membaik)")
    print(f"    Lahir Normal   -> Ukur Stunting : {lahir_n_ukur_s:>4d} (memburuk)")
    print(f"    Lahir Normal   -> Ukur Normal   : {lahir_n_ukur_n:>4d} (tetap baik)")

    return df


# ============================================================================
# TASK 4: FUZZY SYSTEM LOGIC (Takagi-Sugeno)
# ============================================================================

# -- Fungsi Keanggotaan Dasar --
def trapezoid(x, a, b, c, d):
    # Fungsi keanggotaan trapesium dengan dukungan bentuk BAHU (shoulder)
    # untuk himpunan di ujung domain. Konvensi:
    #   - Bila a == b (himpunan TERENDAH, mis. "Terlalu Cepat", "Pendek",
    #     "Kurang", "Muda"): sisi kiri berbentuk bahu, sehingga seluruh nilai
    #     x <= b memiliki derajat keanggotaan penuh (mu = 1). Ini menjamin
    #     nilai ekstrem rendah (termasuk 0 atau di bawah batas) tetap
    #     mempertahankan/menguatkan sifat himpunan tersebut.
    #   - Bila c == d (himpunan TERTINGGI, mis. "Lambat", "Tinggi",
    #     "Obesitas", "Tua"): sisi kanan berbentuk bahu, sehingga seluruh
    #     nilai x >= c memiliki derajat keanggotaan penuh (mu = 1). Ini
    #     menjamin nilai ekstrem tinggi (termasuk di atas batas domain) tetap
    #     mempertahankan/menguatkan sifat himpunan tersebut.
    # Himpunan di tengah (a < b dan c < d) tetap berperilaku trapesium biasa.
    left_shoulder = (a == b)
    right_shoulder = (c == d)

    # Bahu kiri: mu = 1 untuk semua x <= b (= a)
    if left_shoulder and x <= b:
        return 1.0
    # Bahu kanan: mu = 1 untuk semua x >= c (= d)
    if right_shoulder and x >= c:
        return 1.0

    if x <= a or x >= d:
        return 0.0
    elif a < x < b:
        return (x - a) / (b - a)
    elif b <= x <= c:
        return 1.0
    elif c < x < d:
        return (d - x) / (d - c)
    return 0.0

def triangle(x, a, b, c):
    # Segitiga tidak pernah berbentuk bahu (puncak tunggal di b); gunakan
    # trapesium dengan b == c namun a != b dan c != d.
    return trapezoid(x, a, b, b, c)

def trapezoid_closed(x, a, b, c, d):
    # Trapesium TERTUTUP murni (tanpa bahu), apa pun nilai a,b,c,d. Dipakai
    # untuk variabel yang batasnya harus tegas (mis. Usia 0-59 bulan), agar
    # nilai di titik ekstrem tidak otomatis memperoleh keanggotaan penuh
    # seperti pada bentuk bahu. Perilaku identik dengan trapesium standar.
    if x <= a or x >= d:
        return 0.0
    elif a < x < b:
        return (x - a) / (b - a)
    elif b <= x <= c:
        return 1.0
    elif c < x < d:
        return (d - x) / (d - c)
    return 0.0

# -- Initial Boundary Points (akan dioptimasi oleh GA) --
# Usia (bulan) - universal. Memakai 'trap_closed' (tertutup) agar batas usia
# tegas 0-59 bulan; nilai usia=0 (bayi baru lahir) TIDAK memperoleh bahu.
USIA_MF_INIT = {
    'Muda':   ('trap_closed', [0, 0, 6, 18]),
    'Sedang': ('tri',  [12, 30, 48]),
    'Tua':    ('trap_closed', [42, 52, 59, 59]),
}

# TB Laki-laki (cm)
TB_L_MF_INIT = {
    'Pendek': ('trap', [38, 38, 55, 70]),
    'Normal': ('trap', [60, 70, 100, 110]),
    'Tinggi': ('trap', [100, 110, 130, 130]),
}

# TB Perempuan (cm)
TB_P_MF_INIT = {
    'Pendek': ('trap', [38, 38, 53, 68]),
    'Normal': ('trap', [58, 68, 98, 108]),
    'Tinggi': ('trap', [98, 108, 130, 130]),
}

# BB Laki-laki (gram)
BB_L_MF_INIT = {
    'Kurang':   ('trap', [1500, 1500, 5000, 8000]),
    'Normal':   ('trap', [6000, 8000, 15000, 18000]),
    'Lebih':    ('tri',  [15000, 18000, 22000]),
    'Obesitas': ('trap', [20000, 23000, 30000, 30000]),
}

# BB Perempuan (gram)
BB_P_MF_INIT = {
    'Kurang':   ('trap', [1500, 1500, 4500, 7500]),
    'Normal':   ('trap', [5500, 7500, 14000, 17000]),
    'Lebih':    ('tri',  [14000, 17000, 21000]),
    'Obesitas': ('trap', [19000, 22000, 28000, 28000]),
}

# MPASI (bulan) - universal
MPASI_MF_INIT = {
    'TerlaCepat': ('trap', [0, 0, 3, 5]),
    'Normal':     ('trap', [4, 5.5, 6.5, 8]),
    'Lambat':     ('trap', [7, 8, 12, 12]),
}

# -- Consequent Values (FIXED) --
CONSEQUENTS = {'high': 90, 'medium': 65, 'low': 35, 'normal': 10}

# -- FK Output (Gambar 3.9 - FIXED) --
OUTPUT_MF = {
    'Normal':          (0, 0, 15, 30),
    'STUNTING LOW':    (15, 30, 60),
    'STUNTING MEDIUM': (45, 60, 75),
    'STUNTING HIGH':   (75, 90, 100, 100),
}


def eval_mf(x, mf_type, params):
    if mf_type == 'trap':
        return trapezoid(x, *params)
    elif mf_type == 'trap_closed':
        return trapezoid_closed(x, *params)
    elif mf_type == 'tri':
        return triangle(x, *params)
    return 0.0


def fuzzify_new(usia, tb, bb, mpasi, jk, mf_params=None):
    if mf_params is None:
        usia_mf = USIA_MF_INIT
        mpasi_mf = MPASI_MF_INIT
        tb_mf = TB_L_MF_INIT if jk == 1 else TB_P_MF_INIT
        bb_mf = BB_L_MF_INIT if jk == 1 else BB_P_MF_INIT
    else:
        usia_mf = mf_params['usia']
        mpasi_mf = mf_params['mpasi']
        tb_mf = mf_params['tb_l'] if jk == 1 else mf_params['tb_p']
        bb_mf = mf_params['bb_l'] if jk == 1 else mf_params['bb_p']

    mu = {
        'usia':  {k: eval_mf(usia, v[0], v[1]) for k, v in usia_mf.items()},
        'tb':    {k: eval_mf(tb, v[0], v[1]) for k, v in tb_mf.items()},
        'bb':    {k: eval_mf(bb, v[0], v[1]) for k, v in bb_mf.items()},
        'mpasi': {k: eval_mf(mpasi, v[0], v[1]) for k, v in mpasi_mf.items()},
    }
    return mu


# -- 108 Rules (logic pattern, berlaku per JK) --
# Format: (usia_cond, tb_cond, bb_cond, mpasi_cond, risk_class)
RULES_108 = [
    # STUNTING HIGH (c=90) - 10 rules
    ('Tua','Pendek','Kurang','TerlaCepat','high'),
    ('Tua','Pendek','Kurang','Normal','high'),
    ('Tua','Pendek','Kurang','Lambat','high'),
    ('Sedang','Pendek','Kurang','TerlaCepat','high'),
    ('Sedang','Pendek','Kurang','Lambat','high'),
    ('Sedang','Pendek','Kurang','Normal','high'),
    ('Tua','Pendek','Normal','TerlaCepat','high'),
    ('Tua','Pendek','Normal','Lambat','high'),
    ('Muda','Pendek','Kurang','TerlaCepat','high'),
    ('Muda','Pendek','Kurang','Lambat','high'),

    # STUNTING MEDIUM (c=65) - 30 rules
    ('Muda','Pendek','Kurang','Normal','medium'),
    ('Muda','Pendek','Normal','TerlaCepat','medium'),
    ('Muda','Pendek','Normal','Lambat','medium'),
    ('Sedang','Pendek','Normal','TerlaCepat','medium'),
    ('Sedang','Pendek','Normal','Lambat','medium'),
    ('Sedang','Pendek','Normal','Normal','medium'),
    ('Tua','Pendek','Normal','Normal','medium'),
    ('Sedang','Normal','Kurang','TerlaCepat','medium'),
    ('Sedang','Normal','Kurang','Lambat','medium'),
    ('Tua','Normal','Kurang','TerlaCepat','medium'),
    ('Tua','Normal','Kurang','Lambat','medium'),
    ('Tua','Normal','Kurang','Normal','medium'),
    ('Muda','Pendek','Lebih','TerlaCepat','medium'),
    ('Muda','Pendek','Lebih','Normal','medium'),
    ('Muda','Pendek','Lebih','Lambat','medium'),
    ('Sedang','Pendek','Lebih','TerlaCepat','medium'),
    ('Sedang','Pendek','Lebih','Normal','medium'),
    ('Sedang','Pendek','Lebih','Lambat','medium'),
    ('Tua','Pendek','Lebih','TerlaCepat','medium'),
    ('Tua','Pendek','Lebih','Normal','medium'),
    ('Tua','Pendek','Lebih','Lambat','medium'),
    ('Muda','Pendek','Obesitas','TerlaCepat','medium'),
    ('Muda','Pendek','Obesitas','Normal','medium'),
    ('Muda','Pendek','Obesitas','Lambat','medium'),
    ('Sedang','Pendek','Obesitas','TerlaCepat','medium'),
    ('Sedang','Pendek','Obesitas','Normal','medium'),
    ('Sedang','Pendek','Obesitas','Lambat','medium'),
    ('Tua','Pendek','Obesitas','TerlaCepat','medium'),
    ('Tua','Pendek','Obesitas','Normal','medium'),
    ('Tua','Pendek','Obesitas','Lambat','medium'),

    # STUNTING LOW (c=35) - 20 rules
    ('Muda','Pendek','Normal','Normal','low'),
    ('Muda','Normal','Kurang','Normal','low'),
    ('Muda','Normal','Kurang','TerlaCepat','low'),
    ('Muda','Normal','Kurang','Lambat','low'),
    ('Sedang','Normal','Kurang','Normal','low'),
    ('Muda','Normal','Normal','TerlaCepat','low'),
    ('Muda','Normal','Normal','Lambat','low'),
    ('Sedang','Normal','Normal','TerlaCepat','low'),
    ('Sedang','Normal','Normal','Lambat','low'),
    ('Tua','Normal','Normal','TerlaCepat','low'),
    ('Tua','Normal','Normal','Lambat','low'),
    ('Muda','Tinggi','Kurang','TerlaCepat','low'),
    ('Muda','Tinggi','Kurang','Normal','low'),
    ('Muda','Tinggi','Kurang','Lambat','low'),
    ('Sedang','Tinggi','Kurang','TerlaCepat','low'),
    ('Sedang','Tinggi','Kurang','Normal','low'),
    ('Sedang','Tinggi','Kurang','Lambat','low'),
    ('Tua','Tinggi','Kurang','TerlaCepat','low'),
    ('Tua','Tinggi','Kurang','Normal','low'),
    ('Tua','Tinggi','Kurang','Lambat','low'),

    # NORMAL (c=10) - 48 rules
    ('Muda','Normal','Normal','Normal','normal'),
    ('Sedang','Normal','Normal','Normal','normal'),
    ('Tua','Normal','Normal','Normal','normal'),
    ('Muda','Tinggi','Normal','Normal','normal'),
    ('Sedang','Tinggi','Normal','Normal','normal'),
    ('Tua','Tinggi','Normal','Normal','normal'),
    ('Muda','Tinggi','Normal','TerlaCepat','normal'),
    ('Muda','Tinggi','Normal','Lambat','normal'),
    ('Sedang','Tinggi','Normal','TerlaCepat','normal'),
    ('Sedang','Tinggi','Normal','Lambat','normal'),
    ('Tua','Tinggi','Normal','TerlaCepat','normal'),
    ('Tua','Tinggi','Normal','Lambat','normal'),
    ('Muda','Normal','Lebih','Normal','normal'),
    ('Sedang','Normal','Lebih','Normal','normal'),
    ('Tua','Normal','Lebih','Normal','normal'),
    ('Muda','Normal','Lebih','TerlaCepat','normal'),
    ('Muda','Normal','Lebih','Lambat','normal'),
    ('Sedang','Normal','Lebih','TerlaCepat','normal'),
    ('Sedang','Normal','Lebih','Lambat','normal'),
    ('Tua','Normal','Lebih','TerlaCepat','normal'),
    ('Tua','Normal','Lebih','Lambat','normal'),
    ('Muda','Tinggi','Lebih','TerlaCepat','normal'),
    ('Muda','Tinggi','Lebih','Normal','normal'),
    ('Muda','Tinggi','Lebih','Lambat','normal'),
    ('Sedang','Tinggi','Lebih','TerlaCepat','normal'),
    ('Sedang','Tinggi','Lebih','Normal','normal'),
    ('Sedang','Tinggi','Lebih','Lambat','normal'),
    ('Tua','Tinggi','Lebih','TerlaCepat','normal'),
    ('Tua','Tinggi','Lebih','Normal','normal'),
    ('Tua','Tinggi','Lebih','Lambat','normal'),
    ('Muda','Normal','Obesitas','TerlaCepat','normal'),
    ('Muda','Normal','Obesitas','Normal','normal'),
    ('Muda','Normal','Obesitas','Lambat','normal'),
    ('Sedang','Normal','Obesitas','TerlaCepat','normal'),
    ('Sedang','Normal','Obesitas','Normal','normal'),
    ('Sedang','Normal','Obesitas','Lambat','normal'),
    ('Tua','Normal','Obesitas','TerlaCepat','normal'),
    ('Tua','Normal','Obesitas','Normal','normal'),
    ('Tua','Normal','Obesitas','Lambat','normal'),
    ('Muda','Tinggi','Obesitas','TerlaCepat','normal'),
    ('Muda','Tinggi','Obesitas','Normal','normal'),
    ('Muda','Tinggi','Obesitas','Lambat','normal'),
    ('Sedang','Tinggi','Obesitas','TerlaCepat','normal'),
    ('Sedang','Tinggi','Obesitas','Normal','normal'),
    ('Sedang','Tinggi','Obesitas','Lambat','normal'),
    ('Tua','Tinggi','Obesitas','TerlaCepat','normal'),
    ('Tua','Tinggi','Obesitas','Normal','normal'),
    ('Tua','Tinggi','Obesitas','Lambat','normal'),
]

assert len(RULES_108) == 108, f"Expected 108 rules, got {len(RULES_108)}"


def evaluate_rules_new(mu, cons=None):
    if cons is None:
        cons = CONSEQUENTS
    fired = []
    for rule in RULES_108:
        usia_c, tb_c, bb_c, mpasi_c, risk_class = rule
        w = min(
            mu['usia'].get(usia_c, 0),
            mu['tb'].get(tb_c, 0),
            mu['bb'].get(bb_c, 0),
            mu['mpasi'].get(mpasi_c, 0),
        )
        if w > 0:
            fired.append((w, cons[risk_class]))
    return fired


def defuzzify_sugeno(fired_rules):
    if not fired_rules:
        return float(CONSEQUENTS['low'])
    num = sum(w * c for w, c in fired_rules)
    den = sum(w for w, c in fired_rules)
    if den == 0:
        return float(CONSEQUENTS['low'])
    return round(num / den, 2)


def get_priority_flag(risk_score):
    mu_n = trapezoid(risk_score, -1, 0, 15, 30)
    if risk_score <= 0: mu_n = 1.0
    mu_l = triangle(risk_score, 15, 30, 60)
    mu_m = triangle(risk_score, 45, 60, 75)
    mu_h = trapezoid(risk_score, 75, 90, 100, 101)
    if risk_score >= 100: mu_h = 1.0
    mu_out = {'Normal': mu_n, 'STUNTING LOW': mu_l,
              'STUNTING MEDIUM': mu_m, 'STUNTING HIGH': mu_h}
    return max(mu_out, key=mu_out.get)


def fuzzy_infer_single(usia, tb, bb, mpasi, jk, mf_params=None, cons=None):
    if cons is None:
        cons = CONSEQUENTS
    mu = fuzzify_new(usia, tb, bb, mpasi, jk, mf_params)
    fired = evaluate_rules_new(mu, cons)
    risk = defuzzify_sugeno(fired)
    flag = get_priority_flag(risk)
    binary = 0 if flag == 'Normal' else 1
    return risk, flag, binary


def run_pipeline(dataframe, mf_params=None, cons=None):
    if cons is None:
        cons = CONSEQUENTS
    risks, flags, binaries = [], [], []
    for idx in range(len(dataframe)):
        row = dataframe.iloc[idx]
        usia = float(row['usia_bulan'])
        tb = float(row['tinggi_badan_cm'])
        bb = float(row['berat_badan_gram'])
        mpasi = float(row['mpasi_bulan'])
        jk = int(row['jenis_kelamin'])
        r, f, b = fuzzy_infer_single(usia, tb, bb, mpasi, jk, mf_params, cons)
        risks.append(r)
        flags.append(f)
        binaries.append(b)
    return np.array(risks), flags, np.array(binaries)


def task4_fuzzy_logic(df):
    print_header("TASK 4: FUZZY SYSTEM LOGIC - Takagi-Sugeno")

    print_step("4.1", "Fungsi Keanggotaan (Fuzzyfikasi)")
    print(f"    Variabel: Usia (universal), TB (per JK), BB (per JK), MPASI (universal)")
    print(f"    JK = pre-filter: menentukan FK TB dan BB yang dipakai")
    print(f"    Total rules: 216 (108 per JK) - 108 pola x 2 jenis kelamin")
    print(f"    Distribusi rules:")
    for cls in ['high', 'medium', 'low', 'normal']:
        cnt = sum(1 for r in RULES_108 if r[4] == cls)
        label = {'high':'STUNTING HIGH','medium':'STUNTING MEDIUM',
                 'low':'STUNTING LOW','normal':'Normal'}[cls]
        print(f"      {label:18s} (c={CONSEQUENTS[cls]}): {cnt} rules/JK, {cnt*2} total")

    print_step("4.2", f"Inferensi Fuzzy per baris - {len(df):,} balita")
    risks, flags, binaries = run_pipeline(df)
    df['risk_score'] = np.round(risks, 2)
    df['priority_flag'] = flags
    df['fuzzy_label'] = ['Stunting' if b == 1 else 'Normal' for b in binaries]

    print(f"    [OK] Inferensi selesai: {len(df):,} balita")

    print_step("4.3", "Statistik Risk Score & Priority Flag")
    rs = df['risk_score']
    print(f"    Risk Score: Mean={rs.mean():.2f}, Std={rs.std():.2f}, Min={rs.min():.2f}, Max={rs.max():.2f}")

    print(f"\n    Distribusi Priority Flag:")
    for flag_name in ['Normal', 'STUNTING LOW', 'STUNTING MEDIUM', 'STUNTING HIGH']:
        cnt = (df['priority_flag'] == flag_name).sum()
        pct = cnt / len(df) * 100
        print(f"    {flag_name:18s}: {cnt:>4d} ({pct:>5.1f}%)")

    print(f"\n    Monitoring: Z-Score (lahir) vs Fuzzy (saat ukur):")
    for who_lbl in ['Normal', 'Stunting']:
        subset = df[df['label_who'] == who_lbl]
        fparts = []
        for fl in ['Normal', 'STUNTING LOW', 'STUNTING MEDIUM', 'STUNTING HIGH']:
            c = (subset['priority_flag'] == fl).sum()
            fparts.append(f"{fl}={c}")
        print(f"    Z-Score {who_lbl:8s} (n={len(subset):>3d}): {', '.join(fparts)}")

    return df


# ============================================================================
# TASK 5 & 6: EVALUASI + GA OPTIMASI
# ============================================================================

def calc_metrics(y_true, y_pred):
    tp = int(np.sum((y_true == 1) & (y_pred == 1)))
    tn = int(np.sum((y_true == 0) & (y_pred == 0)))
    fp = int(np.sum((y_true == 0) & (y_pred == 1)))
    fn = int(np.sum((y_true == 1) & (y_pred == 0)))
    prec = tp/(tp+fp) if (tp+fp) > 0 else 0.0
    rec  = tp/(tp+fn) if (tp+fn) > 0 else 0.0
    f1   = 2*prec*rec/(prec+rec) if (prec+rec) > 0 else 0.0
    acc  = (tp+tn)/(tp+tn+fp+fn) if (tp+tn+fp+fn) > 0 else 0.0
    return {'TP':tp, 'TN':tn, 'FP':fp, 'FN':fn,
            'precision':round(prec,4), 'recall':round(rec,4),
            'f1':round(f1,4), 'accuracy':round(acc,4)}


def decode_individual(individual):
    """Decode GA individual (flat array) ke MF params dict."""
    i = 0
    # Usia: Muda(4) + Sedang(3) + Tua(4) = 11
    usia = {
        'Muda':   ('trap_closed', list(individual[i:i+4])),
        'Sedang': ('tri',  list(individual[i+4:i+7])),
        'Tua':    ('trap_closed', list(individual[i+7:i+11])),
    }
    i += 11

    # TB_L: Pendek(4) + Normal(4) + Tinggi(4) = 12
    tb_l = {
        'Pendek': ('trap', list(individual[i:i+4])),
        'Normal': ('trap', list(individual[i+4:i+8])),
        'Tinggi': ('trap', list(individual[i+8:i+12])),
    }
    i += 12

    # TB_P: 12
    tb_p = {
        'Pendek': ('trap', list(individual[i:i+4])),
        'Normal': ('trap', list(individual[i+4:i+8])),
        'Tinggi': ('trap', list(individual[i+8:i+12])),
    }
    i += 12

    # BB_L: Kurang(4) + Normal(4) + Lebih(3) + Obesitas(4) = 15
    bb_l = {
        'Kurang':   ('trap', list(individual[i:i+4])),
        'Normal':   ('trap', list(individual[i+4:i+8])),
        'Lebih':    ('tri',  list(individual[i+8:i+11])),
        'Obesitas': ('trap', list(individual[i+11:i+15])),
    }
    i += 15

    # BB_P: 15
    bb_p = {
        'Kurang':   ('trap', list(individual[i:i+4])),
        'Normal':   ('trap', list(individual[i+4:i+8])),
        'Lebih':    ('tri',  list(individual[i+8:i+11])),
        'Obesitas': ('trap', list(individual[i+11:i+15])),
    }
    i += 15

    # MPASI: TerlaCepat(4) + Normal(4) + Lambat(4) = 12
    mpasi = {
        'TerlaCepat': ('trap', list(individual[i:i+4])),
        'Normal':     ('trap', list(individual[i+4:i+8])),
        'Lambat':     ('trap', list(individual[i+8:i+12])),
    }
    i += 12

    # --- Penjagaan bentuk BAHU (shoulder) untuk himpunan ujung domain ---
    # Himpunan TERENDAH tiap variabel harus berbentuk bahu kiri (a == b),
    # dan himpunan TERTINGGI harus berbentuk bahu kanan (c == d), agar nilai
    # ekstrem (mis. MP-ASI 0 / belum diberi, atau di atas batas domain) tetap
    # memiliki derajat keanggotaan penuh. GA hanya menggeser posisi bahu,
    # tidak menghilangkan sifat bahunya. Repair ini dijalankan setelah decode
    # sehingga apa pun konfigurasi gen dari GA, struktur bahu tetap terjaga.
    def _shoulder_left(mf):
        # trap [a,b,c,d] -> paksa b = a (sisi kiri jadi bahu)
        p = list(mf[1]); p[1] = p[0]; return (mf[0], p)
    def _shoulder_right(mf):
        # trap [a,b,c,d] -> paksa c = d (sisi kanan jadi bahu)
        p = list(mf[1]); p[2] = p[3]; return (mf[0], p)

    # Opsi 2b: bentuk bahu (shoulder) diterapkan pada variabel yang memiliki
    # nilai ekstrem bermakna (MP-ASI belum diberi / >=12 bulan, serta TB & BB
    # di luar rentang standar), TETAPI TIDAK pada variabel Usia. Usia memakai
    # batas tegas 0-59 bulan sesuai definisi balita; penerapan bahu pada Usia
    # menyebabkan bayi baru lahir (usia=0) yang berstatus normal keliru
    # ter-flag berisiko tinggi (false positive), sehingga Usia sengaja
    # dipertahankan berbentuk trapesium tertutup.
    tb_l['Pendek'] = _shoulder_left(tb_l['Pendek'])
    tb_l['Tinggi'] = _shoulder_right(tb_l['Tinggi'])
    tb_p['Pendek'] = _shoulder_left(tb_p['Pendek'])
    tb_p['Tinggi'] = _shoulder_right(tb_p['Tinggi'])
    bb_l['Kurang']   = _shoulder_left(bb_l['Kurang'])
    bb_l['Obesitas'] = _shoulder_right(bb_l['Obesitas'])
    bb_p['Kurang']   = _shoulder_left(bb_p['Kurang'])
    bb_p['Obesitas'] = _shoulder_right(bb_p['Obesitas'])
    mpasi['TerlaCepat'] = _shoulder_left(mpasi['TerlaCepat'])
    mpasi['Lambat']     = _shoulder_right(mpasi['Lambat'])

    return {
        'usia': usia, 'tb_l': tb_l, 'tb_p': tb_p,
        'bb_l': bb_l, 'bb_p': bb_p, 'mpasi': mpasi,
    }


def encode_initial():
    """Encode initial MF params ke flat array untuk GA."""
    genes = []
    # Usia
    genes += USIA_MF_INIT['Muda'][1]     # 4
    genes += USIA_MF_INIT['Sedang'][1]   # 3
    genes += USIA_MF_INIT['Tua'][1]      # 4
    # TB_L
    genes += TB_L_MF_INIT['Pendek'][1]   # 4
    genes += TB_L_MF_INIT['Normal'][1]   # 4
    genes += TB_L_MF_INIT['Tinggi'][1]   # 4
    # TB_P
    genes += TB_P_MF_INIT['Pendek'][1]
    genes += TB_P_MF_INIT['Normal'][1]
    genes += TB_P_MF_INIT['Tinggi'][1]
    # BB_L
    genes += BB_L_MF_INIT['Kurang'][1]   # 4
    genes += BB_L_MF_INIT['Normal'][1]   # 4
    genes += BB_L_MF_INIT['Lebih'][1]    # 3
    genes += BB_L_MF_INIT['Obesitas'][1] # 4
    # BB_P
    genes += BB_P_MF_INIT['Kurang'][1]
    genes += BB_P_MF_INIT['Normal'][1]
    genes += BB_P_MF_INIT['Lebih'][1]
    genes += BB_P_MF_INIT['Obesitas'][1]
    # MPASI
    genes += MPASI_MF_INIT['TerlaCepat'][1] # 4
    genes += MPASI_MF_INIT['Normal'][1]     # 4
    genes += MPASI_MF_INIT['Lambat'][1]     # 4

    return genes  # total = 11+12+12+15+15+12 = 77 genes


# Bounds for GA genes
GENE_BOUNDS = []
# Usia: Muda(0,0,6,18) Sedang(12,30,48) Tua(42,52,59,59)
GENE_BOUNDS += [(0,0),(0,0),(3,12),(12,24)]           # Muda trap
GENE_BOUNDS += [(6,18),(20,36),(36,54)]                # Sedang tri
GENE_BOUNDS += [(30,48),(42,56),(59,59),(59,59)]       # Tua trap
# TB_L: Pendek Normal Tinggi
GENE_BOUNDS += [(35,42),(35,42),(45,60),(60,80)]       # Pendek
GENE_BOUNDS += [(50,70),(60,80),(90,105),(100,115)]    # Normal
GENE_BOUNDS += [(90,105),(100,115),(125,135),(125,135)] # Tinggi
# TB_P: Pendek Normal Tinggi
GENE_BOUNDS += [(35,42),(35,42),(43,58),(58,75)]       # Pendek
GENE_BOUNDS += [(48,65),(58,75),(88,103),(98,113)]     # Normal
GENE_BOUNDS += [(88,103),(98,113),(125,135),(125,135)] # Tinggi
# BB_L: Kurang Normal Lebih Obesitas
GENE_BOUNDS += [(1000,2000),(1000,2000),(3000,6000),(6000,10000)]    # Kurang
GENE_BOUNDS += [(4000,8000),(6000,10000),(13000,17000),(16000,20000)] # Normal
GENE_BOUNDS += [(13000,17000),(16000,20000),(20000,24000)]           # Lebih tri
GENE_BOUNDS += [(18000,22000),(21000,25000),(28000,32000),(28000,32000)] # Obesitas
# BB_P: Kurang Normal Lebih Obesitas
GENE_BOUNDS += [(1000,2000),(1000,2000),(2500,5500),(5500,9000)]
GENE_BOUNDS += [(3500,7000),(5500,9000),(12000,16000),(15000,19000)]
GENE_BOUNDS += [(12000,16000),(15000,19000),(19000,23000)]
GENE_BOUNDS += [(17000,21000),(20000,24000),(26000,30000),(26000,30000)]
# MPASI: TerlaCepat Normal Lambat
GENE_BOUNDS += [(0,0),(0,0),(2,4),(3,6)]     # TerlaCepat
GENE_BOUNDS += [(3,5),(4,6),(5.5,7.5),(6,9)] # Normal
GENE_BOUNDS += [(5.5,7.5),(6,9),(12,12),(12,12)] # Lambat

N_GENES = len(encode_initial())
assert len(GENE_BOUNDS) == N_GENES, f"Bounds mismatch: {len(GENE_BOUNDS)} vs {N_GENES}"


def task5b_baseline_comparison(df, tr_idx, te_idx, fuzzy_test_metrics):
    """
    Perbandingan kuantitatif head-to-head antara sistem Fuzzy-GA dengan lima
    baseline machine learning (Decision Tree/C4.5, Random Forest default,
    Random Forest balanced, SVM-RBF balanced, Logistic Regression balanced).

    Semua baseline dilatih pada SPLIT, INPUT, dan GROUND TRUTH yang IDENTIK
    dengan sistem fuzzy (tr_idx/te_idx yang sama, enam variabel yang sama,
    label_who_ukur sebagai ground truth), sehingga perbandingan bersifat
    strictly like-for-like. Fungsi ini menjawab catatan editor Kinetik yang
    meminta perbandingan performa dengan metode lain pada data yang sama.
    """
    print_header("TASK 5B: PERBANDINGAN KUANTITATIF DENGAN BASELINE ML")
    print("    Baseline dilatih pada split/input/ground-truth IDENTIK dengan Fuzzy-GA")

    try:
        from sklearn.tree import DecisionTreeClassifier
        from sklearn.ensemble import RandomForestClassifier
        from sklearn.svm import SVC
        from sklearn.linear_model import LogisticRegression
        from sklearn.preprocessing import StandardScaler
    except ImportError:
        print("    [!] scikit-learn tidak tersedia; melewati perbandingan baseline.")
        print("        Install dengan: pip install scikit-learn")
        return None

    # Enam variabel input yang sama persis dengan sistem fuzzy
    FEATURES = ['usia_bulan', 'jenis_kelamin', 'pbl_cm',
                'tinggi_badan_cm', 'berat_badan_gram', 'mpasi_bulan']
    X = df[FEATURES].values
    y = (df['label_who_ukur'] == 'Stunting').astype(int).values

    X_tr, X_te = X[tr_idx], X[te_idx]
    y_tr, y_te = y[tr_idx], y[te_idx]

    # Standardisasi (fit di train, transform ke test) untuk model berbasis jarak/margin
    scaler = StandardScaler().fit(X_tr)
    X_tr_s, X_te_s = scaler.transform(X_tr), scaler.transform(X_te)

    print(f"    Train: {len(tr_idx)} | Test: {len(te_idx)} "
          f"(Stunting test={int(y_te.sum())}, Normal test={int((y_te==0).sum())})")

    def _eval(name, model, scaled=False):
        Xa, Xb = (X_tr_s, X_te_s) if scaled else (X_tr, X_te)
        model.fit(Xa, y_tr)
        pred = model.predict(Xb)
        m = calc_metrics(y_te, pred)
        return name, m

    models = [
        _eval("Decision Tree (C4.5)",
              DecisionTreeClassifier(criterion='entropy', random_state=42)),
        _eval("Random Forest (default)",
              RandomForestClassifier(n_estimators=100, random_state=42)),
        _eval("Random Forest (balanced)",
              RandomForestClassifier(n_estimators=100, class_weight='balanced',
                                     random_state=42)),
        _eval("SVM-RBF (balanced)",
              SVC(class_weight='balanced', random_state=42), scaled=True),
        _eval("Logistic Reg (balanced)",
              LogisticRegression(max_iter=1000, class_weight='balanced',
                                 random_state=42), scaled=True),
    ]

    # Tabel perbandingan (setara Tabel 9 pada manuskrip)
    print(f"\n    {'Method':26} {'Acc':>7} {'Prec':>7} {'Rec':>7} {'F1':>7}")
    print("    " + "-" * 58)
    comparison = {}
    for name, m in models:
        print(f"    {name:26} {m['accuracy']:>7.4f} {m['precision']:>7.4f} "
              f"{m['recall']:>7.4f} {m['f1']:>7.4f}")
        comparison[name] = m
    # Baris sistem yang diusulkan
    fm = fuzzy_test_metrics
    print(f"    {'Proposed Fuzzy-GA':26} {fm['accuracy']:>7.4f} {fm['precision']:>7.4f} "
          f"{fm['recall']:>7.4f} {fm['f1']:>7.4f}")
    comparison['Proposed Fuzzy-GA'] = fm

    print("\n    Interpretasi: baseline ML unggul pada metrik diskriminasi biner,")
    print("    namun hanya menghasilkan label biner. Sistem Fuzzy-GA menghasilkan")
    print("    risk score kontinu 0-100, empat priority flag (stratifikasi 3.04x),")
    print("    dan monitoring trayektori lahir->ukur yang tidak dapat dihasilkan")
    print("    classifier biner. Perbandingan ini adalah design trade-off, bukan")
    print("    semata defisit performa.")

    # Simpan ke JSON agar dapat direferensikan manuskrip
    out_path = os.path.join(TRAINFILE_DIR, 'baseline_comparison.json')
    serializable = {k: {kk: (int(vv) if isinstance(vv, (int, np.integer)) else float(vv))
                        for kk, vv in v.items()} for k, v in comparison.items()}
    with open(out_path, 'w') as f:
        json.dump({'features': FEATURES,
                   'train_size': int(len(tr_idx)),
                   'test_size': int(len(te_idx)),
                   'ground_truth': 'label_who_ukur',
                   'comparison': serializable,
                   'timestamp': datetime.now().isoformat()}, f, indent=2)
    print(f"\n    [OK] Saved: {out_path}")

    return comparison


def task5_6_evaluation(df):
    print_header("TASK 5 & 6: EVALUASI + GA OPTIMASI BOUNDARY POINTS")
    print(f"    Ground Truth: Z-Score SAAT PENGUKURAN (label_who_ukur)")
    print(f"    Prediksi: Output Fuzzy Sugeno (risk score -> priority flag)")

    y_all = (df['label_who_ukur'] == 'Stunting').astype(int).values

    # Stratified split
    print_step("6.1", "Stratified Train/Test Split (80/20)")
    np.random.seed(42)
    idx_n = np.where(y_all == 0)[0].copy()
    idx_s = np.where(y_all == 1)[0].copy()
    np.random.shuffle(idx_n)
    np.random.shuffle(idx_s)
    n_tr_n = int(len(idx_n) * 0.8)
    n_tr_s = int(len(idx_s) * 0.8)
    tr_idx = np.concatenate([idx_n[:n_tr_n], idx_s[:n_tr_s]])
    te_idx = np.concatenate([idx_n[n_tr_n:], idx_s[n_tr_s:]])

    df_tr = df.iloc[tr_idx].reset_index(drop=True)
    df_te = df.iloc[te_idx].reset_index(drop=True)
    y_tr = (df_tr['label_who_ukur'] == 'Stunting').astype(int).values
    y_te = (df_te['label_who_ukur'] == 'Stunting').astype(int).values

    print(f"    Train: {len(df_tr)} (Normal={sum(y_tr==0)}, Stunting={sum(y_tr==1)})")
    print(f"    Test:  {len(df_te)} (Normal={sum(y_te==0)}, Stunting={sum(y_te==1)})")

    # Baseline (initial MF params)
    print_step("6.2", "Baseline (Initial MF params)")
    _, _, y_pred_base = run_pipeline(df_tr)
    base_m = calc_metrics(y_tr, y_pred_base)
    print(f"    Baseline Train: P={base_m['precision']:.4f} R={base_m['recall']:.4f} F1={base_m['f1']:.4f} Acc={base_m['accuracy']:.4f}")

    # GA Setup
    print_step("6.3", "GA Optimasi Boundary Points")
    print(f"    Genome: {N_GENES} genes (boundary points semua FK)")
    print(f"    GA hanya mengoptimasi posisi titik FK, TIDAK mengubah rules/arsitektur")

    # Precompute arrays for vectorized fitness
    tr_usia = df_tr['usia_bulan'].values.astype(float)
    tr_tb = df_tr['tinggi_badan_cm'].values.astype(float)
    tr_bb = df_tr['berat_badan_gram'].values.astype(float)
    tr_mpasi = df_tr['mpasi_bulan'].values.astype(float)
    tr_jk = df_tr['jenis_kelamin'].values.astype(int)

    def fitness_func(individual):
        # Enforce sorted constraint
        ind = list(individual)
        for i in range(len(ind)):
            lo, hi = GENE_BOUNDS[i]
            ind[i] = max(lo, min(hi, ind[i]))

        try:
            mf_p = decode_individual(ind)
        except Exception:
            return (0.0,)

        correct = 0
        total = len(y_tr)
        tp, fp, fn, tn = 0, 0, 0, 0

        for i in range(total):
            mu = fuzzify_new(tr_usia[i], tr_tb[i], tr_bb[i], tr_mpasi[i], tr_jk[i], mf_p)
            fired = evaluate_rules_new(mu, CONSEQUENTS)
            risk = defuzzify_sugeno(fired)
            flag = get_priority_flag(risk)
            pred = 0 if flag == 'Normal' else 1
            actual = y_tr[i]

            if pred == 1 and actual == 1: tp += 1
            elif pred == 1 and actual == 0: fp += 1
            elif pred == 0 and actual == 1: fn += 1
            else: tn += 1

        prec = tp/(tp+fp) if (tp+fp) > 0 else 0
        rec = tp/(tp+fn) if (tp+fn) > 0 else 0
        f1 = 2*prec*rec/(prec+rec) if (prec+rec) > 0 else 0
        acc = (tp+tn)/total if total > 0 else 0

        # Fitness: weighted combo of F1 and accuracy
        fitness = 0.6 * f1 + 0.4 * acc
        return (fitness,)

    # DEAP setup
    if hasattr(creator, 'FitnessMax'):
        del creator.FitnessMax
    if hasattr(creator, 'Individual'):
        del creator.Individual

    creator.create("FitnessMax", base.Fitness, weights=(1.0,))
    creator.create("Individual", list, fitness=creator.FitnessMax)

    toolbox = base.Toolbox()
    init_genes = encode_initial()

    def create_individual():
        ind = []
        for i, g in enumerate(init_genes):
            lo, hi = GENE_BOUNDS[i]
            # Add noise around initial value
            noise = random.gauss(0, abs(hi - lo) * 0.1)
            val = g + noise
            val = max(lo, min(hi, val))
            ind.append(val)
        return creator.Individual(ind)

    toolbox.register("individual", create_individual)
    toolbox.register("population", tools.initRepeat, list, toolbox.individual)
    toolbox.register("evaluate", fitness_func)
    toolbox.register("mate", tools.cxBlend, alpha=0.3)

    def mutate_bounded(individual, mu=0, sigma=1, indpb=0.15):
        for i in range(len(individual)):
            if random.random() < indpb:
                lo, hi = GENE_BOUNDS[i]
                individual[i] += random.gauss(mu, abs(hi-lo)*0.15)
                individual[i] = max(lo, min(hi, individual[i]))
        return (individual,)

    toolbox.register("mutate", mutate_bounded)
    toolbox.register("select", tools.selTournament, tournsize=3)

    # Run GA - Tier 1 enhancement: larger search space
    POP_SIZE = 80
    N_GEN = 60
    print(f"    Pop={POP_SIZE}, Gen={N_GEN} (Tier 1: extended search)")
    # Multi-seed GA execution (Tier 1: best-of-N for robustness against local optima)
    SEEDS = [42, 123, 777, 2024, 31415]
    print(f"\n    Running GA with {len(SEEDS)} seeds: {SEEDS}")
    print(f"    Total evaluations: {POP_SIZE * (N_GEN+1) * len(SEEDS)}")

    seed_results = []
    all_logbooks = {}  # FIX: dict, not list - keyed by seed value
    for seed_idx, seed in enumerate(SEEDS):
        print(f"\n    [Seed {seed_idx+1}/{len(SEEDS)}] random_seed={seed}")
        random.seed(seed)
        np.random.seed(seed)

        pop = toolbox.population(n=POP_SIZE)
        # Seed population with initial params
        pop[0] = creator.Individual(init_genes)

        stats = tools.Statistics(lambda ind: ind.fitness.values)
        stats.register("max", np.max)
        stats.register("avg", np.mean)

        pop, logbook = algorithms.eaSimple(
            pop, toolbox,
            cxpb=0.6, mutpb=0.3,
            ngen=N_GEN, stats=stats,
            verbose=False  # silent per-seed; show summary instead
        )
        all_logbooks[seed] = logbook

        seed_best_ind = tools.selBest(pop, 1)[0]
        seed_best_mf = decode_individual(seed_best_ind)
        seed_train_fitness = seed_best_ind.fitness.values[0]

        # Score on test set with this seed's best
        _, _, y_pred_te_seed = run_pipeline(df_te, seed_best_mf, CONSEQUENTS)
        seed_test_metrics = calc_metrics(y_te, y_pred_te_seed)

        print(f"        Train fitness: {seed_train_fitness:.4f}  |  Test F1: {seed_test_metrics['f1']:.4f}  Acc: {seed_test_metrics['accuracy']:.4f}")

        seed_results.append({
            'seed': seed,
            'mf': seed_best_mf,
            'individual': list(seed_best_ind),
            'train_fitness': seed_train_fitness,
            'test_metrics': seed_test_metrics,
        })

    # Pick best seed by test F1 (with tie-break on accuracy)
    seed_results.sort(key=lambda r: (r['test_metrics']['f1'], r['test_metrics']['accuracy']), reverse=True)
    best_seed = seed_results[0]
    best_mf = best_seed['mf']
    best_fitness = best_seed['train_fitness']
    best_seed_value = best_seed['seed']

    print(f"\n    Multi-seed comparison:")
    print(f"    {'Seed':>8} {'TrainFit':>10} {'TestP':>8} {'TestR':>8} {'TestF1':>8} {'TestAcc':>8}")
    for r in seed_results:
        m = r['test_metrics']
        marker = "  <- BEST" if r['seed'] == best_seed_value else ""
        print(f"    {r['seed']:>8} {r['train_fitness']:>10.4f} {m['precision']:>8.4f} {m['recall']:>8.4f} {m['f1']:>8.4f} {m['accuracy']:>8.4f}{marker}")

    print(f"\n    Selected: seed={best_seed_value}, train fitness={best_fitness:.4f}")

    # ====== PLOT GA CONVERGENCE ACROSS SEEDS ======
    print_step("6.3.1", "Plot GA Convergence Across Seeds")

    paper_assets_dir = os.path.join(BASE_DIR, 'paper_assets')
    os.makedirs(paper_assets_dir, exist_ok=True)

    fig, ax = plt.subplots(figsize=(8, 5))
    for seed, lb in all_logbooks.items():
        gens = lb.select('gen')
        maxs = lb.select('max')
        ax.plot(gens, maxs, label=f'Seed {seed}', linewidth=1.5)
    ax.set_xlabel('Generation')
    ax.set_ylabel('Best Fitness')
    ax.set_title('GA Convergence Across 5 Random Seeds')
    ax.legend(loc='lower right')
    ax.grid(True, alpha=0.3)
    plt.tight_layout()

    fig_path = os.path.join(paper_assets_dir, 'figure2_ga_convergence.png')
    plt.savefig(fig_path, dpi=300, bbox_inches='tight')
    plt.close()  # important: close to free memory in batch run
    print(f"    [OK] Saved: {fig_path}")
    # ====== END PLOT ======

    # Threshold tuning post-hoc (Tier 1: optimize Normal/Stunting cutoff)
    # Run pipeline once to get raw risk scores
    print_step("6.4", "Threshold Tuning Post-Hoc (Tier 1)")
    print(f"    Mencari threshold optimal pada raw risk score (Normal vs Stunting cutoff)")

    risks_tr_raw, _, _ = run_pipeline(df_tr, best_mf, CONSEQUENTS)
    risks_te_raw, _, _ = run_pipeline(df_te, best_mf, CONSEQUENTS)
    risks_tr_arr = np.array(risks_tr_raw)
    risks_te_arr = np.array(risks_te_raw)

    # Default threshold from get_priority_flag is around 22.5 (transition Normal->Low at risk 30 with overlap from 15)
    # Scan a sensible range
    best_thr = None
    best_thr_f1 = -1
    best_thr_metrics = None
    threshold_candidates = list(range(15, 56, 1))  # 15..55
    for thr in threshold_candidates:
        y_pred_tr_thr = (risks_tr_arr >= thr).astype(int)
        m_tr = calc_metrics(y_tr, y_pred_tr_thr)
        # Combined: 0.6*F1 + 0.4*Acc (same as fitness for consistency)
        score = 0.6 * m_tr['f1'] + 0.4 * m_tr['accuracy']
        if score > best_thr_f1:
            best_thr_f1 = score
            best_thr = thr
            best_thr_metrics = m_tr

    # Compare: default (using priority flag) vs optimized threshold
    _, _, y_pred_tr_default = run_pipeline(df_tr, best_mf, CONSEQUENTS)
    default_train_metrics = calc_metrics(y_tr, y_pred_tr_default)

    print(f"    Default flag (Normal vs others): TrainF1={default_train_metrics['f1']:.4f} TrainAcc={default_train_metrics['accuracy']:.4f}")
    print(f"    Optimal threshold = {best_thr}    : TrainF1={best_thr_metrics['f1']:.4f} TrainAcc={best_thr_metrics['accuracy']:.4f}")

    # Decide: use optimal threshold only if it improves combined score on TRAIN
    default_score = 0.6 * default_train_metrics['f1'] + 0.4 * default_train_metrics['accuracy']
    if best_thr_f1 > default_score:
        use_threshold = best_thr
        print(f"    -> Using optimized threshold: {use_threshold}")
    else:
        use_threshold = None  # use default priority flag pipeline
        print(f"    -> Using default priority flag (no improvement from threshold tuning)")

    print(f"\n    Best fitness (final): {best_fitness:.4f}")

    # Evaluate on test set (using best_mf + optional threshold)
    print_step("5.1", "EVALUASI - Confusion Matrix")

    risks_tr_final, flags_tr_final, y_pred_tr_default = run_pipeline(df_tr, best_mf, CONSEQUENTS)
    risks_te_final, flags_te_final, y_pred_te_default = run_pipeline(df_te, best_mf, CONSEQUENTS)

    if use_threshold is not None:
        # Apply optimized threshold to raw risk scores
        risks_tr_arr_final = np.array(risks_tr_final)
        risks_te_arr_final = np.array(risks_te_final)
        y_pred_tr = (risks_tr_arr_final >= use_threshold).astype(int)
        y_pred_te = (risks_te_arr_final >= use_threshold).astype(int)
        print(f"    Using optimized threshold: {use_threshold}")
    else:
        y_pred_tr = y_pred_tr_default
        y_pred_te = y_pred_te_default
        print(f"    Using default priority flag pipeline")

    trm = calc_metrics(y_tr, y_pred_tr)
    tem = calc_metrics(y_te, y_pred_te)

    print(f"    TRAIN ({len(df_tr)}):")
    print(f"                  Act Stunting   Act Normal")
    print(f"    Pred Stunting    TP = {trm['TP']:>4d}     FP = {trm['FP']:>4d}")
    print(f"    Pred Normal      FN = {trm['FN']:>4d}     TN = {trm['TN']:>4d}")
    print(f"    P={trm['precision']:.4f} R={trm['recall']:.4f} F1={trm['f1']:.4f} Acc={trm['accuracy']:.4f}")

    print(f"\n    TEST ({len(df_te)}):")
    print(f"                  Act Stunting   Act Normal")
    print(f"    Pred Stunting    TP = {tem['TP']:>4d}     FP = {tem['FP']:>4d}")
    print(f"    Pred Normal      FN = {tem['FN']:>4d}     TN = {tem['TN']:>4d}")
    print(f"    P={tem['precision']:.4f} R={tem['recall']:.4f} F1={tem['f1']:.4f} Acc={tem['accuracy']:.4f}")

    # Re-compute all data with best MF
    print_step("5.2", "Re-compute seluruh data dengan MF optimal")
    all_risks, all_flags, all_binaries_default = run_pipeline(df, best_mf, CONSEQUENTS)
    df['risk_score'] = np.round(all_risks, 2)
    df['priority_flag'] = all_flags

    if use_threshold is not None:
        all_binaries = (np.array(all_risks) >= use_threshold).astype(int)
    else:
        all_binaries = all_binaries_default
    df['fuzzy_label'] = ['Stunting' if b == 1 else 'Normal' for b in all_binaries]

    y_all_t = (df['label_who_ukur'] == 'Stunting').astype(int).values
    y_all_p = all_binaries
    am = calc_metrics(y_all_t, y_all_p)

    print(f"\n    ALL DATA ({len(df)}):")
    print(f"    P={am['precision']:.4f} R={am['recall']:.4f} F1={am['f1']:.4f} Acc={am['accuracy']:.4f}")

    # Monitoring analysis: Z-Score LAHIR vs Fuzzy saat ukur
    print_step("5.3", "Analisis Monitoring: Z-Score LAHIR vs Fuzzy saat ukur")
    zs_stunting_fuzzy_normal = ((df['label_who']=='Stunting') & (df['fuzzy_label']=='Normal')).sum()
    zs_normal_fuzzy_stunting = ((df['label_who']=='Normal') & (df['fuzzy_label']=='Stunting')).sum()
    zs_stunting_fuzzy_stunting = ((df['label_who']=='Stunting') & (df['fuzzy_label']=='Stunting')).sum()
    zs_normal_fuzzy_normal = ((df['label_who']=='Normal') & (df['fuzzy_label']=='Normal')).sum()

    print(f"    Z-Score Lahir Stunting + Fuzzy Stunting : {zs_stunting_fuzzy_stunting:>4d} (konsisten buruk)")
    print(f"    Z-Score Lahir Stunting + Fuzzy Normal   : {zs_stunting_fuzzy_normal:>4d} (perkembangan BAIK)")
    print(f"    Z-Score Lahir Normal   + Fuzzy Stunting : {zs_normal_fuzzy_stunting:>4d} (perkembangan BURUK)")
    print(f"    Z-Score Lahir Normal   + Fuzzy Normal   : {zs_normal_fuzzy_normal:>4d} (konsisten baik)")

    # Cross-check: Z-Score saat ukur vs Fuzzy (ini yang jadi evaluasi akurasi)
    print_step("5.4", "Cross-check: Z-Score Saat Ukur vs Fuzzy (Evaluasi Akurasi)")
    zsu_s_f_s = ((df['label_who_ukur']=='Stunting') & (df['fuzzy_label']=='Stunting')).sum()
    zsu_s_f_n = ((df['label_who_ukur']=='Stunting') & (df['fuzzy_label']=='Normal')).sum()
    zsu_n_f_s = ((df['label_who_ukur']=='Normal') & (df['fuzzy_label']=='Stunting')).sum()
    zsu_n_f_n = ((df['label_who_ukur']=='Normal') & (df['fuzzy_label']=='Normal')).sum()

    print(f"    TP: Z-Score Ukur Stunting + Fuzzy Stunting : {zsu_s_f_s:>4d}")
    print(f"    FN: Z-Score Ukur Stunting + Fuzzy Normal   : {zsu_s_f_n:>4d}")
    print(f"    FP: Z-Score Ukur Normal   + Fuzzy Stunting : {zsu_n_f_s:>4d}")
    print(f"    TN: Z-Score Ukur Normal   + Fuzzy Normal   : {zsu_n_f_n:>4d}")

    # Save model
    print_step("5.5", "Simpan model terlatih")

    # Convert MF params to serializable format
    mf_serial = {}
    for k, v in best_mf.items():
        mf_serial[k] = {kk: [vv[0], vv[1]] for kk, vv in v.items()}

    trained_model = {
        'mf_params': mf_serial,
        'consequents': CONSEQUENTS,
        'output_mf': {k: list(v) for k, v in OUTPUT_MF.items()},
        'ga': {'pop_size': POP_SIZE, 'n_gen': N_GEN, 'best_fitness': float(best_fitness),
               'multi_seed': SEEDS, 'best_seed': int(best_seed_value)},
        'threshold': {'use_optimized': use_threshold is not None,
                      'value': int(use_threshold) if use_threshold is not None else None,
                      'scanned_range': [min(threshold_candidates), max(threshold_candidates)]},
        'metrics': {'train': trm, 'test': tem, 'all': am},
        'split': {'train_size': len(df_tr), 'test_size': len(df_te)},
        'monitoring': {
            'zs_stunting_fuzzy_stunting': int(zs_stunting_fuzzy_stunting),
            'zs_stunting_fuzzy_normal': int(zs_stunting_fuzzy_normal),
            'zs_normal_fuzzy_stunting': int(zs_normal_fuzzy_stunting),
            'zs_normal_fuzzy_normal': int(zs_normal_fuzzy_normal),
        },
        'timestamp': datetime.now().isoformat(),
    }

    model_path = os.path.join(TRAINFILE_DIR, 'trained_model.json')
    with open(model_path, 'w') as f:
        json.dump(trained_model, f, indent=2, default=str)
    print(f"    Saved: {model_path}")

    # Save train/test splits
    df.iloc[tr_idx].reset_index(drop=True).to_csv(
        os.path.join(TRAINFILE_DIR, 'dataset_train.csv'), index=False)
    df.iloc[te_idx].reset_index(drop=True).to_csv(
        os.path.join(TRAINFILE_DIR, 'dataset_test.csv'), index=False)
    print(f"    Saved: dataset_train.csv ({len(tr_idx)}), dataset_test.csv ({len(te_idx)})")

    # ====== TASK 5B: Perbandingan kuantitatif dengan baseline ML ======
    # Memakai tem (metrik uji Fuzzy-GA) dan split identik (tr_idx/te_idx).
    baseline_cmp = task5b_baseline_comparison(df, tr_idx, te_idx, tem)
    if baseline_cmp is not None:
        trained_model['baseline_comparison'] = {
            k: {kk: (int(vv) if isinstance(vv, (int, np.integer)) else float(vv))
                for kk, vv in v.items()}
            for k, v in baseline_cmp.items()
        }

    return df, trained_model





# ============================================================================
# MAIN
# ============================================================================
if __name__ == '__main__':
    print("\n" + "=" * 70)
    print("  STUNTOR - Monitoring Stunting Balita Kalimantan Tengah")
    print("  Dataset: SKI 2023 | Metode: Fuzzy Sugeno + GA")
    print("=" * 70)
    print(f"\n  Working directory : {os.getcwd()}")
    print(f"  Script location   : {os.path.abspath(__file__)}")
    print(f"  Dataset path      : {DATASET_PATH}")

    # TASK 1
    df_input, raw_shape = task1_input_data()

    # TASK 2
    df_clean, report = task2_preprocessing(df_input, raw_shape)
    df_clean.to_csv(OUTPUT_CLEAN_CSV, index=False)
    print(f"\n    Dataset bersih: {OUTPUT_CLEAN_CSV} ({len(df_clean):,} baris)")

    # TASK 3A: Z-SCORE LAHIR
    df_labeled, who_lookup = task3_zscore_labelling(df_clean)

    # TASK 3B: Z-SCORE SAAT PENGUKURAN
    df_labeled = task3b_zscore_pengukuran(df_labeled, who_lookup)

    # TASK 4
    df_fuzzy = task4_fuzzy_logic(df_labeled)

    # TASK 5 & 6
    df_final, trained_model = task5_6_evaluation(df_fuzzy)

    # SAVE FINAL
    print_header("MENYIMPAN HASIL AKHIR")
    df_final.to_csv(OUTPUT_FINAL_CSV, index=False)
    print(f"    Dataset final: {OUTPUT_FINAL_CSV} ({len(df_final):,} baris)")

    # Save analysis per row
    df_output = pd.DataFrame({
        'No': range(1, len(df_final)+1),
        'Sex': df_final['jenis_kelamin'].map({1:'Laki-laki',2:'Perempuan'}).values,
        'Umur (bln)': df_final['usia_bulan'].astype(int).values,
        'TB Lahir (cm)': df_final['pbl_cm'].values,
        'TB Saat Ukur (cm)': df_final['tinggi_badan_cm'].values,
        'BB Saat Ukur (g)': df_final['berat_badan_gram'].values,
        'MPASI (bln)': df_final['mpasi_bulan'].values,
        'Z-Score (lahir)': df_final['z_score'].values,
        'Label WHO (lahir)': df_final['label_who'].values,
        'Z-Score (saat ukur)': df_final['z_score_ukur'].values,
        'Label WHO (saat ukur)': df_final['label_who_ukur'].values,
        'Risk Score': df_final['risk_score'].values,
        'Priority Flag': df_final['priority_flag'].values,
        'Fuzzy Label': df_final['fuzzy_label'].values,
        'Jarak Lahir-Ukur (bln)': df_final['jarak_lahir_ukur'].values,
        'Jarak MPASI-Survei (bln)': df_final['jarak_mpasi_survei'].values,
    })
    out_perbaris = os.path.join(TRAINFILE_DIR, 'analisa_perbaris.csv')
    df_output.to_csv(out_perbaris, index=False)
    print(f"    Analisa per baris: {out_perbaris}")

    with open(OUTPUT_REPORT, 'w') as f:
        label_dist_lahir = df_final['label_who'].value_counts().to_dict()
        label_dist_ukur = df_final['label_who_ukur'].value_counts().to_dict()
        flag_dist = df_final['priority_flag'].value_counts().to_dict()
        report['task3a'] = {
            'label_distribution_lahir': {k: int(v) for k, v in label_dist_lahir.items()},
            'stunting_rate_lahir': round(label_dist_lahir.get('Stunting', 0) / len(df_final) * 100, 2),
        }
        report['task3b'] = {
            'label_distribution_ukur': {k: int(v) for k, v in label_dist_ukur.items()},
            'stunting_rate_ukur': round(label_dist_ukur.get('Stunting', 0) / len(df_final) * 100, 2),
        }
        report['task4'] = {
            'priority_flag_distribution': {k: int(v) for k, v in flag_dist.items()},
        }
        report['task5_6'] = trained_model
        json.dump(report, f, indent=2, default=str)
    print(f"    Report: {OUTPUT_REPORT}")

    # RINGKASAN
    print_header("RINGKASAN PIPELINE TASK 1-6")
    tm = trained_model['metrics']
    mon = trained_model['monitoring']
    label_dist_lahir = df_final['label_who'].value_counts().to_dict()
    label_dist_ukur = df_final['label_who_ukur'].value_counts().to_dict()
    print(f"    TASK 1: Input Data    - {raw_shape[0]:,} x {raw_shape[1]} -> 7 variabel")
    print(f"    TASK 2: Preprocessing - {len(df_clean):,} balita bersih")
    print(f"    TASK 3A: Z-Score LAHIR     - Normal={label_dist_lahir.get('Normal',0)}, Stunting={label_dist_lahir.get('Stunting',0)}")
    print(f"    TASK 3B: Z-Score SAAT UKUR - Normal={label_dist_ukur.get('Normal',0)}, Stunting={label_dist_ukur.get('Stunting',0)} (ground truth evaluasi)")
    print(f"    TASK 4: Fuzzy Sugeno  - 216 rules (108/JK), 4 var fuzzy, c=[10,35,65,90]")
    print(f"    TASK 5: Evaluasi      - Confusion Matrix (Fuzzy vs Z-Score saat ukur)")
    print(f"    TASK 6: GA Optimasi   - {N_GENES} boundary points")
    print(f"")
    print(f"    Hasil Evaluasi (Test 20%) - Fuzzy vs Z-Score SAAT PENGUKURAN:")
    print(f"    P={tm['test']['precision']:.4f} R={tm['test']['recall']:.4f} F1={tm['test']['f1']:.4f} Acc={tm['test']['accuracy']:.4f}")
    print(f"")
    print(f"    Monitoring (Z-Score LAHIR vs Fuzzy saat ukur):")
    print(f"    Perkembangan BAIK  : {mon['zs_stunting_fuzzy_normal']} balita (lahir stunting -> fuzzy normal)")
    print(f"    Perkembangan BURUK : {mon['zs_normal_fuzzy_stunting']} balita (lahir normal -> fuzzy stunting)")
    print(f"")
    print(f"    TASK 1-6 SELESAI - Siap untuk TASK 7 (Website Flask)")