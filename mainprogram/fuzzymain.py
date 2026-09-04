#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from flask import Flask, render_template, request, redirect, url_for, jsonify
import pandas as pd
import numpy as np
import json
import os
import csv
import sys
from datetime import datetime

# ============================================================================
# PATH CONFIGURATION
# ============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TRAIN_DIR = os.path.join(os.path.dirname(BASE_DIR), 'train')

MODEL_PATH      = os.path.join(TRAIN_DIR, 'trainfile', 'trained_model.json')
DATASET_FINAL   = os.path.join(TRAIN_DIR, 'trainfile', 'dataset_final.csv')
WHO_ZS_DIR      = os.path.join(TRAIN_DIR, 'zs')
TRAIN_SCRIPT    = os.path.join(TRAIN_DIR, 'fuzzytrainv2.py')

HISTORY_DIR     = os.path.join(BASE_DIR, 'history')
HISTORY_CSV     = os.path.join(HISTORY_DIR, 'history.csv')

os.makedirs(HISTORY_DIR, exist_ok=True)

WHO_TABLE_PATHS = {
    'boys_0_24':   os.path.join(WHO_ZS_DIR, 'boys02.xlsx'),
    'boys_24_60':  os.path.join(WHO_ZS_DIR, 'boys25.xlsx'),
    'girls_0_24':  os.path.join(WHO_ZS_DIR, 'girls02.xlsx'),
    'girls_24_60': os.path.join(WHO_ZS_DIR, 'girls25.xlsx'),
}


# ============================================================================
# INITIAL MF PARAMS (fallback jika trained_model.json belum ada)
# Identik dengan yang ada di fuzzytrainv2.py
# ============================================================================
DEFAULT_MF_PARAMS = {
    'usia': {
        'Muda':   ('trap_closed', [0, 0, 6, 18]),
        'Sedang': ('tri',  [12, 30, 48]),
        'Tua':    ('trap_closed', [42, 52, 59, 59]),
    },
    'tb_l': {
        'Pendek': ('trap', [38, 38, 55, 70]),
        'Normal': ('trap', [60, 70, 100, 110]),
        'Tinggi': ('trap', [100, 110, 130, 130]),
    },
    'tb_p': {
        'Pendek': ('trap', [38, 38, 53, 68]),
        'Normal': ('trap', [58, 68, 98, 108]),
        'Tinggi': ('trap', [98, 108, 130, 130]),
    },
    'bb_l': {
        'Kurang':   ('trap', [1500, 1500, 5000, 8000]),
        'Normal':   ('trap', [6000, 8000, 15000, 18000]),
        'Lebih':    ('tri',  [15000, 18000, 22000]),
        'Obesitas': ('trap', [20000, 23000, 30000, 30000]),
    },
    'bb_p': {
        'Kurang':   ('trap', [1500, 1500, 4500, 7500]),
        'Normal':   ('trap', [5500, 7500, 14000, 17000]),
        'Lebih':    ('tri',  [14000, 17000, 21000]),
        'Obesitas': ('trap', [19000, 22000, 28000, 28000]),
    },
    'mpasi': {
        'TerlaCepat': ('trap', [0, 0, 3, 5]),
        'Normal':     ('trap', [4, 5.5, 6.5, 8]),
        'Lambat':     ('trap', [7, 8, 12, 12]),
    },
}

CONSEQUENTS = {'high': 90, 'medium': 65, 'low': 35, 'normal': 10}


# ============================================================================
# FUZZY LOGIC (identik dengan fuzzytrainv2.py)
# ============================================================================
def trapezoid(x, a, b, c, d):
    # Fungsi keanggotaan trapesium dengan dukungan bentuk BAHU (shoulder)
    # untuk himpunan di ujung domain. Konsisten dengan definisi pada pipeline
    # pelatihan (fuzzytrainv2.py):
    #   - Bila a == b (himpunan terendah): sisi kiri berbentuk bahu, sehingga
    #     seluruh x <= b memiliki derajat keanggotaan penuh (mu = 1). Ini
    #     menjamin nilai ekstrem rendah (mis. MP-ASI 0 / belum diberi) tetap
    #     memiliki keanggotaan penuh.
    #   - Bila c == d (himpunan tertinggi): sisi kanan berbentuk bahu, sehingga
    #     seluruh x >= c memiliki keanggotaan penuh (mis. MP-ASI >= 12 bulan).
    # Himpunan tengah (a < b dan c < d) tetap trapesium biasa.
    left_shoulder = (a == b)
    right_shoulder = (c == d)
    if left_shoulder and x <= b:
        return 1.0
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
    return trapezoid(x, a, b, b, c)

def trapezoid_closed(x, a, b, c, d):
    # Trapesium TERTUTUP murni (tanpa bahu) - dipakai untuk variabel Usia agar
    # batas 0-59 bulan tegas. Konsisten dengan fuzzytrainv2.py.
    if x <= a or x >= d:
        return 0.0
    elif a < x < b:
        return (x - a) / (b - a)
    elif b <= x <= c:
        return 1.0
    elif c < x < d:
        return (d - x) / (d - c)
    return 0.0

def eval_mf(x, mf_type, params):
    if mf_type == 'trap':
        return trapezoid(x, *params)
    elif mf_type == 'trap_closed':
        return trapezoid_closed(x, *params)
    elif mf_type == 'tri':
        return triangle(x, *params)
    return 0.0


# 108 Rules eksplisit (identik dengan fuzzytrainv2.py)
RULES_108 = [
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


# ============================================================================
# MODEL & WHO LOADING
# ============================================================================
def load_trained_model():
    if not os.path.exists(MODEL_PATH):
        print(f"[!] trained_model.json tidak ditemukan: {MODEL_PATH}")
        print(f"[!] Menggunakan DEFAULT MF params (belum dioptimasi GA)")
        return None
    with open(MODEL_PATH, 'r') as f:
        model = json.load(f)
    print(f"[OK] Model terlatih dimuat: {MODEL_PATH}")
    ts = model.get('timestamp', 'unknown')
    fitness = model.get('ga', {}).get('best_fitness', 'N/A')
    print(f"     Timestamp: {ts}")
    print(f"     GA fitness: {fitness}")
    return model


def parse_mf_params(model):
    """Parse mf_params dari trained_model.json.

    Format di JSON (dari fuzzytrainv2.py serialisasi):
      mf_params: {
        "usia": {"Muda": ["trap", [0, 0, 6, 18]], ...},
        "tb_l": {"Pendek": ["trap", [38, 38, 55, 70]], ...},
        ...
      }
    Setiap entry: [mf_type_string, [param_values]]
    Perlu dikonversi ke: (mf_type_string, [param_values]) tuple
    """
    if model is None:
        return DEFAULT_MF_PARAMS

    mf_raw = model.get('mf_params', None)
    if mf_raw is None:
        return DEFAULT_MF_PARAMS

    mf_params = {}
    for var_key, categories in mf_raw.items():
        mf_params[var_key] = {}
        for cat_name, val in categories.items():
            # val = ["trap", [a, b, c, d]] atau ["tri", [a, b, c]]
            mf_type = val[0]         # string: "trap" atau "tri"
            params = val[1]          # list of numbers
            mf_params[var_key][cat_name] = (mf_type, params)
    return mf_params


def load_who_tables():
    who_lookup = {}
    for table_key, fpath in WHO_TABLE_PATHS.items():
        if not os.path.exists(fpath):
            print(f"[!] WHO table tidak ditemukan: {fpath}")
            continue
        jk = 1 if 'boys' in table_key else 2
        df_who = pd.read_excel(fpath)
        df_who.columns = df_who.columns.str.strip()
        for _, row in df_who.iterrows():
            month = int(row['Month'])
            who_lookup[(jk, month)] = {
                'L': float(row['L']), 'M': float(row['M']), 'S': float(row['S'])
            }
    print(f"[OK] WHO LMS tables dimuat: {len(who_lookup)} entries")
    return who_lookup


# ============================================================================
# INFERENCE
# ============================================================================
def calc_zscore(x, L, M, S):
    if L != 0:
        return round(((x / M) ** L - 1) / (L * S), 3)
    else:
        return round(np.log(x / M) / S, 3)


def get_zscore_lahir(jk, pbl_cm, who_lookup):
    key = (jk, 0)  # usia = 0 bulan (saat lahir)
    if key not in who_lookup:
        return None, None
    p = who_lookup[key]
    z = calc_zscore(pbl_cm, p['L'], p['M'], p['S'])
    label = 'Stunting' if z < -2.0 else 'Normal'
    return z, label


def get_zscore_ukur(jk, usia, tb_cm, who_lookup):
    key = (jk, int(usia))  # usia aktual saat pengukuran
    if key not in who_lookup:
        return None, None
    p = who_lookup[key]
    z = calc_zscore(tb_cm, p['L'], p['M'], p['S'])
    label = 'Stunting' if z < -2.0 else 'Normal'
    return z, label


def fuzzify(usia, tb, bb, mpasi, jk, mf_params):
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


def evaluate_rules(mu):
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
            fired.append((w, CONSEQUENTS[risk_class]))
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
    if risk_score <= 0:
        mu_n = 1.0
    mu_l = triangle(risk_score, 15, 30, 60)
    mu_m = triangle(risk_score, 45, 60, 75)
    mu_h = trapezoid(risk_score, 75, 90, 100, 101)
    if risk_score >= 100:
        mu_h = 1.0
    mu_out = {'Normal': mu_n, 'STUNTING LOW': mu_l,
              'STUNTING MEDIUM': mu_m, 'STUNTING HIGH': mu_h}
    return max(mu_out, key=mu_out.get)


def predict(usia, jk, tb_cm, bb_gram, pbl_cm, mpasi_bulan, mf_params, who_lookup):
    # Z-Score Lahir
    z_lahir, label_lahir = get_zscore_lahir(jk, pbl_cm, who_lookup)

    # Z-Score Saat Ukur
    z_ukur, label_ukur = get_zscore_ukur(jk, usia, tb_cm, who_lookup)

    # Fuzzy Inference
    # Deteksi kasus MP-ASI kosong / tidak tahu (nilai sentinel -1). Pada kasus
    # ini nilai MP-ASI tidak diketahui, sehingga penilaian mengikuti mekanisme
    # breakdown dua kondisi berdasarkan usia (lihat penjelasan Gambar 3.9):
    #   - usia < 6 bulan  -> wajar belum diberi MP-ASI (ASI eksklusif ideal) ->
    #                        label akhir dipaksa Normal.
    #   - usia >= 6 bulan -> MP-ASI seharusnya sudah diperkenalkan; ketiadaan
    #                        MP-ASI menjadi indikator risiko -> label akhir
    #                        dipaksa Stunting. Tingkat keparahan (LOW/MEDIUM/
    #                        HIGH) tetap ditentukan oleh mesin fuzzy melalui
    #                        risk_score, bukan dipaksa.
    mpasi_kosong = (mpasi_bulan is not None and float(mpasi_bulan) < 0)

    # Untuk inferensi fuzzy, MP-ASI kosong diperlakukan sebagai 0 (belum diberi)
    # agar mesin fuzzy tetap menghasilkan risk_score untuk gradasi tingkat.
    mpasi_infer = 0.0 if mpasi_kosong else mpasi_bulan

    mu = fuzzify(usia, tb_cm, bb_gram, mpasi_infer, jk, mf_params)
    fired = evaluate_rules(mu)
    risk_score = defuzzify_sugeno(fired)
    priority_flag = get_priority_flag(risk_score)
    fuzzy_label = 'Normal' if priority_flag == 'Normal' else 'Stunting'

    # Override breakdown dua kondisi khusus MP-ASI kosong
    if mpasi_kosong:
        if usia < 6:
            # Usia < 6 bulan: belum diberi MP-ASI adalah wajar -> Normal
            priority_flag = 'Normal'
            fuzzy_label = 'Normal'
        else:
            # Usia >= 6 bulan: ketiadaan MP-ASI menjadi indikator risiko ->
            # Stunting. Tingkat (LOW/MEDIUM/HIGH) diambil dari fuzzy; bila fuzzy
            # menilai Normal, tingkat minimal dinaikkan ke STUNTING LOW.
            fuzzy_label = 'Stunting'
            if priority_flag == 'Normal':
                priority_flag = 'STUNTING LOW'

    # Klasifikasi Risiko berdasarkan Z-Score saat ukur
    if label_ukur == 'Stunting' and z_ukur is not None and z_ukur < -3.0:
        klasifikasi_risiko = 'Stunting Berat'
    elif label_ukur == 'Stunting':
        klasifikasi_risiko = 'Stunting'
    else:
        klasifikasi_risiko = 'Normal'

    # Monitoring
    if label_lahir == 'Stunting' and fuzzy_label == 'Stunting':
        monitoring = 'Konsisten Buruk - Lahir stunting, saat ini masih berisiko'
    elif label_lahir == 'Stunting' and fuzzy_label == 'Normal':
        monitoring = 'Perkembangan BAIK - Lahir stunting, kondisi saat ini membaik'
    elif label_lahir == 'Normal' and fuzzy_label == 'Stunting':
        monitoring = 'Perkembangan BURUK - Lahir normal, saat ini berisiko stunting'
    else:
        monitoring = 'Konsisten Baik - Lahir normal, kondisi saat ini tetap baik'

    return {
        'z_score_lahir': z_lahir,
        'label_lahir': label_lahir,
        'z_score_ukur': z_ukur,
        'label_ukur': label_ukur,
        'klasifikasi_risiko': klasifikasi_risiko,
        'risk_score': risk_score,
        'priority_flag': priority_flag,
        'fuzzy_label': fuzzy_label,
        'monitoring': monitoring,
        'rules_fired': len(fired),
    }


# ============================================================================
# HISTORY & AUTO-RETRAIN
# ============================================================================
HISTORY_HEADER = [
    'timestamp', 'nama', 'usia_bulan', 'jenis_kelamin', 'tinggi_badan_cm',
    'berat_badan_gram', 'pbl_cm', 'mpasi_bulan',
    'z_score_lahir', 'label_lahir', 'z_score_ukur', 'label_ukur',
    'risk_score', 'priority_flag', 'fuzzy_label', 'monitoring'
]


def save_to_history(input_data, result):
    file_exists = os.path.exists(HISTORY_CSV)
    with open(HISTORY_CSV, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(HISTORY_HEADER)
        writer.writerow([
            datetime.now().isoformat(),
            input_data.get('nama', ''),
            input_data['usia'], input_data['jk'], input_data['tb'],
            input_data['bb'], input_data['pbl'], input_data['mpasi'],
            result['z_score_lahir'], result['label_lahir'],
            result['z_score_ukur'], result['label_ukur'],
            result['risk_score'], result['priority_flag'],
            result['fuzzy_label'], result['monitoring'],
        ])


def _read_history_df():
    """Baca history.csv secara tahan-galat TANPA mengubah file.

    File history bisa memiliki baris dengan jumlah kolom berbeda:
      - 15 kolom  -> skema lama (tanpa 'nama')
      - 16 kolom  -> skema baru (dengan 'nama' di posisi ke-2)
    Setiap baris dipetakan ke skema baru berdasarkan JUMLAH kolomnya, sehingga
    tidak ada data yang bergeser. Baris tanpa nama diberi nilai kosong.
    """
    if not os.path.exists(HISTORY_CSV):
        return pd.DataFrame(columns=HISTORY_HEADER)

    with open(HISTORY_CSV, 'r', newline='', encoding='utf-8') as f:
        raw = list(csv.reader(f))
    if not raw:
        return pd.DataFrame(columns=HISTORY_HEADER)

    # Kolom skema lama (15, tanpa 'nama') dan baru (16, dengan 'nama')
    OLD_COLS = [c for c in HISTORY_HEADER if c != 'nama']  # 15 kolom

    # Deteksi & lewati baris header bila ada (baris yang memuat 'timestamp')
    data_rows = []
    for r in raw:
        if not r:
            continue
        if r[0].strip().lower() == 'timestamp':
            continue  # baris header, lewati
        data_rows.append(r)

    records = []
    for r in data_rows:
        if len(r) >= len(HISTORY_HEADER):
            # 16+ kolom -> skema baru (ambil 16 pertama)
            rec = dict(zip(HISTORY_HEADER, r[:len(HISTORY_HEADER)]))
        elif len(r) == len(OLD_COLS):
            # 15 kolom -> skema lama, sisipkan nama kosong
            rec = dict(zip(OLD_COLS, r))
            rec['nama'] = ''
        else:
            # jumlah kolom tak terduga: petakan seadanya lalu lengkapi
            rec = dict(zip(HISTORY_HEADER, r))
            for c in HISTORY_HEADER:
                rec.setdefault(c, '')
            if 'nama' not in rec or rec.get('nama') is None:
                rec['nama'] = ''
        records.append(rec)

    df = pd.DataFrame(records, columns=HISTORY_HEADER)
    if 'nama' not in df.columns:
        df['nama'] = ''
    df['nama'] = df['nama'].fillna('')
    return df


def get_history():
    df = _read_history_df()
    if df.empty:
        return []
    return df.to_dict('records')


def auto_retrain():
    """Auto-retrain: gabungkan dataset_final.csv + history.csv,
    hitung Z-score saat ukur untuk history sebagai ground truth,
    lalu jalankan GA optimasi boundary points pada dataset gabungan.

    Cara kerja:
    1. Baca dataset_final.csv (data training original)
    2. Baca history.csv (data input user dari website)
    3. Hitung Z-score saat ukur untuk history -> label ground truth
    4. Konversi history ke format dataset_final
    5. Gabungkan keduanya
    6. Jalankan GA optimasi pada dataset gabungan
    7. Simpan trained_model.json baru
    """
    print("\n[AUTO-RETRAIN] Memulai proses retrain...")

    if not os.path.exists(DATASET_FINAL):
        print(f"[!] dataset_final.csv tidak ditemukan: {DATASET_FINAL}")
        return False, "Dataset final tidak ditemukan"

    if not os.path.exists(HISTORY_CSV):
        print(f"[!] Tidak ada history untuk ditambahkan")
        return False, "Tidak ada history"

    # 1. Baca dataset final original
    df_orig = pd.read_csv(DATASET_FINAL)
    print(f"[AUTO-RETRAIN] Dataset original: {len(df_orig)} baris")

    # 2. Baca history (tahan-galat terhadap skema kolom lama)
    df_hist = _read_history_df()
    if len(df_hist) < 5:
        msg = f"History terlalu sedikit ({len(df_hist)} records, minimal 5)"
        print(f"[!] {msg}")
        return False, msg
    print(f"[AUTO-RETRAIN] History: {len(df_hist)} baris")

    # 3. Konversi history ke format dataset_final
    # Kolom yang dibutuhkan: usia_bulan, jenis_kelamin, tinggi_badan_cm,
    #   berat_badan_gram, pbl_cm, mpasi_bulan, label_who_ukur
    # History sudah punya z_score_ukur dan label_ukur dari saat predict
    hist_rows = []
    for _, row in df_hist.iterrows():
        try:
            usia = int(row['usia_bulan'])
            jk = int(row['jenis_kelamin'])
            tb = float(row['tinggi_badan_cm'])
            bb = float(row['berat_badan_gram'])
            pbl = float(row['pbl_cm'])
            mpasi = float(row['mpasi_bulan'])
            label_ukur = str(row.get('label_ukur', ''))
            nama = str(row.get('nama', '') or '')

            # Validasi (batasan rentang tb/bb/pbl tidak diberlakukan agar
            # konsisten dengan input bebas pada form; MF bahu menangani nilai
            # ekstrem). Usia tetap dibatasi di bawah 60 bulan.
            if usia < 0 or usia > 59: continue
            if jk not in [1, 2]: continue
            if mpasi < 0 or mpasi > 12: continue
            if label_ukur not in ['Normal', 'Stunting']: continue

            hist_rows.append({
                'nama': nama,
                'usia_bulan': usia,
                'jenis_kelamin': jk,
                'tinggi_badan_cm': tb,
                'berat_badan_gram': bb,
                'pbl_cm': pbl,
                'mpasi_bulan': mpasi,
                'label_who_ukur': label_ukur,
            })
        except (ValueError, KeyError):
            continue

    if len(hist_rows) == 0:
        return False, "Tidak ada history valid untuk retrain"

    df_hist_clean = pd.DataFrame(hist_rows)
    print(f"[AUTO-RETRAIN] History valid: {len(df_hist_clean)} baris")

    # 4. Pastikan dataset original punya kolom yang sama
    needed_cols = ['usia_bulan', 'jenis_kelamin', 'tinggi_badan_cm',
                   'berat_badan_gram', 'pbl_cm', 'mpasi_bulan', 'label_who_ukur']

    # Cek apakah semua kolom ada di df_orig
    missing_cols = [c for c in needed_cols if c not in df_orig.columns]
    if missing_cols:
        print(f"[!] Kolom tidak ditemukan di dataset_final: {missing_cols}")
        # Coba pakai label_who_ukur, kalau tidak ada coba label_who
        if 'label_who_ukur' not in df_orig.columns and 'label_who' in df_orig.columns:
            print("[AUTO-RETRAIN] Menggunakan label_who sebagai fallback")

    df_orig_subset = df_orig[
        [c for c in needed_cols if c in df_orig.columns]
    ].copy()

    # Jika label_who_ukur tidak ada, gunakan label alternatif
    if 'label_who_ukur' not in df_orig_subset.columns:
        if 'label_who' in df_orig.columns:
            df_orig_subset['label_who_ukur'] = df_orig['label_who']
        else:
            return False, "Dataset tidak memiliki kolom label yang valid"

    # 5. Gabungkan
    df_combined = pd.concat([df_orig_subset, df_hist_clean], ignore_index=True)
    df_combined = df_combined.dropna(subset=needed_cols).reset_index(drop=True)
    print(f"[AUTO-RETRAIN] Dataset gabungan: {len(df_combined)} baris "
          f"(original={len(df_orig_subset)}, history={len(df_hist_clean)})")

    # 6. Jalankan GA optimasi pada dataset gabungan
    print(f"[AUTO-RETRAIN] Menjalankan GA optimasi...")

    y_all = (df_combined['label_who_ukur'] == 'Stunting').astype(int).values
    n_total = len(df_combined)

    # Stratified split 80/20
    np.random.seed(42)
    idx_n = np.where(y_all == 0)[0].copy()
    idx_s = np.where(y_all == 1)[0].copy()
    np.random.shuffle(idx_n)
    np.random.shuffle(idx_s)
    n_tr_n = int(len(idx_n) * 0.8)
    n_tr_s = int(len(idx_s) * 0.8)
    tr_idx = np.concatenate([idx_n[:n_tr_n], idx_s[:n_tr_s]])

    df_tr = df_combined.iloc[tr_idx].reset_index(drop=True)
    y_tr = (df_tr['label_who_ukur'] == 'Stunting').astype(int).values

    tr_usia = df_tr['usia_bulan'].values.astype(float)
    tr_tb = df_tr['tinggi_badan_cm'].values.astype(float)
    tr_bb = df_tr['berat_badan_gram'].values.astype(float)
    tr_mpasi = df_tr['mpasi_bulan'].values.astype(float)
    tr_jk = df_tr['jenis_kelamin'].values.astype(int)

    # GA fitness function
    from deap import base as deap_base, creator as deap_creator, tools as deap_tools, algorithms as deap_algo
    import random

    # Gene bounds (sama persis dengan fuzzytrainv2.py)
    gene_bounds = []
    gene_bounds += [(0,0),(0,0),(3,12),(12,24)]
    gene_bounds += [(6,18),(20,36),(36,54)]
    gene_bounds += [(30,48),(42,56),(59,59),(59,59)]
    gene_bounds += [(35,42),(35,42),(45,60),(60,80)]
    gene_bounds += [(50,70),(60,80),(90,105),(100,115)]
    gene_bounds += [(90,105),(100,115),(125,135),(125,135)]
    gene_bounds += [(35,42),(35,42),(43,58),(58,75)]
    gene_bounds += [(48,65),(58,75),(88,103),(98,113)]
    gene_bounds += [(88,103),(98,113),(125,135),(125,135)]
    gene_bounds += [(1000,2000),(1000,2000),(3000,6000),(6000,10000)]
    gene_bounds += [(4000,8000),(6000,10000),(13000,17000),(16000,20000)]
    gene_bounds += [(13000,17000),(16000,20000),(20000,24000)]
    gene_bounds += [(18000,22000),(21000,25000),(28000,32000),(28000,32000)]
    gene_bounds += [(1000,2000),(1000,2000),(2500,5500),(5500,9000)]
    gene_bounds += [(3500,7000),(5500,9000),(12000,16000),(15000,19000)]
    gene_bounds += [(12000,16000),(15000,19000),(19000,23000)]
    gene_bounds += [(17000,21000),(20000,24000),(26000,30000),(26000,30000)]
    gene_bounds += [(0,0),(0,0),(2,4),(3,6)]
    gene_bounds += [(3,5),(4,6),(5.5,7.5),(6,9)]
    gene_bounds += [(5.5,7.5),(6,9),(12,12),(12,12)]

    def decode_individual_local(individual):
        i = 0
        usia = {'Muda':('trap',list(individual[i:i+4])),'Sedang':('tri',list(individual[i+4:i+7])),'Tua':('trap',list(individual[i+7:i+11]))}
        i += 11
        tb_l = {'Pendek':('trap',list(individual[i:i+4])),'Normal':('trap',list(individual[i+4:i+8])),'Tinggi':('trap',list(individual[i+8:i+12]))}
        i += 12
        tb_p = {'Pendek':('trap',list(individual[i:i+4])),'Normal':('trap',list(individual[i+4:i+8])),'Tinggi':('trap',list(individual[i+8:i+12]))}
        i += 12
        bb_l = {'Kurang':('trap',list(individual[i:i+4])),'Normal':('trap',list(individual[i+4:i+8])),'Lebih':('tri',list(individual[i+8:i+11])),'Obesitas':('trap',list(individual[i+11:i+15]))}
        i += 15
        bb_p = {'Kurang':('trap',list(individual[i:i+4])),'Normal':('trap',list(individual[i+4:i+8])),'Lebih':('tri',list(individual[i+8:i+11])),'Obesitas':('trap',list(individual[i+11:i+15]))}
        i += 15
        mpasi_mf = {'TerlaCepat':('trap',list(individual[i:i+4])),'Normal':('trap',list(individual[i+4:i+8])),'Lambat':('trap',list(individual[i+8:i+12]))}
        return {'usia':usia,'tb_l':tb_l,'tb_p':tb_p,'bb_l':bb_l,'bb_p':bb_p,'mpasi':mpasi_mf}

    # Start from current best params
    init_genes = []
    for var_key in ['usia', 'tb_l', 'tb_p', 'bb_l', 'bb_p', 'mpasi']:
        for cat_name in mf_params[var_key]:
            init_genes += list(mf_params[var_key][cat_name][1])

    n_genes = len(init_genes)

    def fitness_func(individual):
        ind = list(individual)
        for i in range(len(ind)):
            lo, hi = gene_bounds[i]
            ind[i] = max(lo, min(hi, ind[i]))
        try:
            mfp = decode_individual_local(ind)
        except Exception:
            return (0.0,)

        tp, fp, fn, tn = 0, 0, 0, 0
        for i in range(len(y_tr)):
            mu = fuzzify(tr_usia[i], tr_tb[i], tr_bb[i], tr_mpasi[i], tr_jk[i], mfp)
            fired = evaluate_rules(mu)
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
        acc = (tp+tn)/len(y_tr) if len(y_tr) > 0 else 0
        return (0.6 * f1 + 0.4 * acc,)

    # DEAP setup
    if hasattr(deap_creator, 'FitnessMaxRetrain'):
        del deap_creator.FitnessMaxRetrain
    if hasattr(deap_creator, 'IndividualRetrain'):
        del deap_creator.IndividualRetrain

    deap_creator.create("FitnessMaxRetrain", deap_base.Fitness, weights=(1.0,))
    deap_creator.create("IndividualRetrain", list, fitness=deap_creator.FitnessMaxRetrain)

    toolbox = deap_base.Toolbox()

    def create_ind():
        ind = []
        for i, g in enumerate(init_genes):
            lo, hi = gene_bounds[i]
            noise = random.gauss(0, abs(hi - lo) * 0.1)
            val = max(lo, min(hi, g + noise))
            ind.append(val)
        return deap_creator.IndividualRetrain(ind)

    toolbox.register("individual", create_ind)
    toolbox.register("population", deap_tools.initRepeat, list, toolbox.individual)
    toolbox.register("evaluate", fitness_func)
    toolbox.register("mate", deap_tools.cxBlend, alpha=0.3)

    def mutate_bounded(individual):
        for i in range(len(individual)):
            if random.random() < 0.15:
                lo, hi = gene_bounds[i]
                individual[i] += random.gauss(0, abs(hi-lo)*0.15)
                individual[i] = max(lo, min(hi, individual[i]))
        return (individual,)

    toolbox.register("mutate", mutate_bounded)
    toolbox.register("select", deap_tools.selTournament, tournsize=3)

    POP_SIZE = 30
    N_GEN = 20
    print(f"[AUTO-RETRAIN] GA: Pop={POP_SIZE}, Gen={N_GEN}, Genes={n_genes}")

    pop = toolbox.population(n=POP_SIZE)
    pop[0] = deap_creator.IndividualRetrain(init_genes)

    stats = deap_tools.Statistics(lambda ind: ind.fitness.values)
    stats.register("max", np.max)
    stats.register("avg", np.mean)

    pop, logbook = deap_algo.eaSimple(
        pop, toolbox, cxpb=0.6, mutpb=0.3,
        ngen=N_GEN, stats=stats, verbose=True
    )

    best_ind = deap_tools.selBest(pop, 1)[0]
    best_mf = decode_individual_local(best_ind)
    best_fitness = best_ind.fitness.values[0]

    print(f"[AUTO-RETRAIN] Best fitness: {best_fitness:.4f}")

    # 7. Evaluasi pada seluruh data gabungan
    tp, fp, fn, tn = 0, 0, 0, 0
    all_usia = df_combined['usia_bulan'].values.astype(float)
    all_tb = df_combined['tinggi_badan_cm'].values.astype(float)
    all_bb = df_combined['berat_badan_gram'].values.astype(float)
    all_mpasi = df_combined['mpasi_bulan'].values.astype(float)
    all_jk = df_combined['jenis_kelamin'].values.astype(int)
    y_all_eval = (df_combined['label_who_ukur'] == 'Stunting').astype(int).values

    for i in range(len(df_combined)):
        mu = fuzzify(all_usia[i], all_tb[i], all_bb[i], all_mpasi[i], all_jk[i], best_mf)
        fired = evaluate_rules(mu)
        risk = defuzzify_sugeno(fired)
        flag = get_priority_flag(risk)
        pred = 0 if flag == 'Normal' else 1
        actual = y_all_eval[i]
        if pred == 1 and actual == 1: tp += 1
        elif pred == 1 and actual == 0: fp += 1
        elif pred == 0 and actual == 1: fn += 1
        else: tn += 1

    prec = tp/(tp+fp) if (tp+fp) > 0 else 0
    rec = tp/(tp+fn) if (tp+fn) > 0 else 0
    f1 = 2*prec*rec/(prec+rec) if (prec+rec) > 0 else 0
    acc = (tp+tn)/len(df_combined) if len(df_combined) > 0 else 0

    print(f"[AUTO-RETRAIN] Evaluasi: P={prec:.4f} R={rec:.4f} F1={f1:.4f} Acc={acc:.4f}")

    # 8. Simpan model baru
    mf_serial = {}
    for k, v in best_mf.items():
        mf_serial[k] = {kk: [vv[0], vv[1]] for kk, vv in v.items()}

    new_model = {
        'mf_params': mf_serial,
        'consequents': CONSEQUENTS,
        'ga': {'pop_size': POP_SIZE, 'n_gen': N_GEN, 'best_fitness': float(best_fitness)},
        'metrics': {
            'all': {'precision': round(prec,4), 'recall': round(rec,4),
                    'f1': round(f1,4), 'accuracy': round(acc,4),
                    'TP': tp, 'TN': tn, 'FP': fp, 'FN': fn},
        },
        'retrain_info': {
            'original_rows': len(df_orig_subset),
            'history_rows': len(df_hist_clean),
            'combined_rows': len(df_combined),
        },
        'timestamp': datetime.now().isoformat(),
    }

    with open(MODEL_PATH, 'w') as f:
        json.dump(new_model, f, indent=2, default=str)

    msg = (f"Retrain berhasil: {len(df_combined)} data "
           f"(+{len(df_hist_clean)} history), "
           f"F1={f1:.4f}, Acc={acc:.4f}")
    print(f"[AUTO-RETRAIN] {msg}")
    print(f"[AUTO-RETRAIN] Model disimpan: {MODEL_PATH}")

    return True, msg


# ============================================================================
# FLASK APP
# ============================================================================
app = Flask(__name__)

# Load model dan WHO tables saat startup
trained_model = load_trained_model()
mf_params = parse_mf_params(trained_model)
who_lookup = load_who_tables()

# Pilihan MP-ASI mengikuti kategori resmi Buku Kode SKI 2023 (kolom I43:
# "Pada saat [NAMA] umur berapa Ibu mulai mengenalkan makanan/minuman selain
# ASI?"). Kategori di bawah 1 bulan pada kuesioner terbagi menjadi dua indikator
# berbasis HARI (0-7 hari dan 8-29 hari). Label ditampilkan dalam satuan hari,
# namun nilai numerik yang dikirim ke mesin fuzzy tetap dalam satuan BULAN
# (0,25 dan 0,5) agar perhitungan fungsi keanggotaan tidak berubah.
# Format tuple: (label_tampilan, nilai_bulan)
MPASI_OPTIONS = [
    ('Kosong / Tidak tahu', -1),
    ('Belum diberi MP-ASI', 0),
    ('0-7 hari', 0.25), ('8-29 hari', 0.5),
    ('1 bulan', 1), ('2 bulan', 2), ('3 bulan', 3), ('4 bulan', 4),
    ('5 bulan', 5), ('6 bulan', 6), ('7 bulan', 7), ('8 bulan', 8),
    ('9 bulan', 9), ('10 bulan', 10), ('11 bulan', 11), ('≥12 bulan', 12),
]


@app.route('/')
def landing():
    return render_template('landing.html')


@app.route('/form')
def index():
    return render_template('index.html', mpasi_options=MPASI_OPTIONS)


@app.route('/predict', methods=['POST'])
def predict_route():
    try:
        nama = request.form.get('nama', '').strip()
        usia = int(request.form['usia'])
        jk = int(request.form['jk'])
        tb = float(request.form['tb'])
        bb = float(request.form['bb'])
        pbl = float(request.form['pbl'])
        mpasi = float(request.form['mpasi'])

        errors = []
        if usia < 0 or usia > 59:
            errors.append('Usia harus di bawah 60 bulan (0-59 bulan)')
        if jk not in [1, 2]:
            errors.append('Jenis kelamin tidak valid')
        if mpasi != -1 and (mpasi < 0 or mpasi > 12):
            errors.append('MP-ASI harus 0-12 bulan atau Kosong/Tidak tahu')
        # Catatan: batasan rentang untuk tinggi badan, berat badan, dan panjang
        # badan lahir sengaja tidak diberlakukan agar sistem dapat memproses
        # nilai antropometri apa pun (termasuk nilai ekstrem) melalui fungsi
        # keanggotaan berbentuk bahu.

        if errors:
            return render_template('index.html', mpasi_options=MPASI_OPTIONS,
                                   errors=errors, form_data={'nama': nama})

        result = predict(usia, jk, tb, bb, pbl, mpasi, mf_params, who_lookup)

        input_data = {'nama': nama, 'usia': usia, 'jk': jk, 'tb': tb, 'bb': bb, 'pbl': pbl, 'mpasi': mpasi}
        save_to_history(input_data, result)

        jk_label = 'Laki-laki' if jk == 1 else 'Perempuan'

        return render_template('result.html',
            nama=nama, usia=usia, jk=jk, jk_label=jk_label,
            tb=tb, bb=bb, pbl=pbl, mpasi=mpasi,
            result=result)

    except (ValueError, KeyError) as e:
        return render_template('index.html', mpasi_options=MPASI_OPTIONS,
                               errors=[f'Input tidak valid: {str(e)}'])


@app.route('/history')
def history_page():
    records = get_history()
    return render_template('history.html', records=records)


@app.route('/retrain', methods=['POST'])
def retrain_route():
    global trained_model, mf_params
    success, message = auto_retrain()
    if success:
        trained_model = load_trained_model()
        mf_params = parse_mf_params(trained_model)
    return jsonify({'success': success, 'message': message})


@app.route('/api/predict', methods=['POST'])
def api_predict():
    data = request.get_json()
    if not data:
        return jsonify({'error': 'No JSON data'}), 400
    try:
        usia = int(data['usia'])
        jk = int(data['jk'])
        tb = float(data['tb'])
        bb = float(data['bb'])
        pbl = float(data['pbl'])
        mpasi = float(data['mpasi'])
        nama = str(data.get('nama', '') or '')

        result = predict(usia, jk, tb, bb, pbl, mpasi, mf_params, who_lookup)
        input_data = {'nama': nama, 'usia': usia, 'jk': jk, 'tb': tb, 'bb': bb, 'pbl': pbl, 'mpasi': mpasi}
        save_to_history(input_data, result)
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 400


@app.route('/api/model-info')
def api_model_info():
    if trained_model is None:
        return jsonify({'status': 'no model', 'using': 'default params'})
    return jsonify({
        'status': 'loaded',
        'timestamp': trained_model.get('timestamp', 'unknown'),
        'ga_fitness': trained_model.get('ga', {}).get('best_fitness', 'N/A'),
        'metrics': trained_model.get('metrics', {}),
    })


# ============================================================================
# MAIN
# ============================================================================
if __name__ == '__main__':
    print("\n" + "=" * 60)
    print("  STUNTOR - Monitoring Stunting Balita Kalimantan Tengah")
    print("  Website Berbasis AI | Fuzzy Sugeno + GA")
    print("=" * 60)
    print(f"  Model       : {MODEL_PATH}")
    print(f"  Dataset     : {DATASET_FINAL}")
    print(f"  WHO Tables  : {WHO_ZS_DIR}")
    print(f"  History     : {HISTORY_CSV}")
    print(f"  Train Script: {TRAIN_SCRIPT}")

    if trained_model:
        m = trained_model.get('metrics', {}).get('test', {})
        print(f"  Evaluasi Test: P={m.get('precision','?')} R={m.get('recall','?')} F1={m.get('f1','?')} Acc={m.get('accuracy','?')}")
        model_src = "GA-optimized"
    else:
        print(f"  [!] Menggunakan DEFAULT params (belum dilatih)")
        model_src = "default (unoptimized)"

    print(f"  MF Source   : {model_src}")
    print(f"\n  Server: http://localhost:5000")
    print("=" * 60)

    app.run(debug=True, host='0.0.0.0', port=5000)