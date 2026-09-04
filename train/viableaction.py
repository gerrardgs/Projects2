# save as: paper_compute_remaining.py
import pandas as pd
import numpy as np
from scipy.stats import spearmanr

df = pd.read_csv('train/trainfile/dataset_final.csv')

print("=" * 70)
print("DATA UNTUK PAPER — JALANKAN INI DAN KIRIM OUTPUT KE CLAUDE")
print("=" * 70)

# Demografi
print("\n[1] Sex distribution:")
print(df['jenis_kelamin'].value_counts().sort_index())

print("\n[2] Age groups:")
df['age_group'] = pd.cut(df['usia_bulan'], bins=[-1, 11, 23, 35, 47, 59],
                         labels=['0-11', '12-23', '24-35', '36-47', '48-59'])
print(df['age_group'].value_counts().sort_index())

print("\n[3] MP-ASI timing:")
df['mp_timing'] = pd.cut(df['mpasi_bulan'], bins=[-0.01, 5.99, 6.01, 12.01],
                         labels=['Too early', 'Normal', 'Late'])
print(df['mp_timing'].value_counts())

# Spearman
rho, pv = spearmanr(df['risk_score'], df['z_score_ukur'])
print(f"\n[4] Spearman ρ(risk_score, HAZ) = {rho:.4f} (p = {pv:.4e})")

# Prevalence per flag
baseline = (df['label_who_ukur'] == 'Stunting').mean()
print(f"\n[5] Baseline stunting rate: {baseline*100:.2f}%")
for flag in ['Normal', 'STUNTING LOW', 'STUNTING MEDIUM', 'STUNTING HIGH']:
    s = df[df['priority_flag'] == flag]
    if len(s) == 0:
        print(f"   {flag}: n=0"); continue
    n_st = (s['label_who_ukur'] == 'Stunting').sum()
    rate = n_st / len(s)
    print(f"   {flag:18s}: n={len(s):3d}, stunting={n_st:2d} ({rate*100:5.2f}%), ratio={rate/baseline:.2f}x")

# Input correlations
print("\n[6] Input vs HAZ correlations:")
for col in ['pbl_cm', 'tinggi_badan_cm', 'berat_badan_gram', 'mpasi_bulan']:
    pr = df[col].corr(df['z_score_ukur'])
    sr = df[col].corr(df['z_score_ukur'], method='spearman')
    print(f"   {col:25s}: Pearson r={pr:+.4f}, Spearman ρ={sr:+.4f}")