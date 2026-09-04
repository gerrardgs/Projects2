# SCRIPT untuk Anda jalankan
import pandas as pd
df = pd.read_csv('train/trainfile/dataset_final.csv')

print("=== Sex distribution ===")
print(df['jenis_kelamin'].value_counts())

print("\n=== Age groups ===")
df['age_group'] = pd.cut(df['usia_bulan'],
                         bins=[-1, 11, 23, 35, 47, 59],
                         labels=['0-11', '12-23', '24-35', '36-47', '48-59'])
print(df['age_group'].value_counts().sort_index())

print("\n=== MP-ASI timing ===")
df['mpasi_timing'] = pd.cut(df['mpasi_bulan'],
                            bins=[-0.01, 5.99, 6.01, 12.01],
                            labels=['Too early (<6)', 'Normal (≈6)', 'Late (>6)'])
print(df['mpasi_timing'].value_counts())