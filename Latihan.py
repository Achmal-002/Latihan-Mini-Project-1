# ===================================================
# Program Gabungan: Import–Export Data Mahasiswa
# Author : Achmal Maulana
# ===================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math

# =============================
# BAGIAN 1 — RATA-RATA DASAR
# =============================
print("=== BAGIAN 1: RATA-RATA DASAR ===")

# 1. Import data dari file Excel (contoh: data_mahasiswa.xlsx)
data_awal = pd.read_excel("data_mahasiswa.xlsx")

# 2. Hitung rata-rata dari tiga nilai
data_awal["Rata-rata"] = data_awal[["Nilai 1", "Nilai 2", "Nilai 3"]].mean(axis=1)

# 3. Pembulatan ke atas dengan 2 digit di belakang koma
data_awal["Rata-rata"] = data_awal["Rata-rata"].apply(lambda x: math.ceil(x * 100) / 100)

# 4. Tampilkan hasil di terminal
print("\nData Mahasiswa dengan Rata-rata (dibulatkan 2 digit):")
print(data_awal[["NIM", "Nama Mahasiswa", "Rata-rata"]].head())

# 5. Simpan hasil ke file baru
data_awal.to_excel("hasil_mahasiswa.xlsx", index=False)
print("\nFile 'hasil_mahasiswa.xlsx' berhasil dibuat!")

# ===================================================
# BAGIAN 2 — LATIHAN LANJUTAN (BOBOT + STATUS + TOP 5)
# ===================================================
print("\n=== BAGIAN 2: LATIHAN LANJUTAN (BOBOT + STATUS + TOP 5) ===")

# 1. Import data baru dari file hasil_mahasiswa.xlsx
data = pd.read_excel("hasil_mahasiswa.xlsx")

# 2. Hitung rata-rata berbobot (Tugas 30%, UTS 30%, UAS 40%)
data["Rata-rata Berbobot"] = (
    data["Nilai 1"] * 0.3 +
    data["Nilai 2"] * 0.3 +
    data["Nilai 3"] * 0.4
)

# 3. Pembulatan ke atas dengan 2 digit di belakang koma
data["Rata-rata Berbobot"] = data["Rata-rata Berbobot"].apply(lambda x: math.ceil(x * 100) / 100)

# 4. Tambahkan kolom status
data["Status"] = data["Rata-rata Berbobot"].apply(lambda x: "Lulus" if x >= 75 else "Tidak Lulus")

# 5. Urutkan dari nilai tertinggi
data = data.sort_values(by="Rata-rata Berbobot", ascending=False)

# 6. Ambil 5 mahasiswa terbaik
top5 = data.head(5)
print("\n=== 5 Mahasiswa Terbaik ===")
print(top5[["Nama Mahasiswa", "Rata-rata Berbobot", "Status"]])

# 7. Visualisasi grafik 5 mahasiswa terbaik
plt.figure(figsize=(8,5))
plt.bar(top5["Nama Mahasiswa"], top5["Rata-rata Berbobot"], color='skyblue')
plt.title("5 Mahasiswa dengan Nilai Tertinggi", fontsize=14)
plt.xlabel("Nama Mahasiswa")
plt.ylabel("Rata-rata Berbobot")
plt.xticks(rotation=30, ha='right')
plt.tight_layout()
plt.show()

# 8. Simpan hasil akhir ke file Excel baru
data.to_excel("rekap_nilai.xlsx", index=False)
print("\nFile 'rekap_nilai.xlsx' berhasil dibuat!")

print("\n=== Program Selesai ===")
