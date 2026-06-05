import pandas as pd
import sqlite3

# Ganti path ini dengan path file CSV yang kamu upload
csv_file = "C:/Users/HP/dashboardpjd/petugas.csv"

# Baca CSV
df = pd.read_csv(csv_file, dtype=str)

# Ubah nama kolom jika perlu
if "nama_karyawan" in df.columns:
    df = df.rename(columns={"nama_karyawan": "nama_pengguna"})

# Tambah kolom default jika belum ada
kolom_tambahan = {
    "nip": "",
    "jabatan": "Casemix",
    "email": "",
    "no_hp": "",
    "status_akun": "Aktif"
}

for kolom, default in kolom_tambahan.items():
    if kolom not in df.columns:
        df[kolom] = default

# Urutkan kolom agar konsisten
df = df[[
    "nama_pengguna", "username", "nip", "jabatan",
    "email", "no_hp", "status_akun", "password"
]]

# Koneksi ke database
db_file = "C:/Users/HP/dashboardpjd/eklaim.db"
conn = sqlite3.connect(db_file)

# Simpan ke tabel pusat_akun, replace kalau sudah ada
df.to_sql("pusat_akun", conn, if_exists="replace", index=True, index_label="id")

# Cek isi tabel
c = conn.cursor()
c.execute("SELECT * FROM pusat_akun LIMIT 5")
rows = c.fetchall()
for r in rows:
    print(r)

conn.close()
print("✅ Tabel pusat_akun berhasil dibuat dan database diisi")