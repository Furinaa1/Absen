import sqlite3
import os

# Sesuaikan path database Anda (biasanya di dalam folder instance/database_absensi.db)
db_path = os.path.join('instance', 'database_absensi.db')

if not os.path.exists(db_path):
    # Jika tidak ada di folder instance, coba cari langsung di direktori utama
    db_path = 'database_absensi.db'

print(f"Menghubungkan ke database: {db_path}")
conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Daftar kolom baru yang ingin ditambahkan ke tabel peserta
kolom_baru = [
    ("tempat_lahir", "TEXT"),
    ("tanggal_lahir", "TEXT"),
    ("jenis_kelamin", "TEXT"),
    ("hoby", "TEXT"),
    ("jurusan_kelas", "TEXT"),
    ("no_hp", "TEXT"),
    ("email", "TEXT"),
    ("alamat", "TEXT"),
    ("bidang_penempatan", "TEXT"),
    ("tanggal_mulai", "TEXT"),
    ("tanggal_selesai", "TEXT"),
    ("nama_pembimbing", "TEXT"),
    ("nip", "TEXT")
]

for kol, tipe in kolom_baru:
    try:
        cursor.execute(f"ALTER TABLE peserta ADD COLUMN {kol} {tipe};")
        print(f"Berhasil menambahkan kolom: {kol}")
    except sqlite3.OperationalError as e:
        print(f"Kolom '{kol}' mungkin sudah ada: {e}")

conn.commit()
conn.close()
print("Upgrade database selesai! Sekarang Anda bisa menjalankan ulang app.py.")