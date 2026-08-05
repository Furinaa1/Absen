import sqlite3, os
db_path = os.path.join('instance', 'database_absensi.db')
if not os.path.exists(db_path): db_path = 'database_absensi.db'
conn = sqlite3.connect(db_path)
cursor = conn.cursor()
kolom_baru = [
    ("tempat_lahir", "TEXT"), ("tanggal_lahir", "TEXT"), ("jenis_kelamin", "TEXT"),
    ("hoby", "TEXT"), ("jurusan_kelas", "TEXT"), ("no_hp", "TEXT"), ("email", "TEXT"),
    ("alamat", "TEXT"), ("bidang_penempatan", "TEXT"), ("tanggal_mulai", "TEXT"),
    ("tanggal_selesai", "TEXT"), ("nama_pembimbing", "TEXT"), ("nip", "TEXT")
]
for kol, tipe in kolom_baru:
    try:
        cursor.execute(f"ALTER TABLE peserta ADD COLUMN {kol} {tipe};")
        print(f"Kolom {kol} ditambahkan.")
    except Exception as e:
        print(f"Kolom {kol} sudah ada / aman.")
conn.commit()
conn.close()
print("Selesai! Database aman.")