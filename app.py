import os
import base64
from io import BytesIO
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
import qrcode

from models import db, Admin, Peserta, Absensi

app = Flask(__name__)
app.config['SECRET_KEY'] = 'rahasia-absen-pkl-magang'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///database_absensi.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

UPLOAD_FOLDER = os.path.join('static', 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

db.init_app(app)

login_manager = LoginManager()
login_manager.login_view = 'login'
login_manager.init_app(app)

@login_manager.user_loader
def load_user(user_id):
    return Admin.query.get(int(user_id))


# --- ROUTE PRESENSI DAN PERIZINAN ---
@app.route('/presensi', methods=['GET', 'POST'])
def presensi():
    if request.method == 'POST':
        nomor_induk = request.form.get('nomor_induk', '').strip()
        status = request.form.get('status', 'Hadir')
        jurnal = request.form.get('jurnal', '').strip()
        alasan_izin = request.form.get('alasan_izin', '').strip()
        lat = request.form.get('latitude')
        long = request.form.get('longitude')
        foto_base64 = request.form.get('foto_base64')
        ttd_base64 = request.form.get('ttd_base64')

        peserta = Peserta.query.filter_by(nomor_induk=nomor_induk).first()
        if not peserta:
            flash("NIM / NISN tidak terdaftar! Periksa kembali.", "danger")
            return redirect(url_for('presensi'))

        # 1. Olah Simpan Foto Selfie (Jika Hadir)
        filename_foto = None
        if status == 'Hadir' and foto_base64 and ',' in foto_base64:
            try:
                header, encoded = foto_base64.split(',', 1)
                data = base64.b64decode(encoded)
                filename_foto = f"selfie_{nomor_induk}_{int(datetime.now().timestamp())}.jpg"
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename_foto)
                with open(filepath, "wb") as f:
                    f.write(data)
            except Exception as e:
                print(f"Gagal menyimpan foto: {e}")

        # 2. Olah Simpan Tanda Tangan
        filename_ttd = None
        if ttd_base64 and ',' in ttd_base64:
            try:
                header, encoded = ttd_base64.split(',', 1)
                data = base64.b64decode(encoded)
                filename_ttd = f"ttd_{nomor_induk}_{int(datetime.now().timestamp())}.png"
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename_ttd)
                with open(filepath, "wb") as f:
                    f.write(data)
            except Exception as e:
                print(f"Gagal menyimpan TTD: {e}")

        # 3. Olah Simpan File Surat Izin / Sakit (Jika Izin atau Sakit)
        filename_surat = None
        if status != 'Hadir':
            file_surat = request.files.get('surat_izin')
            if file_surat and file_surat.filename != '':
                filename_surat = f"surat_{nomor_induk}_{int(datetime.now().timestamp())}.jpg"
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename_surat)
                file_surat.save(filepath)

        # Simpan Data Presensi / Izin Ke DB
        absen_baru = Absensi(
            peserta_id=peserta.id,
            status=status,
            jurnal_harian=jurnal if status == 'Hadir' else None,
            alasan_izin=alasan_izin if status != 'Hadir' else None,
            foto_selfie=filename_foto,
            tanda_tangan=filename_ttd,
            surat_izin=filename_surat,  # <-- Menyimpan nama file surat ke database
            latitude=float(lat) if lat and lat != '' else None,
            longitude=float(long) if long and long != '' else None
        )
        db.session.add(absen_baru)
        db.session.commit()

        flash(f"Pengajuan {status} Berhasil Dikirim untuk {peserta.nama}!", "success")
        return redirect(url_for('presensi'))

    return render_template('presensi.html')


@app.route('/api/get-nama/<nomor_induk>')
def get_nama(nomor_induk):
    peserta = Peserta.query.filter_by(nomor_induk=nomor_induk).first()
    if peserta:
        return jsonify({"success": True, "nama": peserta.nama, "kategori": peserta.kategori})
    else:
        return jsonify({"success": False, "message": "NIM / NISN tidak ditemukan!"})


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        admin = Admin.query.filter_by(username=username).first()
        
        if admin and check_password_hash(admin.password, password):
            login_user(admin)
            return redirect(url_for('admin_dashboard'))
        else:
            flash("Username atau password salah!", "danger")
            
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/admin')
@login_required
def admin_dashboard():
    list_peserta = Peserta.query.all()
    rekap_absensi = Absensi.query.order_by(Absensi.waktu_masuk.desc()).all()
    return render_template('admin_dashboard.html', peserta=list_peserta, absensi=rekap_absensi)

@app.route('/admin/peserta/tambah', methods=['POST'])
@login_required
def tambah_peserta():
    nama = request.form.get('nama')
    nomor_induk = request.form.get('nomor_induk', '').strip()
    kategori = request.form.get('kategori')
    instansi = request.form.get('instansi')
    
    if Peserta.query.filter_by(nomor_induk=nomor_induk).first():
        flash("NIM/NISN sudah terdaftar!", "warning")
    else:
        peserta_baru = Peserta(nama=nama, nomor_induk=nomor_induk, kategori=kategori, instansi=instansi)
        db.session.add(peserta_baru)
        db.session.commit()
        flash("Peserta berhasil ditambahkan!", "success")
        
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/peserta/hapus/<int:id>')
@login_required
def hapus_peserta(id):
    peserta = Peserta.query.get_or_404(id)
    db.session.delete(peserta)
    db.session.commit()
    flash("Data peserta berhasil dihapus!", "success")
    return redirect(url_for('admin_dashboard'))

@app.route('/admin/generate-qr')
@login_required
def generate_qr():
    domain_https = "https://mollusk-flammable-delegate.ngrok-free.dev"
    url_presensi = f"{domain_https}/presensi"
    qr = qrcode.make(url_presensi)
    os.makedirs('static', exist_ok=True)
    qr.save('static/qr_absensi.png')
    flash("QR Code berhasil dibuat ulang!", "success")
    return redirect(url_for('admin_dashboard'))

@app.route('/export_excel')
@login_required
def export_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Rekap Absensi & Izin"

    ws.append(["No", "Waktu", "Nama", "Kategori", "Instansi", "NIM/NISN", "Status", "Jurnal / Alasan Izin", "Koordinat GPS", "Nama File Foto", "Nama File TTD", "Nama File Surat"])

    absensi_list = Absensi.query.order_by(Absensi.waktu_masuk.desc()).all()
    for idx, item in enumerate(absensi_list, 1):
        keterangan = item.jurnal_harian if item.status == 'Hadir' else item.alasan_izin
        ws.append([
            idx,
            item.waktu_masuk.strftime('%Y-%m-%d %H:%M:%S'),
            item.peserta.nama,
            item.peserta.kategori,
            item.peserta.instansi,
            item.peserta.nomor_induk,
            item.status,
            keterangan if keterangan else "-",
            f"{item.latitude}, {item.longitude}" if item.latitude else "-",
            item.foto_selfie if item.foto_selfie else "-",
            item.tanda_tangan if item.tanda_tangan else "-",
            item.surat_izin if item.surat_izin else "-"
        ])

    file_stream = BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Rekap_Absensi_Izin.xlsx'
    )

def init_db():
    with app.app_context():
        # db.drop_all()  <-- Baris ini sudah dihapus/dimatikan
        db.create_all()   # Hanya membuat tabel jika belum ada, data lama akan aman tersimpan
        
        # Opsional: Memastikan akun admin tetap ada tanpa mereset data lain
        if not Admin.query.filter_by(username='instruktur').first():
            hashed_pw = generate_password_hash('admin1234')
            admin_default = Admin(username='instruktur', password=hashed_pw)
            db.session.add(admin_default)
            db.session.commit()
            print("Akun admin default dibuat!")

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)