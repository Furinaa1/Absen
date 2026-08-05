import os
import base64
from io import BytesIO
from datetime import datetime
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import qrcode

# Import openpyxl untuk manipulasi sheet Excel
from openpyxl import Workbook

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


# --- ROUTE CETAK BIODATA PESERTA ---
@app.route('/admin/peserta/cetak/<int:id>')
@login_required
def cetak_biodata_peserta(id):
    peserta = Peserta.query.get_or_404(id)
    return render_template('cetak_biodata.html', peserta=peserta)


# --- ROUTE UTAMA ---
@app.route('/')
def index():
    return redirect(url_for('presensi'))


# --- ROUTE PRESENSI DAN PERIZINAN ---
@app.route('/presensi', methods=['GET', 'POST'])
def presensi():
    if request.method == 'POST':
        nomor_induk = request.form.get('nomor_induk', '').strip()
        status = request.form.get('status', 'Hadir Pagi')
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

        # 1. Olah Simpan Foto Selfie (Jika Hadir Pagi atau Pulang Sore)
        filename_foto = None
        if status in ['Hadir Pagi', 'Pulang Sore'] and foto_base64 and ',' in foto_base64:
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
        if status not in ['Hadir Pagi', 'Pulang Sore']:
            file_surat = request.files.get('surat_izin')
            if file_surat and file_surat.filename != '':
                filename_surat = f"surat_{nomor_induk}_{int(datetime.now().timestamp())}.jpg"
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename_surat)
                file_surat.save(filepath)

        # Simpan Data Presensi / Izin Ke DB
        absen_baru = Absensi(
            peserta_id=peserta.id,
            status=status,
            jurnal_harian=jurnal if status == 'Pulang Sore' else None,
            alasan_izin=alasan_izin if status not in ['Hadir Pagi', 'Pulang Sore'] else None,
            foto_selfie=filename_foto,
            tanda_tangan=filename_ttd,
            surat_izin=filename_surat,  
            latitude=float(lat) if lat and lat != '' else None,
            longitude=float(long) if long and long != '' else None
        )
        db.session.add(absen_baru)
        db.session.commit()

        flash(f"Pengajuan {status} Berhasil Dikirim untuk {peserta.nama}!", "success")
        return redirect(url_for('presensi'))

    return render_template('presensi.html')


# --- ROUTE FORM & PENYIMPANAN BIODATA SISWA/MAHASISWA (VIA QR) ---
@app.route('/biodata', methods=['GET', 'POST'])
def biodata():
    if request.method == 'POST':
        nomor_induk = request.form.get('nomor_induk', '').strip()
        
        peserta = Peserta.query.filter_by(nomor_induk=nomor_induk).first()
        if not peserta:
            flash('NIM / NISN belum terdaftar di sistem admin! Silakan hubungi pembimbing.', 'danger')
            return redirect(url_for('biodata'))
        
        peserta.nama = request.form.get('nama')
        peserta.tempat_lahir = request.form.get('tempat_lahir')
        peserta.tanggal_lahir = request.form.get('tanggal_lahir')
        peserta.jenis_kelamin = request.form.get('jenis_kelamin')
        peserta.hoby = request.form.get('hoby')
        peserta.instansi = request.form.get('instansi')
        peserta.jurusan_kelas = request.form.get('jurusan_kelas')
        peserta.no_hp = request.form.get('no_hp')
        peserta.email = request.form.get('email')
        peserta.alamat = request.form.get('alamat')
        peserta.bidang_penempatan = request.form.get('bidang_penempatan')
        peserta.tanggal_mulai = request.form.get('tanggal_mulai')
        peserta.tanggal_selesai = request.form.get('tanggal_selesai')
        peserta.nama_pembimbing = request.form.get('nama_pembimbing')
        peserta.nip = request.form.get('nip')
        peserta.kategori = request.form.get('kategori', 'PKL')
        
        db.session.commit()
        flash('Biodata lengkap berhasil disimpan!', 'success')
        return redirect(url_for('biodata'))

    return render_template('biodata.html')


# --- ROUTE REKAP & CETAK BIODATA LENGKAP ---
@app.route('/admin/biodata-peserta')
@login_required
def admin_biodata_peserta():
    all_peserta = Peserta.query.all()
    return render_template('admin_biodata_peserta.html', peserta_list=all_peserta)


# --- ROUTE DOWNLOAD REKAP WORD BIODATA ---
@app.route('/admin/export-word-biodata')
@login_required
def export_word_biodata():
    doc = Document()
    
    list_peserta = Peserta.query.all()
    if not list_peserta:
        doc.add_heading('Rekap Biodata Lengkap Peserta PKL/Magang', level=1)
        doc.add_paragraph('Belum ada data peserta.')
    else:
        for idx, p in enumerate(list_peserta, 1):
            # Judul Dokumen di setiap halaman baru
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_title = p_title.add_run('Rekap Biodata Lengkap Peserta PKL/Magang\n')
            run_title.bold = True
            run_title.font.size = Pt(14)
            
            # Tabel Biodata
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Kolom Data'
            hdr_cells[1].text = 'Keterangan'
            
            # Nama ditaruh di dalam tabel, tepat di atas NIM / NISN
            data_fields = [
                ("Nama Lengkap", p.nama),
                ("NIM / NISN", p.nomor_induk),
                ("Kategori", p.kategori),
                ("Tempat Lahir", p.tempat_lahir),
                ("Tanggal Lahir", p.tanggal_lahir),
                ("Jenis Kelamin", p.jenis_kelamin),
                ("Hobi", p.hoby),
                ("Asal Sekolah / Kampus", p.instansi),
                ("Jurusan / Kelas", p.jurusan_kelas),
                ("No HP", p.no_hp),
                ("Email", p.email),
                ("Alamat", p.alamat),
                ("Bidang Penempatan", p.bidang_penempatan),
                ("Tanggal Mulai", p.tanggal_mulai),
                ("Tanggal Selesai", p.tanggal_selesai),
                ("Nama Pembimbing", p.nama_pembimbing),
                ("NIP Pembimbing", p.nip)
            ]
            
            for label, val in data_fields:
                row_cells = table.add_row().cells
                row_cells[0].text = label
                row_cells[1].text = str(val if val else '-')
                
            # Spasi pemisah sebelum bagian tanda tangan
            doc.add_paragraph()

            # Tabel untuk Pas Foto dan Tanda Tangan (1 Baris, 2 Kolom)
            ttd_table = doc.add_table(rows=1, cols=2)
            
            cell_kiri = ttd_table.rows[0].cells[0]
            cell_kanan = ttd_table.rows[0].cells[1]

            # Kolom Kiri: Kotak Pas Foto 3x4
            p_foto = cell_kiri.paragraphs[0]
            p_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_foto.add_run("Pas foto\n3 x 4\nBerwarna")

            # Kolom Kanan: Tempat, Tanggal, dan Tanda Tangan
            p_ttd = cell_kanan.paragraphs[0]
            p_ttd.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p_ttd.add_run("Samarinda, ...................................\n")
            p_ttd.add_run("Hormat saya\n\n\n\n")
            run_ttd_nama = p_ttd.add_run(f"( {p.nama if p.nama else '...................................'} )")
            run_ttd_nama.bold = True

            # Page Break: Pastikan peserta berikutnya berada di 1 kertas/halaman baru
            if idx < len(list_peserta):
                doc.add_page_break()

    file_stream = BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='Rekap_Biodata_Peserta.docx'
    )


# --- ROUTE DOWNLOAD REKAP EXCEL PER PESERTA (MULTI-SHEET + LINK FOTO/TTD) ---
@app.route('/export_excel')
@login_required
def export_excel():
    list_peserta = Peserta.query.all()
    file_path = 'rekap_absensi_per_peserta.xlsx'
    
    base_url = "http://127.0.0.1:5000"
    
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    if not list_peserta:
        ws = wb.create_sheet(title='Tidak Ada Peserta')
        ws.append([
            'Waktu Masuk', 'Nama Peserta', 'Nomor Induk', 
            'Kategori', 'Instansi', 'Status', 
            'Keterangan / Jurnal / Alasan', 'Koordinat GPS', 
            'Foto Selfie', 'Tanda Tangan'
        ])
    else:
        for p in list_peserta:
            nama_sheet = "".join(c for c in p.nama if c.isalnum() or c in (' ', '_'))[:30].strip()
            if not nama_sheet:
                nama_sheet = f"Peserta_{p.id}"
            
            ws = wb.create_sheet(title=nama_sheet)
            
            headers = [
                'Waktu Masuk', 'Nama Peserta', 'Nomor Induk', 
                'Kategori', 'Instansi', 'Status', 
                'Keterangan / Jurnal / Alasan', 'Koordinat GPS', 
                'Foto Selfie', 'Tanda Tangan'
            ]
            ws.append(headers)
            
            ws.column_dimensions['I'].width = 25
            ws.column_dimensions['J'].width = 25
            
            data_absensi_peserta = Absensi.query.filter_by(peserta_id=p.id).order_by(Absensi.waktu_masuk.desc()).all()
            
            for index, a in enumerate(data_absensi_peserta, start=2):
                waktu_str = a.waktu_masuk.strftime('%Y-%m-%d %H:%M:%S') if a.waktu_masuk else ''
                keterangan = a.jurnal_harian if a.status == 'Pulang Sore' else a.alasan_izin
                koordinat = f"{a.latitude}, {a.longitude}" if a.latitude and a.longitude else ''
                
                link_foto = f'=HYPERLINK("{base_url}/static/uploads/{a.foto_selfie}", "Lihat Foto")' if a.foto_selfie else '-'
                link_ttd = f'=HYPERLINK("{base_url}/static/uploads/{a.tanda_tangan}", "Lihat TTD")' if a.tanda_tangan else '-'
                
                row_data = [
                    waktu_str,
                    p.nama or '',
                    p.nomor_induk or '',
                    p.kategori or '',
                    p.instansi or '',
                    a.status,
                    keterangan or '',
                    koordinat,
                    link_foto, 
                    link_ttd  
                ]
                ws.append(row_data)

    wb.save(file_path)
    return send_file(file_path, as_attachment=True)


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
    
    tanggal_mulai = request.args.get('tanggal_mulai')
    tanggal_selesai = request.args.get('tanggal_selesai')
    
    query = Absensi.query
    
    if tanggal_mulai:
        try:
            dt_mulai = datetime.strptime(tanggal_mulai, '%Y-%m-%d')
            query = query.filter(Absensi.waktu_masuk >= dt_mulai)
        except ValueError:
            pass
            
    if tanggal_selesai:
        try:
            dt_selesai = datetime.strptime(tanggal_selesai + ' 23:59:59', '%Y-%m-%d %H:%M:%S')
            query = query.filter(Absensi.waktu_masuk <= dt_selesai)
        except ValueError:
            pass
            
    rekap_absensi = query.order_by(Absensi.waktu_masuk.desc()).all()
    
    return render_template(
        'admin_dashboard.html', 
        peserta=list_peserta, 
        absensi=rekap_absensi,
        tanggal_mulai=tanggal_mulai,
        tanggal_selesai=tanggal_selesai
    )


# --- ROUTE KELOLA PESERTA ---
@app.route('/admin/peserta')
@login_required
def admin_peserta():
    all_peserta = Peserta.query.all()
    return render_template('admin_peserta.html', peserta_list=all_peserta)


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
    os.makedirs('static', exist_ok=True)
    
    # Generate QR Presensi
    url_presensi = f"{domain_https}/presensi"
    qr_presensi = qrcode.make(url_presensi)
    qr_presensi.save('static/qr_absensi.png')

    # Generate QR Biodata
    url_biodata = f"{domain_https}/biodata"
    qr_biodata = qrcode.make(url_biodata)
    qr_biodata.save('static/qr_biodata.png')

    flash("Semua QR Code (Presensi & Biodata) berhasil dibuat ulang!", "success")
    return redirect(url_for('admin_dashboard'))

def init_db():
    with app.app_context():
        db.create_all()   
        if not Admin.query.filter_by(username='instruktur').first():
            hashed_pw = generate_password_hash('admin1234')
            admin_default = Admin(username='instruktur', password=hashed_pw)
            db.session.add(admin_default)
            db.session.commit()
            print("Akun admin default dibuat!")

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000, debug=True)