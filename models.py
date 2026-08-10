from datetime import datetime
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin

db = SQLAlchemy()

class Admin(db.Model, UserMixin):
    __tablename__ = 'admin'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(255), nullable=False)


class Peserta(db.Model):
    __tablename__ = 'peserta'
    id = db.Column(db.Integer, primary_key=True)
    nama = db.Column(db.String(100), nullable=True)
    nomor_induk = db.Column(db.String(50), unique=True, nullable=False)
    kategori = db.Column(db.String(50), nullable=True)
    instansi = db.Column(db.String(150), nullable=True)
    
    tempat_lahir = db.Column(db.String(100), nullable=True)
    tanggal_lahir = db.Column(db.String(50), nullable=True)
    jenis_kelamin = db.Column(db.String(20), nullable=True)
    hoby = db.Column(db.String(100), nullable=True)
    jurusan_kelas = db.Column(db.String(100), nullable=True)
    no_hp = db.Column(db.String(20), nullable=True)
    email = db.Column(db.String(100), nullable=True)
    alamat = db.Column(db.Text, nullable=True)
    bidang_penempatan = db.Column(db.String(100), nullable=True)
    tanggal_mulai = db.Column(db.String(50), nullable=True)
    tanggal_selesai = db.Column(db.String(50), nullable=True)
    nama_pembimbing = db.Column(db.String(100), nullable=True)
    nip = db.Column(db.String(50), nullable=True)
    
    # Relasi ke Absensi
    absensi = db.relationship('Absensi', backref='peserta', lazy=True, cascade="all, delete-orphan")


class Absensi(db.Model):
    __tablename__ = 'absensi'
    id = db.Column(db.Integer, primary_key=True)
    peserta_id = db.Column(db.Integer, db.ForeignKey('peserta.id'), nullable=False)
    waktu_masuk = db.Column(db.DateTime, default=datetime.now)
    status = db.Column(db.String(20), default='Hadir Pagi')
    jurnal_harian = db.Column(db.Text, nullable=True)
    alasan_izin = db.Column(db.Text, nullable=True)
    
    # Diubah menjadi db.Text agar muat menampung URL Cloudinary yang panjang
    foto_selfie = db.Column(db.Text, nullable=True)
    tanda_tangan = db.Column(db.Text, nullable=True)
    surat_izin = db.Column(db.Text, nullable=True)
    
    latitude = db.Column(db.Float, nullable=True)
    longitude = db.Column(db.Float, nullable=True)