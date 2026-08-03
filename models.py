from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()

class Admin(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(255), nullable=False)

class Peserta(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nama = db.Column(db.String(100), nullable=False)
    nomor_induk = db.Column(db.String(50), unique=True, nullable=False)
    kategori = db.Column(db.String(20), nullable=False) # 'PKL' atau 'MAGANG'
    instansi = db.Column(db.String(100), nullable=False)
    
    absensi_list = db.relationship('Absensi', back_populates='peserta', lazy=True, cascade="all, delete-orphan")

class Absensi(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    peserta_id = db.Column(db.Integer, db.ForeignKey('peserta.id'), nullable=False)
    waktu_masuk = db.Column(db.DateTime, default=datetime.now)
    status = db.Column(db.String(20), default='Hadir') # 'Hadir', 'Izin', atau 'Sakit'
    jurnal_harian = db.Column(db.Text, nullable=True)
    alasan_izin = db.Column(db.Text, nullable=True)
    foto_selfie = db.Column(db.String(255), nullable=True)
    tanda_tangan = db.Column(db.String(255), nullable=True)
    surat_izin = db.Column(db.String(255), nullable=True) # <-- Kolom baru untuk file surat
    latitude = db.Column(db.Float, nullable=True)
    longitude = db.Column(db.Float, nullable=True)

    peserta = db.relationship('Peserta', back_populates='absensi_list')