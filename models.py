from flask_login import UserMixin
from extensions import db
from extensions import db
from datetime import datetime

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), nullable=False, unique=True)
    email = db.Column(db.String(150), nullable=False, unique=True)
    phone = db.Column(db.String(20), nullable=False)
    password = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(10), default='staff')

class Ticket(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    kanal_pengaduan = db.Column(db.String(100), nullable=True)
    kategori_pengaduan = db.Column(db.String(100), nullable=True)
    jenis_pengaduan = db.Column(db.String(100), nullable=True)
    detail_pengaduan = db.Column(db.Text, nullable=True)
    tanggal = db.Column(db.DateTime, default=datetime.utcnow, nullable=True)
    #nomor_ticket = db.Column(db.String(100), nullable=True)
    nama_nasabah = db.Column(db.String(150), nullable=True)
    email = db.Column(db.String(150), nullable=True)
    nomor_utama = db.Column(db.String(50), nullable=True)
    nomor_kontak = db.Column(db.String(50), nullable=True)
    nik = db.Column(db.String(50), nullable=True)
    order_no = db.Column(db.String(100), nullable=True)
    deskripsi_pengaduan = db.Column(db.String(1000), nullable=True)
    
    input_by = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    user = db.relationship('User', backref=db.backref('tickets', lazy=True))
    
    status_ticket = db.Column(db.String(50), default='1', nullable=True)
    sla = db.Column(db.Integer, default=10, nullable=True)
    hasil_tindak = db.Column(db.Text, nullable=True)
    hasil_feedback = db.Column(db.Text, nullable=True)
    konfirmasi_nasabah = db.Column(db.Text(100), nullable=True)  
    notes = db.Column(db.Text, nullable=True)
    nama_dc = db.Column(db.String(100), nullable=True)
    nama_os = db.Column(db.String(100), nullable=True)
    nama_bucket = db.Column(db.String(100), nullable=True)
    punishment = db.Column(db.Text, nullable=True)
    hasil_punishment = db.Column(db.Text, nullable=True)
    bukti_chat = db.Column(db.String(300), nullable=True) 
    tahapan = db.Column(db.String(100), nullable=True)
    tahapan_2 = db.Column(db.String(100), nullable=True)
    created_time = db.Column(db.DateTime, default=datetime.utcnow)
    kronologis = db.Column(db.Text, nullable=True)
    status_case = db.Column(db.String(50), nullable=True)
    document = db.Column(db.Text, nullable=True)
    catatan = db.Column(db.Text, nullable=True)
    tanggal_catatan = db.Column(db.String(10)) 


    nomor_ticket_id = db.Column(db.Integer, db.ForeignKey('nomor_ticket.id'), nullable=True)
    nomor_ticket = db.relationship('NomorTicket', back_populates='tickets')

    def __repr__(self):
        return f"<Ticket {self.id}>"
    
class NomorTicket(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nomor_ticket = db.Column(db.String(100), unique=True, nullable=False)
    status = db.Column(db.String(20), default='aktif')

    tickets = db.relationship('Ticket', back_populates='nomor_ticket')

    def __repr__(self):
        return f"<NomorTicket {self.nomor_ticket}>"

class Kontak(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nama_lengkap = db.Column(db.String(150), nullable=False)
    nik = db.Column(db.String(50), nullable=False)
    phone = db.Column(db.String(50), nullable=False)
    phone_2 = db.Column(db.String(50), nullable=True)
    email = db.Column(db.String(150), nullable=True)
    id_ticket = db.Column(db.Integer, db.ForeignKey('ticket.id'), nullable=False)

    ticket = db.relationship('Ticket', backref=db.backref('kontaks', lazy=True))

    def __repr__(self):
        return f"<Kontak {self.nama_lengkap}>"

class History(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nomor_ticket = db.Column(db.String(100), nullable=False)  
    tanggal = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    order_number = db.Column(db.String(100), nullable=True)   
    status_ticket = db.Column(db.String(50), nullable=True)   
    tahapan = db.Column(db.String(100), nullable=True)
    nama_os = db.Column(db.String(100), nullable=True)
    catatan = db.Column(db.String(100), nullable=True)       
    create_by = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)  

    user = db.relationship('User', backref=db.backref('histories', lazy=True))

    def __repr__(self):
        return f"<History {self.nomor_ticket} - {self.tanggal}>"
