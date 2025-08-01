from flask import Flask, render_template, redirect, url_for, request, flash, send_file, send_from_directory
from flask_login import login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
import os
from extensions import db, migrate, login_manager
from datetime import datetime
from models import Ticket, NomorTicket, User, Kontak, History, db
from sqlalchemy.orm import joinedload, aliased
from sqlalchemy import func, or_, and_, asc, distinct
from werkzeug.utils import secure_filename
import pandas as pd
from io import BytesIO
from flask_apscheduler import APScheduler
from pytz import timezone
import pytz
from werkzeug.serving import is_running_from_reloader

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://root:@localhost/dashboard-cs3'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = 'b35dfe6ce150230940bd145823034486'
app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['MAX_CONTENT_LENGTH'] = 150 * 1024 * 1024 

UPLOAD_FOLDER = os.path.join('static', 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename  

db.init_app(app)
migrate.init_app(app, db)
login_manager.init_app(app)
login_manager.login_view = 'login'

@app.context_processor
def inject_sla_warning_tickets():
    subquery = (
        db.session.query(
            Ticket.nomor_ticket_id,
            func.min(Ticket.sla).label("min_sla")
        )
        .filter(Ticket.sla.between(1, 3))
        .group_by(Ticket.nomor_ticket_id)
    ).subquery()

    TicketAlias = aliased(Ticket)

    sla_warning_tickets = (
        db.session.query(TicketAlias)
        .join(subquery, and_(
            TicketAlias.nomor_ticket_id == subquery.c.nomor_ticket_id,
            TicketAlias.sla == subquery.c.min_sla
        ))
        .join(NomorTicket, NomorTicket.id == TicketAlias.nomor_ticket_id)
        .filter(NomorTicket.status.in_(['aktif', 'Reopen']))
        .order_by(TicketAlias.sla.asc())
        .all()
    )

    return {'sla_warning_tickets': sla_warning_tickets}

class Config:
    SCHEDULER_API_ENABLED = True

app.config.from_object(Config())

scheduler = APScheduler()
scheduler.init_app(app)

@scheduler.task('cron', id='decrease_sla_daily', hour=0, minute=0, timezone='Asia/Jakarta')
def decrease_sla():
    with app.app_context():
        tickets = Ticket.query.filter(Ticket.sla > 0).all()
        for ticket in tickets:
            ticket.sla -= 1
        db.session.commit()
        print(f"SLA updated at {datetime.now(timezone('Asia/Jakarta'))} — {len(tickets)} ticket(s) updated.")

# Job 2: Jalan setiap 1 menit
def update_ticket_fields():
    with app.app_context():
        tickets = Ticket.query.filter(
            (Ticket.nama_os.in_([None, "-", "None"])) |
            (Ticket.nama_bucket.in_([None, "-", "None"]))
        ).all()

        for ticket in tickets:
            if ticket.nama_os in [None, "-", "None"]:
                ticket.nama_os = ""
            if ticket.nama_bucket in [None, "-", "None"]:
                ticket.nama_bucket = ""

        db.session.commit()
        print(f"Updated {len(tickets)} ticket(s).")

# Tambahkan job 2 ke scheduler
scheduler.add_job(
    id='update_none_fields',
    func=update_ticket_fields,
    trigger='interval',
    minutes=1
)

# Hanya jalankan scheduler jika belum jalan
if not scheduler.running:
    scheduler.start()

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

@app.errorhandler(404)
def page_not_found(error):
    return render_template('404.html'), 404

@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method == 'POST':
        username = request.form['username']
        email = request.form['email']
        phone = request.form['phone']
        password = request.form['password']
        hashed_pw = generate_password_hash(password)

        if User.query.filter_by(username=username).first():
            flash('Username sudah terdaftar.')
            return redirect(url_for('register'))
        if User.query.filter_by(email=email).first():
            flash('Email sudah terdaftar.')
            return redirect(url_for('register'))

        user = User(username=username, email=email, phone=phone, password=hashed_pw)
        db.session.add(user)
        db.session.commit()
        flash('Registrasi berhasil! Silakan login.')
        return redirect(url_for('login'))

    return render_template('register.html')

@app.route('/')
def home():
    return redirect(url_for('login'))

@app.route('/history')
@login_required
def history():
    page = request.args.get('page', 1, type=int)
    history_list = History.query.order_by(History.tanggal.desc()).paginate(page=page, per_page=10, error_out=False)
    return render_template('history.html', user=current_user, history_list=history_list)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        login_input = request.form['username'] 
        password = request.form['password']

        user = User.query.filter(
            (User.username == login_input) | (User.email == login_input)
        ).first()

        if user and check_password_hash(user.password, password):
            login_user(user)
            flash('Login berhasil!')
            if user.role == 'admin':
                return redirect(url_for('admin_dashboard'))
            elif user.role == 'qc':
                return redirect(url_for('qc_dashboard'))
            else:
                return redirect(url_for('staff_dashboard'))
        else:
            flash('Username/email atau password salah.')

    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Anda telah logout.')
    return redirect(url_for('login'))

@app.route('/admin_dashboard')
@login_required
def admin_dashboard():
    if current_user.role != 'admin':
        flash('Akses ditolak: Anda bukan admin.')
        return redirect(url_for('staff_dashboard'))
    
    staff_users = User.query.filter(User.role != 'admin').all()
    
    return render_template('admin_dashboard.html', user=current_user, users=staff_users)

@app.route('/qc-dashboard')
@login_required
def qc_dashboard():
    if current_user.role != 'qc':
        flash('Akses ditolak: Anda bukan QC!!!')
        return redirect(url_for('index'))

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(
        NomorTicket.id_qc == current_user.id,
        or_(
            NomorTicket.status == None,
            and_(
                NomorTicket.status != 'close',
                NomorTicket.status != 'reopen'
            )
        )
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))
        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla != 0)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(NomorTicket.status == 'aktif', Ticket.sla != 0, NomorTicket.id_qc == current_user.id)\
        .distinct()\
        .count()

    return render_template(
        'qc_dashboard.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif
    )

@app.route('/qc/nomor-ticket/<int:nomor_ticket_id>')
@login_required
def list_ticket_by_nomor_qc(nomor_ticket_id):
    if current_user.role != 'qc':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.filter_by(id=nomor_ticket_id, id_qc=current_user.id).first_or_404()

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }

    qc_users = User.query.filter_by(role='qc').all()

    return render_template(
        'list_ticket_qc.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map,
        qc_users=qc_users
    )

@app.route('/follow-up-pengaduan-qc/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def follow_up_pengaduan_qc(nomor_ticket_id):
    if current_user.role != 'qc':
        return redirect(request.referrer)

    deskripsi_qc = request.form.get("deskripsi_qc")  
    uploaded_files = request.files.getlist("file_qc")  

    existing = request.form.getlist("existing_images")
    deleted = request.form.getlist("deleted_images")

    new_filenames = []
    for file in uploaded_files:
        if file and file.filename:
            filename = secure_filename(file.filename)
            save_path = os.path.join("static/uploads", filename)
            file.save(save_path)
            new_filenames.append(filename)

    final_files = [f for f in existing if f not in deleted] + new_filenames
    joined_filenames = ",".join(final_files)

    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket.id).all()

    for ticket in tickets:
        ticket.deskripsi_qc = deskripsi_qc
        ticket.file_qc = joined_filenames

    db.session.commit()
    flash("Data berhasil diupdate", "success")
    return redirect(url_for("list_ticket_by_nomor_qc", nomor_ticket_id=nomor_ticket_id))

@app.route('/list_user')
@login_required
def list_user():
    if current_user.role != 'admin':
        flash('Akses ditolak: Hanya admin yang bisa melihat daftar user.')
        return redirect(url_for('staff_dashboard'))

    staff_users = User.query.filter_by(role='staff').all()
    return render_template('list_user.html', user=current_user, users=staff_users)

@app.route('/add_user', methods=['POST'])
@login_required
def add_user():
    if current_user.role != 'admin':
        flash('Akses ditolak: Anda bukan admin.')
        return redirect(url_for('list_user'))

    username = request.form['username']
    email = request.form['email']
    phone = request.form['phone']
    password = request.form['password']
    role = request.form.get('role') 
    hashed_pw = generate_password_hash(password)

    if User.query.filter_by(username=username).first():
        flash('Username sudah terdaftar.')
        return redirect(url_for('list_user'))
    if User.query.filter_by(email=email).first():
        flash('Email sudah terdaftar.')
        return redirect(url_for('list_user'))

    user = User(username=username, email=email, phone=phone, password=hashed_pw, role=role)
    db.session.add(user)
    db.session.commit()

    flash(f'User {role} berhasil ditambahkan.')
    return redirect(url_for('admin_dashboard'))

@app.route('/delete_user/<int:user_id>', methods=['POST'])
@login_required
def delete_user(user_id):
    if current_user.role != 'admin':
        flash('Akses ditolak: Anda bukan admin.', 'error')
        return redirect(url_for('list_user'))

    user = User.query.get_or_404(user_id)

    if user.id == current_user.id:
        flash('Anda tidak dapat menghapus akun Anda sendiri.', 'error')
        return redirect(url_for('list_user'))

    db.session.delete(user)
    db.session.commit()
    flash(f'User {user.username} berhasil dihapus.', 'success')
    return redirect(url_for('list_user'))

@app.route('/filtering', methods=['GET'])
@login_required
def filtering():
    os_selected = request.args.getlist('os') or []
    bucket_selected = request.args.getlist('bucket') or []
    range1 = request.args.get('range1', '')
    range2 = request.args.get('range2', '')

    def parse_range(date_range):
        try:
            start, end = date_range.split(' - ')
            return start.strip(), end.strip()
        except ValueError:
            return None, None

    def format_range_label(start, end):
        try:
            start_fmt = datetime.strptime(start, "%Y-%m-%d").strftime("%d %b")
            end_fmt = datetime.strptime(end, "%Y-%m-%d").strftime("%d %b")
            return f"{start_fmt} - {end_fmt}"
        except:
            return "Range"

    range1_start, range1_end = parse_range(range1)
    range2_start, range2_end = parse_range(range2)

    label_range1 = format_range_label(range1_start, range1_end)
    label_range2 = format_range_label(range2_start, range2_end) if range2_start and range2_end else None

    def get_filtered_data(start_date, end_date):
        query = Ticket.query
        query = query.filter(Ticket.nama_os.isnot(None)).filter(Ticket.nama_os != '')

        if os_selected:
            query = query.filter(Ticket.nama_os.in_(os_selected))
        if bucket_selected:
            query = query.filter(Ticket.nama_bucket.in_(bucket_selected))
        if start_date and end_date:
            try:
                start = datetime.strptime(start_date, "%Y-%m-%d")
                end = datetime.strptime(end_date, "%Y-%m-%d")
                query = query.filter(Ticket.tanggal.between(start, end))
            except ValueError:
                pass

        data_grouped = query.with_entities(
            Ticket.nama_os,
            Ticket.nama_bucket,
            func.count(Ticket.id)
        ).group_by(Ticket.nama_os, Ticket.nama_bucket).all()

        os_totals = {}
        os_buckets = {}

        for os, bucket, count in data_grouped:
            if os:
                os_totals[os] = os_totals.get(os, 0) + count
                if bucket: 
                    if os not in os_buckets:
                        os_buckets[os] = []
                    os_buckets[os].append(f"{bucket}: {count}")

        return os_totals, os_buckets

    os_count1, bucket_info1 = get_filtered_data(range1_start, range1_end)
    os_count2, bucket_info2 = get_filtered_data(range2_start, range2_end) if range2_start and range2_end else ({}, {})

    chart_labels = sorted(list(set(os_count1.keys()) | set(os_count2.keys())))

    chart_series = []

    chart_series.append({
        "name": label_range1,
        "data": [os_count1.get(os, 0) for os in chart_labels],
        "bucket_info": [bucket_info1.get(os, []) for os in chart_labels]
    })

    if os_count2:
        chart_series.append({
            "name": label_range2,
            "data": [os_count2.get(os, 0) for os in chart_labels],
            "bucket_info": [bucket_info2.get(os, []) for os in chart_labels]
        })

    list_os = db.session.query(Ticket.nama_os).distinct().all()
    list_bucket = db.session.query(Ticket.nama_bucket).distinct().all()

    default_colors = [
        "#1E90FF", "#28a745", "#ffc107", "#dc3545", "#6f42c1",
        "#20c997", "#fd7e14", "#6610f2", "#17a2b8", "#343a40"
    ]
    color_map = {label: default_colors[i % len(default_colors)] for i, label in enumerate(chart_labels)}
    chart_colors = [color_map[os] for os in chart_labels]

    return render_template('filtering.html',
        user=current_user,
        chart_labels=chart_labels,
        chart_series=chart_series,
        chart_colors=chart_colors,
        list_os=[os[0] for os in list_os if os[0]],
        list_bucket=[b[0] for b in list_bucket if b[0]],
        os_selected=os_selected,
        bucket_selected=bucket_selected,
        range1=range1,
        range2=range2
    )
from datetime import datetime, timedelta
from collections import defaultdict
@app.route('/filtering-kanal', methods=['GET'])
@login_required
def filtering_kanal():
    range1 = request.args.get('range1', '')
    range2 = request.args.get('range2', '')

    def parse_range(date_range):
        try:
            start, end = date_range.split(' - ')
            start_dt = datetime.strptime(start.strip(), '%Y-%m-%d')
            end_dt = datetime.strptime(end.strip(), '%Y-%m-%d') + timedelta(days=1)
            return start_dt, end_dt
        except:
            return None, None

    start1, end1 = parse_range(range1)
    start2, end2 = parse_range(range2)

    # Ambil semua kanal (normalized ke lowercase dan strip spasi)
    kanal_raw = db.session.query(Ticket.kanal_pengaduan)\
        .filter(Ticket.kanal_pengaduan.isnot(None), Ticket.kanal_pengaduan != '')\
        .distinct().all()

    kanal_set = set()
    for k in kanal_raw:
        kanal_normalized = (k[0] or '').strip().lower()
        if kanal_normalized:
            kanal_set.add(kanal_normalized)

    # Untuk tampilan chart: kapitalisasi tiap kata
    kanal_list = sorted(kanal_set)
    chart_labels = [k.title() for k in kanal_list]

    def get_data_by_range(start, end):
        q = Ticket.query.join(NomorTicket)\
            .filter(Ticket.kanal_pengaduan.isnot(None), Ticket.kanal_pengaduan != '')
        if start and end:
            q = q.filter(Ticket.tanggal >= start, Ticket.tanggal < end)

        result = q.with_entities(
            func.lower(func.trim(Ticket.kanal_pengaduan)).label('kanal'),
            func.count(distinct(Ticket.nomor_ticket_id))
        ).group_by('kanal').all()

        # Buat dict hasil
        data_dict = {kanal: count for kanal, count in result}
        return [data_dict.get(k, 0) for k in kanal_list]

    chart_series = []

    if not range1 and not range2:
        total_data = get_data_by_range(None, None)
        chart_series = [{
            "name": "Total",
            "data": total_data,
            "bucket_info": [[] for _ in total_data]
        }]
    else:
        data1 = get_data_by_range(start1, end1)
        data2 = get_data_by_range(start2, end2)
        if range1:
            chart_series.append({
                "name": f"Range {range1}",
                "data": data1,
                "bucket_info": [[] for _ in data1]
            })
        if range2:
            chart_series.append({
                "name": f"Range {range2}",
                "data": data2,
                "bucket_info": [[] for _ in data2]
            })

    kanal_colors = [
        "#3081D0", "#FF6768", "#00C49F", "#FFBB28", "#AF7AC5",
        "#2ECC71", "#F39C12", "#E74C3C", "#17A589", "#5D6D7E"
    ]
    default_colors = kanal_colors[:len(kanal_list)]

    return render_template(
        'filtering_kanal.html',
        chart_labels=chart_labels,   # pakai label kapitalisasi
        chart_series=chart_series,
        chart_colors=default_colors,
        range1=range1,
        range2=range2,
        user=current_user
    )

@app.route('/staff_dashboard')
@login_required
def staff_dashboard():
    if current_user.role != 'staff':
        flash('Akses ditolak: Anda bukan staff.')
        return redirect(url_for('admin_dashboard'))

    date_range = request.args.get('date_range')
    selected_os = request.args.getlist('os')
    selected_bucket = request.args.getlist('bucket')
    selected_jenis_pengaduan = request.args.getlist('jenis_pengaduan')
    chart_by = request.args.get('chart_by')

    start_date = end_date = None
    if date_range:
        try:
            start_str, end_str = date_range.split(' - ')
            start_date = datetime.strptime(start_str.strip(), '%Y-%m-%d')
            end_date = datetime.strptime(end_str.strip(), '%Y-%m-%d')
        except ValueError:
            flash("Format tanggal tidak valid.", "danger")

    all_os = [os[0] for os in db.session.query(Ticket.nama_os).distinct().all() if os[0]]
    all_bucket = [b[0] for b in db.session.query(Ticket.nama_bucket).distinct().all() if b[0]]

    os_filter = selected_os if selected_os else all_os
    bucket_filter = selected_bucket if selected_bucket else all_bucket
    jenis_pengaduan_filter = [int(j) for j in selected_jenis_pengaduan if j.strip().isdigit()]

    jenis_pengaduan_labels = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    if chart_by == 'jenis_pengaduan' and jenis_pengaduan_filter:
        chart_query = db.session.query(
            Ticket.jenis_pengaduan,
            func.count(Ticket.id)
        )

        if start_date and end_date:
            chart_query = chart_query.filter(Ticket.tanggal.between(start_date, end_date))
        if selected_os:
            chart_query = chart_query.filter(Ticket.nama_os.in_(selected_os))
        if selected_bucket:
            chart_query = chart_query.filter(Ticket.nama_bucket.in_(selected_bucket))
        if jenis_pengaduan_filter:
            chart_query = chart_query.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))

        chart_query = chart_query.group_by(Ticket.jenis_pengaduan)
        chart_data = chart_query.all()

        chart_labels = [
            jenis_pengaduan_labels.get(int(item[0]), f"Jenis {item[0]}") if item[0] else "Tidak Diketahui"
            for item in chart_data
        ]

        chart_values = [item[1] for item in chart_data]

        chart_series = [{
            "name": "Jumlah Order",
            "data": chart_values
        }]
        chart_title = "Jumlah Tiket Berdasarkan Jenis Pengaduan"

    else:
        from collections import defaultdict
        filter_by_bucket = bool(selected_bucket)

        if filter_by_bucket:
            chart_query = db.session.query(
                Ticket.nama_os,
                Ticket.nama_bucket,
                func.count(Ticket.id)
            ).filter(
                Ticket.nama_bucket.in_(bucket_filter),
                Ticket.nama_os.in_(os_filter)
            )

            if start_date and end_date:
                chart_query = chart_query.filter(Ticket.tanggal.between(start_date, end_date))
            if jenis_pengaduan_filter:
                chart_query = chart_query.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))

            chart_query = chart_query.group_by(Ticket.nama_bucket, Ticket.nama_os)
            chart_data = chart_query.all()

            all_os_labels = sorted(set([item[0] or "Tidak Diketahui" for item in chart_data]))
            grouped = defaultdict(lambda: defaultdict(int))

            for os_name, bucket, count in chart_data:
                os_label = os_name or "Tidak Diketahui"
                bucket_label = bucket or "Tidak Diketahui"
                grouped[bucket_label][os_label] = count

            chart_series = []
            for bucket_label in bucket_filter:
                bucket_label = bucket_label or "Tidak Diketahui"
                series_data = [grouped[bucket_label].get(os_label, 0) for os_label in all_os_labels]
                chart_series.append({
                    "name": bucket_label,
                    "data": series_data
                })

            chart_labels = all_os_labels
            chart_title = "Jumlah Tiket per OS berdasarkan Bucket"

        else:
            chart_query = db.session.query(
                Ticket.nama_os,
                func.count(Ticket.id)
            ).filter(Ticket.nama_os.in_(os_filter))

            if start_date and end_date:
                chart_query = chart_query.filter(Ticket.tanggal.between(start_date, end_date))
            if jenis_pengaduan_filter:
                chart_query = chart_query.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))

            chart_query = chart_query.group_by(Ticket.nama_os)
            chart_data = chart_query.all()

            chart_labels = [item[0] or "Tidak Diketahui" for item in chart_data]
            chart_values = [item[1] for item in chart_data]

            chart_series = [{
                "name": "Jumlah Order",
                "data": chart_values
            }]
            chart_title = "Jumlah Order per Tiket"

    query = db.session.query(NomorTicket).join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)

    if start_date and end_date:
        query = query.filter(Ticket.tanggal.between(start_date, end_date))
    if selected_os:
        query = query.filter(Ticket.nama_os.in_(selected_os))
    if selected_bucket:
        query = query.filter(Ticket.nama_bucket.in_(selected_bucket))
    if jenis_pengaduan_filter:
        query = query.filter(Ticket.jenis_pengaduan.in_(jenis_pengaduan_filter))

    total_nomor_ticket = query.distinct(NomorTicket.id).count()
    total_open = query.filter(or_(
        NomorTicket.status == 'aktif',
        NomorTicket.status == 'Reopen'
    )).distinct(NomorTicket.id).count()
    total_close = query.filter(NomorTicket.status == 'close').distinct(NomorTicket.id).count()

    # Tambahan: Chart jumlah NomorTicket berdasarkan jenis_pengaduan
    jenis_pengaduan_chart_query = db.session.query(
        Ticket.jenis_pengaduan,
        func.count(distinct(Ticket.nomor_ticket_id))
    ).join(NomorTicket, Ticket.nomor_ticket_id == NomorTicket.id).filter(
        Ticket.jenis_pengaduan.in_(range(1, 11))  
    )

    if start_date and end_date:
        jenis_pengaduan_chart_query = jenis_pengaduan_chart_query.filter(
            Ticket.tanggal.between(start_date, end_date)
        )

    jenis_pengaduan_chart_query = jenis_pengaduan_chart_query.group_by(Ticket.jenis_pengaduan)
    jenis_pengaduan_chart_data = jenis_pengaduan_chart_query.all()

    jenis_pengaduan_chart_labels = [item[0] for item in jenis_pengaduan_chart_data]

    jenis_pengaduan_chart_values = [item[1] for item in jenis_pengaduan_chart_data]

    jenis_pengaduan_chart_series = [{
        "name": "Jumlah NomorTicket",
        "data": jenis_pengaduan_chart_values
    }]

    return render_template(
        'staff_dashboard.html',
        user=current_user,
        total_nomor_ticket=total_nomor_ticket,
        total_open=total_open,
        total_close=total_close,
        selected_date_range=date_range or "Semua",
        selected_os=selected_os,
        selected_bucket=selected_bucket,
        selected_jenis_pengaduan=selected_jenis_pengaduan,
        all_os=all_os,
        all_bucket=all_bucket,
        chart_labels=chart_labels,
        chart_series=chart_series,
        chart_title=chart_title,
        jenis_pengaduan_chart_labels=jenis_pengaduan_chart_labels,
        jenis_pengaduan_chart_series=jenis_pengaduan_chart_series,
        ticket_chart_title="Jumlah Ticket per Status"
    )

@app.route('/pengaduan')
@login_required
def pengaduan():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')
    tahapan = request.args.get('tahapan')

    tahapan_options = [row[0] for row in db.session.query(Ticket.tahapan).distinct().all() if row[0]]

    nomor_ticket_query = NomorTicket.query.filter(
        NomorTicket.id_qc == None,
        NomorTicket.status.in_(['aktif', 'reopen'])
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla != 0)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass
        if tahapan:
            query = query.filter_by(tahapan=tahapan)

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]
    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(
            NomorTicket.status == 'aktif',
            Ticket.sla != 0,
            NomorTicket.id_qc == None
        )\
        .distinct()\
        .count()

    return render_template(
        'pengaduan.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif,
        tahapan_options=tahapan_options
    )

@app.route('/export-ticket-excel')
@login_required
def export_ticket_excel():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    date_range = request.args.get('date', '') 

    try:
        start_date_str, end_date_str = date_range.split(' - ')
        start_date = datetime.strptime(start_date_str.strip(), "%Y-%m-%d")
        end_date = datetime.strptime(end_date_str.strip(), "%Y-%m-%d")
    except ValueError:
        flash('Format tanggal tidak valid. Gunakan format: YYYY-MM-DD - YYYY-MM-DD', 'danger')
        return redirect(url_for('pengaduan'))

    tickets = Ticket.query \
        .filter(Ticket.tanggal >= start_date, Ticket.tanggal <= end_date) \
        .order_by(Ticket.tanggal.desc()).all()

    if not tickets:
        flash('Tidak ada data ticket pada rentang tanggal tersebut.', 'warning')
        return redirect(url_for('pengaduan'))

    status_ticket_map = {
        '1': 'Aktif',
        '2': 'Perpanjangan',
        '3': 'Keberatan',
        '4': 'Tutup',
        '5': 'Reopen'
    }

    jenis_pengaduan_map = {
        '1': "Informasi Pengajuan",
        '2': "Permintaan Kode OTP",
        '3': "Informasi Tenor",
        '4': "Informasi Tagihan",
        '5': "Informasi Denda",
        '6': "Pembatalan Pinjaman",
        '7': "Informasi Pencairan Dana",
        '8': "Perilaku Petugas Penagihan",
        '9': "Informasi Pembayaran",
        '10': "Discount / Pemutihan"
    }

    data = []
    for t in tickets:
        status_label = status_ticket_map.get(str(t.status_ticket), t.status_ticket)
        jenis_label = jenis_pengaduan_map.get(str(t.jenis_pengaduan), t.jenis_pengaduan)

        file_links = ''
        if t.bukti_chat:
            filenames = [f.strip() for f in t.bukti_chat.split(',') if f.strip()]
            base_url = request.host_url.rstrip('/') + '/static/uploads'
            file_links = ', '.join([f"{base_url}/{filename}" for filename in filenames])
        
        data.append({
            "Channel": t.kanal_pengaduan,
            "Tanggal": t.tanggal.strftime('%Y-%m-%d') if t.tanggal else '',
            "No Ticket": t.nomor_ticket.nomor_ticket if t.nomor_ticket else '',
            "Name": t.nama_nasabah,
            "Customer Phone Number": t.nomor_utama,
            "Email": t.email,
            "NIK": t.nik,
            "Detail Problem": t.detail_pengaduan,
            "Tipe Pengaduan": jenis_label,
            "Detail Pengaduan": t.detail_pengaduan,
            "Deskripsi Pengaduan": t.deskripsi_pengaduan,
            "Status Ticket": status_label,
            "DC": t.nama_dc,
            "OS": t.nama_os,
            "Bucket": t.nama_bucket,
            "Screenshoot Chat": file_links, 
        })

    df = pd.DataFrame(data)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Tickets')

    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=f"export_tickets_{start_date_str}_to_{end_date_str}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

@app.route('/nomor-ticket/<int:nomor_ticket_id>')
@login_required
def list_ticket_by_nomor(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }
    qc_users = User.query.filter_by(role='qc').all()

    return render_template(
        'list_ticket_by_nomor.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map,
        qc_users=qc_users
    )

@app.route('/eskalasi-ticket-qc/<int:nomor_ticket_id>')
@login_required
def eskalasi_ticket_qc(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()
    
    semua_kosong = all(ticket.deskripsi_qc in [None, ""] for ticket in tickets)

    if semua_kosong:
        status_feedback = "Belum ada Feedback"
        badge_feedback = "warning"
    else:
        status_feedback = "Ada Feedback"
        badge_feedback = "success"

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }
    qc_users = User.query.filter_by(role='qc').all()

    return render_template(
        'hasil_eskalasi.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map,
        qc_users=qc_users
    )

@app.route('/follow-up-pengaduan/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def follow_up_pengaduan(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    jenis_pengaduan = request.form.get("jenis_pengaduan")
    detail_pengaduan = request.form.get("detail_pengaduan")
    kronologis = request.form.get("kronologis")
    uploaded_files = request.files.getlist("bukti_chat")

    existing = request.form.getlist("existing_images")
    deleted = request.form.getlist("deleted_images")

    new_filenames = []
    for file in uploaded_files:
        if file and file.filename:
            filename = secure_filename(file.filename)
            save_path = os.path.join("static/uploads", filename)
            file.save(save_path)
            new_filenames.append(filename)

    final_files = [f for f in existing if f not in deleted] + new_filenames
    joined_filenames = ",".join(final_files)

    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket.id).all()

    for ticket in tickets:
        ticket.jenis_pengaduan = jenis_pengaduan
        ticket.detail_pengaduan = detail_pengaduan
        ticket.kronologis = kronologis
        ticket.bukti_chat = joined_filenames

    db.session.commit()
    flash("Data berhasil diupdate", "success")
    return redirect(url_for("list_ticket_by_nomor", nomor_ticket_id=nomor_ticket_id))

@app.route('/follow-up-pengaduan-reopen/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def follow_up_pengaduan_reopen(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    jenis_pengaduan = request.form.get("jenis_pengaduan")
    detail_pengaduan = request.form.get("detail_pengaduan")
    kronologis = request.form.get("kronologis")
    uploaded_files = request.files.getlist("bukti_chat")

    existing = request.form.getlist("existing_images")
    deleted = request.form.getlist("deleted_images")

    new_filenames = []
    for file in uploaded_files:
        if file and file.filename:
            filename = secure_filename(file.filename)
            save_path = os.path.join("static/uploads", filename)
            file.save(save_path)
            new_filenames.append(filename)

    final_files = [f for f in existing if f not in deleted] + new_filenames
    joined_filenames = ",".join(final_files)

    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket.id).all()

    for ticket in tickets:
        ticket.jenis_pengaduan = jenis_pengaduan
        ticket.detail_pengaduan = detail_pengaduan
        ticket.kronologis = kronologis
        ticket.bukti_chat = joined_filenames

    db.session.commit()
    flash("Data berhasil diupdate", "success")
    return redirect(url_for("list_reopen_ticket", nomor_ticket_id=nomor_ticket_id))

@app.route('/add-order/<int:ticket_id>', methods=['POST'])
@login_required
def add_order(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    original_ticket = Ticket.query.get_or_404(ticket_id)

    order_no = request.form.get('order_no')
    nama_os = request.form.get('nama_os')
    nama_dc = request.form.get('nama_dc')
    nama_bucket = request.form.get('nama_bucket')
    deskripsi_pengaduan = request.form.get('deskripsi_pengaduan')
    tanggal_str = request.form.get('tanggal')

    if not deskripsi_pengaduan or not tanggal_str:
        flash('Deskripsi pengaduan dan tanggal wajib diisi.', 'danger')
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=original_ticket.nomor_ticket_id))

    print("Tanggal dari form:", tanggal_str)

    try:
        tanggal = datetime.strptime(tanggal_str, '%Y-%m-%d')
    except Exception as e:
        flash(f'Tanggal tidak valid. {str(e)}', 'danger')
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=original_ticket.nomor_ticket_id))

    new_ticket = Ticket(
        order_no=order_no,
        nama_os=nama_os,
        nama_dc=nama_dc,
        nama_bucket=nama_bucket,
        deskripsi_pengaduan=deskripsi_pengaduan,
        tanggal=tanggal,  

        kanal_pengaduan=original_ticket.kanal_pengaduan,
        kategori_pengaduan=original_ticket.kategori_pengaduan,
        jenis_pengaduan=original_ticket.jenis_pengaduan,
        detail_pengaduan=original_ticket.detail_pengaduan,
        nama_nasabah=original_ticket.nama_nasabah,
        email=original_ticket.email,
        nomor_utama=original_ticket.nomor_utama,
        nomor_kontak=original_ticket.nomor_kontak,
        nik=original_ticket.nik,

        input_by=current_user.id,
        sla=10,
        status_ticket='1',
        nomor_ticket_id=original_ticket.nomor_ticket_id
    )

    db.session.add(new_ticket)
    db.session.commit()

    flash('Order berhasil ditambahkan.', 'success')
    return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=original_ticket.nomor_ticket_id))

@app.route('/add-order-reopen/<int:ticket_id>', methods=['POST'])
@login_required
def add_order_reopen(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    original_ticket = Ticket.query.get_or_404(ticket_id)

    order_no = request.form.get('order_no')
    nama_os = request.form.get('nama_os')
    nama_dc = request.form.get('nama_dc')
    nama_bucket = request.form.get('nama_bucket')
    deskripsi_pengaduan = request.form.get('deskripsi_pengaduan')
    tanggal_str = request.form.get('tanggal')

    if not deskripsi_pengaduan or not tanggal_str:
        flash('Deskripsi pengaduan dan tanggal wajib diisi.', 'danger')
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=original_ticket.nomor_ticket_id))

    print("Tanggal dari form:", tanggal_str)

    try:
        tanggal = datetime.strptime(tanggal_str, '%Y-%m-%d')
    except Exception as e:
        flash(f'Tanggal tidak valid. {str(e)}', 'danger')
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=original_ticket.nomor_ticket_id))

    new_ticket = Ticket(
        order_no=order_no,
        nama_os=nama_os,
        nama_dc=nama_dc,
        nama_bucket=nama_bucket,
        deskripsi_pengaduan=deskripsi_pengaduan,
        tanggal=tanggal,  

        kanal_pengaduan=original_ticket.kanal_pengaduan,
        kategori_pengaduan=original_ticket.kategori_pengaduan,
        jenis_pengaduan=original_ticket.jenis_pengaduan,
        detail_pengaduan=original_ticket.detail_pengaduan,
        nama_nasabah=original_ticket.nama_nasabah,
        email=original_ticket.email,
        nomor_utama=original_ticket.nomor_utama,
        nomor_kontak=original_ticket.nomor_kontak,
        nik=original_ticket.nik,

        input_by=current_user.id,
        sla=10,
        status_ticket='5',
        nomor_ticket_id=original_ticket.nomor_ticket_id
    )

    db.session.add(new_ticket)
    db.session.commit()

    flash('Order berhasil ditambahkan.', 'success')
    return redirect(url_for('list_reopen_ticket', nomor_ticket_id=original_ticket.nomor_ticket_id))

@app.route('/add-kontak/<int:ticket_id>', methods=['POST'])
@login_required
def add_kontak(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    ticket = Ticket.query.get_or_404(ticket_id)

    nama_lengkap = request.form.get('nama_lengkap')
    nik = request.form.get('nik')
    phone = request.form.get('phone')
    phone_2 = request.form.get('phone_2')
    email = request.form.get('email')

    if not all([nama_lengkap, nik, phone]):
        flash('Field wajib diisi: Nama, NIK, dan No HP.', 'danger')
        return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=ticket.nomor_ticket_id))

    kontak = Kontak(
        nama_lengkap=nama_lengkap,
        nik=nik,
        phone=phone,
        phone_2=phone_2,
        email=email,
        id_ticket=ticket.id
    )

    db.session.add(kontak)
    db.session.commit()

    flash('Kontak berhasil ditambahkan.', 'success')
    return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=ticket.nomor_ticket_id))

@app.route('/submit-ticket', methods=['POST'])
@login_required
def submit_ticket():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    try:
        nomor_ticket_value = request.form.get('nomor_ticket')

        nomor_ticket = NomorTicket.query.filter_by(nomor_ticket=nomor_ticket_value).first()
        if not nomor_ticket:
            nomor_ticket = NomorTicket(nomor_ticket=nomor_ticket_value)
            db.session.add(nomor_ticket)
            db.session.flush()  

        ticket = Ticket(
            kanal_pengaduan=request.form.get('country'),
            kategori_pengaduan=request.form.get('kategori'),
            jenis_pengaduan=request.form.get('jenis'),
            detail_pengaduan=request.form.get('detail_pengaduan'),
            tanggal=datetime.strptime(request.form.get('tanggal'), "%Y-%m-%d") if request.form.get('tanggal') else datetime.utcnow(),
            nama_nasabah=request.form.get('nama_nasabah'),
            email=request.form.get('email'),
            nomor_utama=request.form.get('nomor_utama'),
            nomor_kontak=request.form.get('nomor_kontak'),
            nik=request.form.get('nik'),
            nama_os=request.form.get('nama_os').replace(" ", "") if request.form.get('nama_os') else None,
            nama_dc=request.form.get('nama_dc'),
            nama_bucket=request.form.get('nama_bucket').replace(" ", "") if request.form.get('nama_bucket') else None,
            order_no=request.form.get('order_no'),
            deskripsi_pengaduan=request.form.get('deskripsi_pengaduan'),
            input_by=current_user.id,
            sla=10,
            status_ticket='1',
            nomor_ticket=nomor_ticket, 
            created_time=datetime.utcnow()
        )

        db.session.add(ticket)
        db.session.commit()

        flash('Ticket berhasil ditambahkan!', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Terjadi kesalahan: {e}', 'danger')

    return redirect(url_for('pengaduan',))

@app.route('/update-tahapan/<int:nomor_ticket_id>/<int:ticket_id>', methods=['POST'])
@login_required
def update_tahapan(nomor_ticket_id, ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    id_qc = request.form.get('id_qc')
    
    tiket = Ticket.query.get_or_404(ticket_id)

    tahapan = request.form.get('tahapan')
    status_ticket = request.form.get('status_ticket')
    nama_os = request.form.get('nama_os').strip() if request.form.get('nama_os') else None
    nama_bucket = request.form.get('nama_bucket').strip() if request.form.get('nama_bucket') else None
    nama_dc = request.form.get('nama_dc')
    nama_nasabah = request.form.get('nama_nasabah')
    nik = request.form.get('nik')
    nomor_utama = request.form.get('nomor_utama')
    nomor_kontak = request.form.get('nomor_kontak')
    email = request.form.get('email')
    deskripsi_pengaduan = request.form.get('deskripsi_pengaduan')
    order_no = request.form.get('order_no')

    tahapan_2 = None
    if status_ticket == '3':
        date = request.form.get('tahapan_2_date')
        desc = request.form.get('tahapan_2_desc')
        tahapan_2 = f"{date} - {desc}" if date and desc else None
    elif status_ticket == '4':
        followup = request.form.get('tahapan_2_followup')
        tahapan_2 = followup if followup else None

    is_updating_tahapan = bool(status_ticket or tahapan or tahapan_2)

    if is_updating_tahapan:
        tiket.tahapan = tahapan 
        tiket.status_ticket = status_ticket
        tiket.tahapan_2 = tahapan_2

        if tahapan == "Eskalasi ke QC" and id_qc:
            tiket.nomor_ticket.id_qc = int(id_qc)

        new_history = History(
            nomor_ticket=tiket.nomor_ticket.nomor_ticket,
            order_number=tiket.order_no,
            status_ticket=status_ticket,
            tahapan=tahapan,
            create_by=current_user.id,
            nama_os=nama_os
        )
        db.session.add(new_history)

    tiket.nama_os = nama_os
    tiket.nama_bucket = nama_bucket
    tiket.nama_dc = nama_dc
    tiket.nama_nasabah = nama_nasabah
    tiket.nik = nik
    tiket.nomor_utama = nomor_utama
    tiket.nomor_kontak = nomor_kontak
    tiket.email = email
    tiket.deskripsi_pengaduan = deskripsi_pengaduan
    tiket.order_no = order_no

    db.session.commit()

    flash('Data berhasil diperbarui.', 'success')
    return redirect(url_for('list_ticket_by_nomor', nomor_ticket_id=nomor_ticket_id))

@app.route('/update-catatan/<int:ticket_id>', methods=['POST'])
@login_required
def update_catatan(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    tiket = Ticket.query.get_or_404(ticket_id)
    
    catatan = request.form.get('catatan')
    
    if catatan:
        tiket.catatan = catatan
        tiket.tanggal_catatan = datetime.today().strftime('%Y-%m-%d')
        db.session.commit()
        flash('Catatan berhasil disimpan.', 'success')
    else:
        flash('Catatan tidak boleh kosong.', 'danger')

    return redirect(request.referrer or url_for('list_ticket_by_nomor', nomor_ticket_id=tiket.nomor_ticket_id))

@app.route('/mark-case-valid/<int:ticket_id>', methods=['POST'])
@login_required
def mark_case_valid(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    tiket = Ticket.query.get_or_404(ticket_id)
    tiket.status_case = 'valid'
    db.session.commit()
    flash('Status case diubah menjadi VALID', 'success')
    return redirect(request.referrer or url_for('dashboard'))

@app.route('/update-tahapan-reopen/<int:nomor_ticket_id>/<int:ticket_id>', methods=['POST'])
@login_required
def update_tahapan_reopen(nomor_ticket_id, ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    tiket = Ticket.query.get_or_404(ticket_id)

    tahapan = request.form.get('tahapan')
    status_ticket = request.form.get('status_ticket')
    tahapan_2 = None

    if status_ticket == '3':
        date = request.form.get('tahapan_2_date')
        desc = request.form.get('tahapan_2_desc')
        tahapan_2 = f"{date} - {desc}" if date and desc else None
    elif status_ticket == '4':
        followup = request.form.get('tahapan_2_followup')
        tahapan_2 = followup if followup else None

    if not tahapan:
        flash('Tahapan wajib dipilih.', 'danger')
        return redirect(url_for('list_reopen_ticket', nomor_ticket_id=nomor_ticket_id))

    tiket.tahapan = tahapan
    tiket.status_ticket = status_ticket
    tiket.tahapan_2 = tahapan_2
    db.session.commit()

    new_history = History(
        nomor_ticket=tiket.nomor_ticket.nomor_ticket,
        order_number=tiket.order_no,
        status_ticket=status_ticket,
        tahapan=tahapan,
        create_by=current_user.id
    )
    db.session.add(new_history)
    db.session.commit()

    flash('Data berhasil diperbarui.', 'success')
    return redirect(url_for('list_reopen_ticket', nomor_ticket_id=nomor_ticket_id))

@app.route('/close-nomor-ticket/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def close_nomor_ticket(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    
    nomor_ticket.status = 'close'

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket.id).all()
    for ticket in tickets:
        ticket.status_ticket = '4'

    db.session.commit()
    flash("Nomor ticket berhasil ditutup.", "success")
    return redirect(url_for('close_ticket', nomor_ticket_id=nomor_ticket.id))

@app.route('/reopen-nomor-ticket/<int:nomor_ticket_id>', methods=['POST'])
@login_required
def reopen_nomor_ticket(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)
    
    nomor_ticket.status = 'reopen'

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket.id).all()
    for ticket in tickets:
        ticket.status_ticket = '5'

    db.session.commit()
    flash("Nomor ticket berhasil diubah menjadi Reopen.", "success")

    return redirect(url_for('reopen_ticket', nomor_ticket_id=nomor_ticket.id))

@app.route('/ticket-close')
@login_required
def close_ticket():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')

    nomor_ticket_query = NomorTicket.query.filter(NomorTicket.status == 'close')

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_close = NomorTicket.query.filter_by(status='close').count()

    return render_template(
        'ticket_close.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_close=jumlah_tiket_close
    )

@app.route('/closed-ticket/<int:nomor_ticket_id>')
@login_required
def list_closed_ticket(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }

    return render_template(
        'list_closed_ticket.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map
    )

@app.route('/reopen-ticket/<int:nomor_ticket_id>')
@login_required
def list_reopen_ticket(nomor_ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    nomor_ticket = NomorTicket.query.get_or_404(nomor_ticket_id)

    tickets = Ticket.query.filter_by(nomor_ticket_id=nomor_ticket_id)\
        .order_by(Ticket.created_time.asc()).all()

    jenis_pengaduan_map = {
        1: "Informasi Pengajuan",
        2: "Permintaan Kode OTP",
        3: "Informasi Tenor",
        4: "Informasi Tagihan",
        5: "Informasi Denda",
        6: "Pembatalan Pinjaman",
        7: "Informasi Pencairan Dana",
        8: "Perilaku Petugas Penagihan",
        9: "Informasi Pembayaran",
        10: "Discount / Pemutihan"
    }

    detail_pengaduan_map = {
        1: [
            "Hasil Pengajuan",
            "Pengajuan Ditolak",
            "Status Pengajuan sedang ditransfer",
            "Tidak bisa pengajuan ulang karena keterlambatan",
            "Verifikasi Bank gagal",
            "Verifikasi KTP gagal",
            "Cara pengajuan",
            "Perubahan Nomor Handphone",
            "Perubahan Nomor Rekening"
        ],
        2: [
            "OTP Limit",
            "Tidak terima SMS OTP"
        ],
        3: [
            "Informasi Pinjaman"
        ],
        4: [
            "Konsultasi detail pinjaman saat ini",
            "Konsultasi Perpanjangan",
            "Bukti Transfer"
        ],
        5: [
            "Denda Keterlambatan"
        ],
        6: [
            "Hapus Data (Penutupan Akun)",
            "Pembatalan Pinjaman"
        ],
        7: [
            "Pencairan dana berhasil",
            "Status pencairan dana gagal",
            "Status pengajuan pencairan dana ulang",
            "Tidak terima dana",
            "Operasi Gagal (tidak bisa verifikasi wajah dan KTP)"
        ],
        8: [
            "Keluhan Penagihan",
            "Keluhan Reminder",
            "Penipuan"
        ],
        9: [
            "Konfirmasi Pembayaran",
            "Pembayaran belum masuk",
            "Pembayaran bukan ke VA UATAS",
            "Refund (pembayaran double)",
            "Meminta VA (cicilan)",
            "Meminta VA (pelunasan)",
            "Meminta VA (perpanjangan)",
            "Tidak bisa ambil VA"
        ],
        10: [
            "Meminta keringanan pembayaran (cicilan)",
            "Meminta keringanan pembayaran (potongan denda)",
            "Tidak ada dana"
        ]
    }

    return render_template(
        'list_reopen_ticket.html',
        nomor_ticket=nomor_ticket,
        tickets=tickets,
        user=current_user,
        jenis_pengaduan_map=jenis_pengaduan_map,
        detail_pengaduan_map=detail_pengaduan_map
    )

@app.route('/reopen-ticket')
@login_required
def reopen_ticket():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')

    nomor_ticket_query = NomorTicket.query.filter(NomorTicket.status == 'reopen')

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_reopen = NomorTicket.query.filter_by(status='reopen').count()

    return render_template(
        'reopen_ticket.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_reopen=jumlah_tiket_reopen
    )

@app.route('/download-template')
def download_template():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    return send_from_directory(directory='static/files', path='template_cs.xlsx', as_attachment=True)

import re

def clean_alpha_only(val):
    cleaned = re.sub(r"[^a-zA-Z]", "", val) 
    return cleaned if cleaned else None

@app.route('/upload', methods=['POST'])
@login_required
def upload_excel():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    def safe_val(val):
        return None if pd.isna(val) else str(val).strip()

    try:
        file = request.files.get('avatar')  
        if not file:
            flash("Tidak ada file yang diupload", 'danger')
            return redirect(request.referrer)

        df = pd.read_excel(file)

        expected_cols = ['kanal_pengaduan','nomor_ticket', 'tanggal', 'nama_nasabah', 'tipe_pengaduan',
                         'detail_pengaduan', 'order_no', 'os', 'dc', 'bucket']
        if not all(col in df.columns for col in expected_cols):
            flash("Kolom Excel tidak sesuai template.", 'danger')
            return redirect(request.referrer)

        jenis_pengaduan_map = {
            "Informasi Pengajuan": 1,
            "Permintaan Kode OTP": 2,
            "Informasi Tenor": 3,
            "Informasi Tagihan": 4,
            "Informasi Denda": 5,
            "Pembatalan Pinjaman": 6,
            "Informasi Pencairan Dana": 7,
            "Perilaku Petugas Penagihan": 8,
            "Informasi Pembayaran": 9,
            "Discount / Pemutihan": 10
        }

        existing_order_nos = {ticket.order_no for ticket in Ticket.query.with_entities(Ticket.order_no).all()}

        inserted_order_nos = set()

        for index, row in df.iterrows():
            order_no = safe_val(row['order_no'])

            if not order_no or order_no in existing_order_nos or order_no in inserted_order_nos:
                continue

            inserted_order_nos.add(order_no)

            nomor_ticket_str = safe_val(row['nomor_ticket'])
            nomor_ticket = NomorTicket.query.filter_by(nomor_ticket=nomor_ticket_str).first()
            if not nomor_ticket:
                nomor_ticket = NomorTicket(nomor_ticket=nomor_ticket_str)
                db.session.add(nomor_ticket)
                db.session.flush()

            tanggal_value = row['tanggal']
            if isinstance(tanggal_value, str):
                tanggal_value = datetime.strptime(tanggal_value, '%Y-%m-%d')
            elif pd.isna(tanggal_value):
                tanggal_value = datetime.utcnow()

            jenis_pengaduan_str = safe_val(row['tipe_pengaduan'])
            jenis_pengaduan_val = jenis_pengaduan_map.get(jenis_pengaduan_str)
            if not jenis_pengaduan_val:
                raise ValueError(f"Jenis pengaduan tidak valid di baris {index + 2}: '{jenis_pengaduan_str}'")

            ticket = Ticket(
                kanal_pengaduan=safe_val(row['kanal_pengaduan']),
                nomor_ticket=nomor_ticket,
                tanggal=tanggal_value,
                nama_nasabah=safe_val(row['nama_nasabah']),
                jenis_pengaduan=jenis_pengaduan_val,
                detail_pengaduan=safe_val(row['detail_pengaduan']),
                order_no=order_no,
                nama_os=clean_alpha_only(safe_val(row['os']).replace(" ", "")) if safe_val(row['os']) and pd.notna(row['os']) else None,
                nama_dc=safe_val(row['dc']),
                nama_bucket=safe_val(row['bucket']).replace(" ", "") if safe_val(row['bucket']) and pd.notna(row['bucket']) else None,
                input_by=current_user.id,
                sla=10,
                status_ticket='1',
                created_time=datetime.utcnow(),
            )

            db.session.add(ticket)

        db.session.commit()
        flash(f"Berhasil mengimport data dari Excel. {len(inserted_order_nos)} data baru ditambahkan.", 'success')

    except Exception as e:
        db.session.rollback()
        flash(f"Gagal import: {e}", 'danger')

    return redirect(request.referrer)

@app.route("/case-valid")
@login_required
def case_valid():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')

    query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
        .filter(Ticket.status_case == 'valid')

    if jenis:
        query = query.filter(Ticket.jenis_pengaduan == jenis)
    if status:
        query = query.filter(Ticket.status_ticket == status)
    if tanggal:
        try:
            tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
            query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
        except ValueError:
            pass

    query = query.order_by(Ticket.created_time.desc())

    per_page = 10
    pagination = query.paginate(page=page, per_page=per_page, error_out=False)

    return render_template(
        "case_valid.html",
        user=current_user,
        tickets=pagination
    )

@app.route('/upload-document/<int:ticket_id>', methods=['POST'])
@login_required
def upload_document(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    ticket = Ticket.query.get_or_404(ticket_id)
    files = request.files.getlist('documents')
    uploaded_files = []

    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)

            if os.path.exists(filepath):
                name, ext = os.path.splitext(filename)
                filename = f"{name}_{datetime.utcnow().timestamp()}{ext}"
                filepath = os.path.join(UPLOAD_FOLDER, filename)

            file.save(filepath)
            uploaded_files.append(filename)

    if ticket.document:
        existing_files = ticket.document.split(',')
        all_files = existing_files + uploaded_files
    else:
        all_files = uploaded_files

    ticket.document = ','.join(all_files)
    db.session.commit()

    flash(f'{len(uploaded_files)} dokumen berhasil diupload.', 'success')
    return redirect(request.referrer or url_for('case_valid'))

@app.route('/hapus-dokumen/<int:ticket_id>', methods=['POST'])
@login_required
def hapus_dokumen(ticket_id):
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    filename = request.form.get('filename')
    ticket = Ticket.query.get_or_404(ticket_id)

    if not filename or filename not in (ticket.document or ''):
        flash('File tidak ditemukan atau tidak valid.', 'danger')
        return redirect(request.referrer)

    file_path = os.path.join('static/uploads', filename)
    if os.path.exists(file_path):
        os.remove(file_path)

    dokumen_list = ticket.document.split(',')
    dokumen_list.remove(filename)
    ticket.document = ','.join(dokumen_list)
    db.session.commit()

    flash(f'File {filename} berhasil dihapus.', 'success')
    return redirect(request.referrer)

@app.route('/sla')
@login_required
def sla():
    if current_user.role != 'staff':
        return redirect(request.referrer)
    
    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(
        or_(
            NomorTicket.status == None,
            and_(
                NomorTicket.status != 'close',
                NomorTicket.status != 'reopen'
            )
        )
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla == 0)  # Hanya SLA = 0

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).filter(Ticket.sla == 0)  # Filter sesuai SLA
        .group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(NomorTicket.status == 'aktif', Ticket.sla == 0)\
        .distinct()\
        .count()

    return render_template(
        'sla.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif
    )

@app.route('/add-detail-qc/<int:ticket_id>', methods=['POST'])
@login_required
def add_detail_qc(ticket_id):
    if current_user.role != 'qc':
        return redirect(request.referrer)

    tiket = Ticket.query.get_or_404(ticket_id)

    deskripsi_qc = request.form.get('deskripsi_qc')

    uploaded_files = request.files.getlist('file_qc')
    existing_files = request.form.getlist('existing_images') 

    filenames = existing_files.copy()

    for file in uploaded_files:
        if file and file.filename:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.root_path, 'static/uploads', filename)
            file.save(filepath)
            filenames.append(filename)

    tiket.deskripsi_qc = deskripsi_qc
    tiket.file_qc = ','.join(filenames) if filenames else None

    db.session.commit()
    flash('Detail QC berhasil disimpan.', 'success')
    return redirect(request.referrer)

@app.route('/eskalasi-qc')
@login_required
def eskalasi_qc():
    if current_user.role != 'staff':
        return redirect(request.referrer)

    page = request.args.get('page', 1, type=int)
    jenis = request.args.get('jenis')
    status = request.args.get('status')
    tanggal = request.args.get('tanggal')
    q = request.args.get('q')

    nomor_ticket_query = NomorTicket.query.filter(
        NomorTicket.id_qc.isnot(None), 
        or_(
            NomorTicket.status == None,
            and_(
                NomorTicket.status != 'close',
                NomorTicket.status != 'reopen'
            )
        )
    )

    if q:
        matching_ids_from_nama = db.session.query(Ticket.nomor_ticket_id)\
            .filter(Ticket.nama_nasabah.ilike(f"%{q}%"))\
            .distinct().all()
        matching_ids_from_nama = [id[0] for id in matching_ids_from_nama]

        matching_ids_from_nomor = db.session.query(NomorTicket.id)\
            .filter(NomorTicket.nomor_ticket.ilike(f"%{q}%"))\
            .all()
        matching_ids_from_nomor = [id[0] for id in matching_ids_from_nomor]

        all_matching_ids = list(set(matching_ids_from_nama + matching_ids_from_nomor))

        nomor_ticket_query = nomor_ticket_query.filter(NomorTicket.id.in_(all_matching_ids))

    tickets_grouped = []

    for nt in nomor_ticket_query.all():
        query = Ticket.query.options(joinedload(Ticket.nomor_ticket))\
            .filter(Ticket.nomor_ticket_id == nt.id, Ticket.sla != 0)

        if jenis:
            query = query.filter_by(jenis_pengaduan=jenis)
        if status:
            query = query.filter_by(status_ticket=status)
        if tanggal:
            try:
                tanggal_obj = datetime.strptime(tanggal, "%Y-%m-%d")
                query = query.filter(func.date(Ticket.tanggal) == tanggal_obj.date())
            except ValueError:
                pass

        first_ticket = query.order_by(Ticket.created_time.asc()).first()
        if first_ticket:
            ada_feedback_qc = db.session.query(Ticket.id).filter(
                Ticket.nomor_ticket_id == nt.id,
                or_(
                    Ticket.deskripsi_qc.isnot(None),
                    Ticket.file_qc.isnot(None)
                )
            ).first() is not None

            first_ticket.feedback_qc_status = "Check Feedback QC" if ada_feedback_qc else "Belum ada Feedback QC"
            first_ticket.feedback_qc_badge = "success" if ada_feedback_qc else "warning"

            tickets_grouped.append(first_ticket)

    tickets_grouped.sort(
        key=lambda x: x.created_time or datetime.min,
        reverse=True
    )

    per_page = 10
    total = len(tickets_grouped)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_items = tickets_grouped[start:end]

    class Pagination:
        def __init__(self, items, page, per_page, total):
            self.items = items
            self.page = page
            self.per_page = per_page
            self.total = total
            self.pages = (total + per_page - 1) // per_page
            self.has_prev = page > 1
            self.has_next = page < self.pages
            self.prev_num = page - 1
            self.next_num = page + 1

    pagination = Pagination(paginated_items, page, per_page, total)

    count_by_nomor_ticket = dict(
        db.session.query(
            Ticket.nomor_ticket_id,
            func.count(Ticket.id)
        ).group_by(Ticket.nomor_ticket_id).all()
    )

    jumlah_tiket_aktif = db.session.query(NomorTicket)\
        .join(Ticket, Ticket.nomor_ticket_id == NomorTicket.id)\
        .filter(NomorTicket.status == 'aktif', Ticket.sla != 0, NomorTicket.id_qc.isnot(None))\
        .distinct()\
        .count()

    return render_template(
        'eskalasi.html',
        user=current_user,
        tickets=pagination,
        count_by_nomor_ticket=count_by_nomor_ticket,
        jumlah_tiket_aktif=jumlah_tiket_aktif
    )

if __name__ == '__main__':
    if not is_running_from_reloader():
        if not scheduler.running:
            scheduler.start()
    with app.app_context():
        db.create_all()
    app.run(debug=True, port=5007, host='0.0.0.0')