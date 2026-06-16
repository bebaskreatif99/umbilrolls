from flask import Flask, render_template, request, redirect, url_for, flash, Response, session
import csv
import os
import io
import re
from datetime import datetime, timedelta
from functools import wraps

app = Flask(__name__)
app.secret_key = 'umbirolls_super_secret_key_2026'

# --- 1. PENGATURAN PATH FILE ANTI-ERROR ---
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
FILE_COSTING = os.path.join(BASE_DIR, 'Umbirolls drizzle AI.xlsx - Costing.csv')
FILE_PROSES = os.path.join(BASE_DIR, 'Umbirolls drizzle AI.xlsx - Data proses.csv')

# --- 2. KONFIGURASI GLOBAL & DATABASE ---
APP_CONFIG = {
    'margin_target': 17.0, 'overhead_tetap': 21404, 'jam_kerja_hari': 8.0,
    'pembulatan_kelipatan': 500, 'mesin_kukus': 1, 'mesin_goreng': 1, 'user_role': 'Owner'
}
INVENTORY = {'ubi_ungu_gr': 5000, 'kulit_lumpia_lbr': 500}

waktu_sekarang = datetime.now()
TRANSAKSI = [
    {
        'id': 'TRX-001', 
        'tanggal': waktu_sekarang.strftime('%Y-%m-%d %H:%M'), 
        'jumlah': 15, 
        'rasa': {'Coklat': 5, 'Stroberi': 5, 'Tiramissu': 5, 'Keju': 0, 'Oreo': 0}, 
        'modal': 97215, 'pendapatan': 120000, 'laba': 22785, 
        'estimasi_teks': '15 Menit',
        'target_waktu_raw': (waktu_sekarang + timedelta(minutes=15)).strftime('%Y-%m-%d %H:%M:%S'),
        'target_waktu_tampil': (waktu_sekarang + timedelta(minutes=15)).strftime('%H:%M WIB'),
        'status_produksi': 'Sukses',
        'status_pembayaran': 'Lunas'
    }
]

# --- 3. DEKORATOR PENGAMAN LOGIN ---
def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            flash("🔒 Silakan login terlebih dahulu untuk mengakses sistem.", "error")
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

# --- 4. FUNGSI LOAD & SAVE DATA CSV ---
def load_costing():
    data = []
    if os.path.exists(FILE_COSTING):
        try:
            with open(FILE_COSTING, newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                start = False
                for row in reader:
                    if not row or all(c.strip() == "" for c in row): continue
                    if "Nama barang" in row[0] or "Komponen" in row[0]:
                        start = True
                        continue
                    if start:
                        if "TOTAL" in row[0].upper() or "HPP" in row[0].upper(): break
                        if row[0].strip():
                            data.append({
                                'nama': row[0].strip(), 'harga_satuan': row[2].strip() if len(row) > 2 else "0",
                                'kebutuhan': row[3].strip() if len(row) > 3 else "0", 'total': row[4].strip() if len(row) > 4 else "0"
                            })
        except: pass
    if not data:
        data = [
            {'nama': 'Ubi ungu', 'harga_satuan': '25000', 'kebutuhan': '600gr', 'total': '15000'},
            {'nama': 'Kulit lumpia', 'harga_satuan': '28000', 'kebutuhan': '75 lbr', 'total': '26250'},
            {'nama': 'Minyak goreng', 'harga_satuan': '6000', 'kebutuhan': '100mL', 'total': '3000'},
            {'nama': 'Keju gold cheddar', 'harga_satuan': '23000', 'kebutuhan': '20gr', 'total': '2875'},
            {'nama': 'Oreo crumble', 'harga_satuan': '23000', 'kebutuhan': '20gr', 'total': '1840'},
            {'nama': 'Coklat', 'harga_satuan': '20000', 'kebutuhan': '45gr', 'total': '4500'},
            {'nama': 'Stroberi', 'harga_satuan': '20000', 'kebutuhan': '45gr', 'total': '4500'},
            {'nama': 'tiramissu', 'harga_satuan': '20000', 'kebutuhan': '45gr', 'total': '4500'},
            {'nama': 'Piping bag', 'harga_satuan': '12000', 'kebutuhan': '4 lbr', 'total': '480'},
            {'nama': 'Kemasan', 'harga_satuan': '800', 'kebutuhan': '15 kotak', 'total': '12000'}
        ]
    return data

def save_costing(form):
    names = form.getlist('nama[]')
    hargas = form.getlist('harga_satuan[]')
    kebs = form.getlist('kebutuhan[]')
    totals = form.getlist('total[]')
    with open(FILE_COSTING, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(["Nama barang", "Satuan", "Harga Satuan", "Kebutuhan", "Total"])
        for i in range(len(names)):
            if names[i].strip(): writer.writerow([names[i], "-", hargas[i], kebs[i], totals[i]])

def load_proses():
    data = []
    if os.path.exists(FILE_PROSES):
        try:
            with open(FILE_PROSES, newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                start = False
                for row in reader:
                    if not row or all(c.strip() == "" for c in row): continue
                    if "Nama Stasiun" in row[0] or "Stasiun" in row[0]:
                        start = True
                        continue
                    if start:
                        if "HPP" in row[0].upper(): break
                        if row[0].strip():
                            data.append({
                                'stasiun': row[0].strip(), 'kelompok': row[1].strip() if len(row) > 1 else "-",
                                'waktu': row[2].strip() if len(row) > 2 else "0", 'reject': row[4].strip() if len(row) > 4 else "0",
                                'status': row[8].strip().upper() if len(row) > 8 else "OPTIMAL"
                            })
        except: pass
    if not data:
        data = [
            {'stasiun': 'Steaming', 'kelompok': 'Pre_Process', 'waktu': '900', 'reject': '3', 'status': 'BOTTLENECK'},
            {'stasiun': 'Frying', 'kelompok': 'Main', 'waktu': '600', 'reject': '3', 'status': 'OPTIMAL'}
        ]
    return data

def save_proses(form):
    stasiuns = form.getlist('stasiun[]')
    kelompoks = form.getlist('kelompok[]')
    waktus = form.getlist('waktu[]')
    rejects = form.getlist('reject[]')
    statuses = form.getlist('status[]')
    with open(FILE_PROSES, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(["Nama Stasiun", "Kelompok", "Cycle Time", "Efisiensi", "Reject", "Dt", "1", "2", "Status"])
        for i in range(len(stasiuns)):
            if stasiuns[i].strip(): writer.writerow([stasiuns[i], kelompoks[i], waktus[i], "-", rejects[i], "-", "-", "-", statuses[i]])

# --- 5. LOGIKA KALKULASI FINANSIAL & OPERASIONAL ---
def hitung_hpp_dan_harga():
    costing_data = load_costing()
    total_biaya_batch = 0
    for item in costing_data:
        try:
            val = re.sub(r'[^\d]', '', str(item['total']))
            if val: total_biaya_batch += float(val)
        except: pass
    base_hpp_per_pack = total_biaya_batch / 15 if total_biaya_batch > 0 else 6481
    margin_decimal = APP_CONFIG['margin_target'] / 100.0
    if margin_decimal >= 1.0: margin_decimal = 0.99
    raw_harga_jual = base_hpp_per_pack / (1 - margin_decimal)
    kelipatan = APP_CONFIG['pembulatan_kelipatan']
    harga_jual = int((raw_harga_jual + kelipatan - 1) // kelipatan * kelipatan)
    return int(base_hpp_per_pack), harga_jual

def hitung_durasi_produksi_jam(jumlah_pesanan):
    # REVISI TETAP TERJAGA: Kapasitas harian diset 1085 pack (35 batch per hari)
    standar_batch_per_hari = 35.0
    standar_jam_kerja = 8.0
    pack_per_batch = 31.0
    faktor_mesin = (APP_CONFIG['mesin_kukus'] + APP_CONFIG['mesin_goreng']) / 2.0
    batch_per_jam = (standar_batch_per_hari / standar_jam_kerja) * faktor_mesin
    pack_per_jam = batch_per_jam * pack_per_batch
    total_jam_dibutuhkan = jumlah_pesanan / pack_per_jam
    return total_jam_dibutuhkan

def hitung_kapasitas_harian():
    return int((35.0 * 31.0) * ((APP_CONFIG['mesin_kukus'] + APP_CONFIG['mesin_goreng']) / 2.0) * (APP_CONFIG['jam_kerja_hari'] / 8.0))

# --- 6. ROUTING AUTENTIKASI ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    if 'logged_in' in session: return redirect(url_for('index'))
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        if username == 'owner' and password == 'umbirolls123':
            session['logged_in'] = True
            session['username'] = username
            flash("🔑 Login berhasil! Selamat datang di Drizzle System.", "success")
            return redirect(url_for('index'))
        else:
            flash("Username atau password salah!", "error")
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    session.pop('username', None)
    flash("👋 Anda telah berhasil keluar dari sistem.", "success")
    return redirect(url_for('login'))

# --- 7. ROUTING UTAMA ---
@app.route('/', methods=['GET', 'POST'])
@login_required
def index():
    global TRANSAKSI, INVENTORY
    hpp, harga_jual = hitung_hpp_dan_harga()
    kapasitas_harian = hitung_kapasitas_harian()
    
    if request.method == 'POST':
        q_coklat = int(request.form.get('qty_coklat', 0))
        q_stroberi = int(request.form.get('qty_stroberi', 0))
        q_tiramissu = int(request.form.get('qty_tiramissu', 0))
        q_keju = int(request.form.get('qty_keju', 0))
        q_oreo = int(request.form.get('qty_oreo', 0))
        total_jumlah = q_coklat + q_stroberi + q_tiramissu + q_keju + q_oreo
        
        if total_jumlah > 0:
            butuh_ubi = total_jumlah * 40
            butuh_kulit = total_jumlah * 5
            if INVENTORY['ubi_ungu_gr'] < butuh_ubi or INVENTORY['kulit_lumpia_lbr'] < butuh_kulit:
                flash(f"⚠️ STOK INVENTORY TIDAK CUKUP! Butuh {butuh_ubi}g Ubi & {butuh_kulit}lbr Kulit.", "error")
                return redirect(url_for('index'))
            
            INVENTORY['ubi_ungu_gr'] -= butuh_ubi
            INVENTORY['kulit_lumpia_lbr'] -= butuh_kulit
            modal = total_jumlah * hpp
            pendapatan = total_jumlah * harga_jual
            laba = pendapatan - modal
            
            jam_tambahan = hitung_durasi_produksi_jam(total_jumlah)
            waktu_mulai = datetime.now()
            
            if len(TRANSAKSI) > 0:
                for t in reversed(TRANSAKSI):
                    if t['status_produksi'] == 'Diproses':
                        last_target = datetime.strptime(t['target_waktu_raw'], '%Y-%m-%d %H:%M:%S')
                        if last_target > waktu_mulai:
                            waktu_mulai = last_target
                        break
            
            waktu_target = waktu_mulai + timedelta(hours=jam_tambahan)
            
            jam = int(jam_tambahan)
            menit = int((jam_tambahan - jam) * 60)
            if jam > 0: estimasi_teks = f"{jam} Jam {menit} Menit"
            else: estimasi_teks = f"{menit} Menit"
            
            TRANSAKSI.append({
                'id': f"TRX-{len(TRANSAKSI) + 1:03d}", 
                'tanggal': datetime.now().strftime('%Y-%m-%d %H:%M'),
                'jumlah': total_jumlah, 
                'rasa': {'Coklat': q_coklat, 'Stroberi': q_stroberi, 'Tiramissu': q_tiramissu, 'Keju': q_keju, 'Oreo': q_oreo},
                'modal': modal, 'pendapatan': pendapatan, 'laba': laba, 
                'estimasi_teks': estimasi_teks,
                'target_waktu_raw': waktu_target.strftime('%Y-%m-%d %H:%M:%S'),
                'target_waktu_tampil': waktu_target.strftime('%H:%M WIB'),
                'status_produksi': 'Diproses',
                'status_pembayaran': 'Belum Lunas'
            })
            flash(f"✅ Pesanan diproses! Target selesai dalam {estimasi_teks}.", "success")
        return redirect(url_for('index'))

    filter_waktu = request.args.get('filter_waktu', 'semua')
    filter_status = request.args.get('filter_status', 'semua')
    tanggal_mulai = request.args.get('start_date')
    tanggal_akhir = request.args.get('end_date')
    
    transaksi_terfilter = []
    sekarang = datetime.now()
    
    for t in TRANSAKSI:
        if filter_status == 'belum_lunas' and t['status_pembayaran'] == 'Lunas': continue
        if filter_status == 'lunas' and t['status_pembayaran'] == 'Belum Lunas': continue
        
        tgl_transaksi = datetime.strptime(t['tanggal'], '%Y-%m-%d %H:%M')
        if filter_waktu == 'hari_ini' and tgl_transaksi.date() == sekarang.date(): transaksi_terfilter.append(t)
        elif filter_waktu == '7_hari' and (sekarang - tgl_transaksi).days <= 7: transaksi_terfilter.append(t)
        elif filter_waktu == 'bulan_ini' and tgl_transaksi.month == sekarang.month and tgl_transaksi.year == sekarang.year: transaksi_terfilter.append(t)
        elif filter_waktu == 'kustom' and tanggal_mulai and tanggal_akhir:
            try:
                start = datetime.strptime(tanggal_mulai, '%Y-%m-%d').date()
                end = datetime.strptime(tanggal_akhir, '%Y-%m-%d').date()
                if start <= tgl_transaksi.date() <= end: transaksi_terfilter.append(t)
            except ValueError: transaksi_terfilter.append(t)
        elif filter_waktu == 'semua': transaksi_terfilter.append(t)

    total_order = sum(t['jumlah'] for t in transaksi_terfilter)
    total_pendapatan = sum(t['pendapatan'] for t in transaksi_terfilter)
    total_laba = sum(t['laba'] for t in transaksi_terfilter)
    rekap_rasa = {'Coklat': 0, 'Stroberi': 0, 'Tiramissu': 0, 'Keju': 0, 'Oreo': 0}
    for t in transaksi_terfilter:
        for rasa, qty in t['rasa'].items():
            if rasa in rekap_rasa: rekap_rasa[rasa] += qty
    best_seller = max(rekap_rasa, key=rekap_rasa.get) if total_order > 0 else "-"

    return render_template('index.html', config=APP_CONFIG, transaksi=transaksi_terfilter, inventory=INVENTORY,
                           hpp=hpp, harga_jual=harga_jual, kapasitas_harian=kapasitas_harian,
                           total_order=total_order, total_pendapatan=total_pendapatan, total_laba=total_laba, best_seller=best_seller,
                           filter_aktif=filter_waktu, filter_status=filter_status, start_date=tanggal_mulai, end_date=tanggal_akhir)

# --- REVISI LOGIKA AKURAT 3 SIKLUS STATUS (Diproses -> Selesai -> Sukses) ---
@app.route('/update_status/<trx_id>', methods=['POST'])
@login_required
def update_status(trx_id):
    jenis = request.form.get('jenis')
    for t in TRANSAKSI:
        if t['id'] == trx_id:
            if jenis == 'produksi':
                if t['status_produksi'] == 'Diproses':
                    t['status_produksi'] = 'Selesai'
                elif t['status_produksi'] == 'Selesai':
                    t['status_produksi'] = 'Sukses'
                else:
                    t['status_produksi'] = 'Diproses'
            elif jenis == 'pembayaran':
                t['status_pembayaran'] = 'Lunas' if t['status_pembayaran'] == 'Belum Lunas' else 'Belum Lunas'
            break
    return redirect(url_for('index'))

@app.route('/costing', methods=['GET', 'POST'])
@login_required
def costing():
    if request.method == 'POST':
        save_costing(request.form)
        flash("💾 Data Costing berhasil diperbarui!", "success")
        return redirect(url_for('costing'))
    return render_template('costing.html', data_costing=load_costing())

@app.route('/proses', methods=['GET', 'POST'])
@login_required
def proses():
    if request.method == 'POST':
        save_proses(request.form)
        flash("💾 Data Lini Produksi berhasil diperbarui!", "success")
        return redirect(url_for('proses'))
    return render_template('proses.html', data_proses=load_proses())

@app.route('/settings', methods=['GET', 'POST'])
@login_required
def settings():
    if request.method == 'POST':
        action = request.form.get('action_type')
        if action == 'update_inventory':
            INVENTORY['ubi_ungu_gr'] += int(request.form.get('add_ubi', 0))
            INVENTORY['kulit_lumpia_lbr'] += int(request.form.get('add_kulit', 0))
            flash("📦 Stok Bahan Baku berhasil ditambah!", "success")
        elif action == 'update_config':
            APP_CONFIG['mesin_kukus'] = max(1, int(request.form.get('mesin_kukus', APP_CONFIG['mesin_kukus'])))
            APP_CONFIG['mesin_goreng'] = max(1, int(request.form.get('mesin_goreng', APP_CONFIG['mesin_goreng'])))
            APP_CONFIG['margin_target'] = float(request.form.get('margin_target', APP_CONFIG['margin_target']))
            APP_CONFIG['jam_kerja_hari'] = float(request.form.get('jam_kerja_hari', APP_CONFIG['jam_kerja_hari']))
            APP_CONFIG['overhead_tetap'] = int(request.form.get('overhead_tetap', APP_CONFIG['overhead_tetap']))
            APP_CONFIG['pembulatan_kelipatan'] = int(request.form.get('pembulatan_kelipatan', APP_CONFIG['pembulatan_kelipatan']))
            APP_CONFIG['user_role'] = request.form.get('user_role', APP_CONFIG['user_role'])
            flash("⚙️ Pengaturan & Konfigurasi Berhasil Disimpan!", "success")
        return redirect(url_for('settings'))
    return render_template('settings.html', config=APP_CONFIG, inventory=INVENTORY)

@app.route('/export')
@login_required
def export_csv():
    output = io.StringIO()
    writer = csv.writer(output, delimiter=';')
    writer.writerow(['ID Transaksi', 'Tanggal', 'Volume', 'Coklat', 'Stroberi', 'Tiramissu', 'Keju', 'Oreo', 'Modal Pokok', 'Pendapatan', 'Laba Bersih', 'SLA Target', 'Status Produksi', 'Status Pembayaran'])
    for t in TRANSAKSI: 
        writer.writerow([
            t['id'], t['tanggal'], t['jumlah'], 
            t.get('rasa', {}).get('Coklat', 0), t.get('rasa', {}).get('Stroberi', 0),
            t.get('rasa', {}).get('Tiramissu', 0), t.get('rasa', {}).get('Keju', 0), t.get('rasa', {}).get('Oreo', 0),
            int(t['modal']), int(t['pendapatan']), int(t['laba']), 
            t['target_waktu_tampil'], t['status_produksi'], t['status_pembayaran']
        ])
    excel_friendly_data = "\ufeff" + output.getvalue()
    response = Response(excel_friendly_data, mimetype="text/csv")
    response.headers["Content-Disposition"] = "attachment; filename=Laporan_Umbirolls.csv"
    return response

@app.route('/reset', methods=['POST'])
@login_required
def reset_data():
    global TRANSAKSI
    TRANSAKSI.clear()
    flash("🗑️ Semua data riwayat log berhasil dibersihkan!", "success")
    return redirect(url_for('index'))

@app.template_filter('rupiah')
def rupiah_filter(value):
    return f"Rp {value:,.0f}".replace(",", ".")

if __name__ == '__main__':
    app.run(debug=True)