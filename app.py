import io
import logging
import os
import re
import urllib.parse
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber
from docxtpl import DocxTemplate
from flask import Flask, render_template, request, send_file
from num2words import num2words
from werkzeug.utils import secure_filename

# --------------------------------------------------------------------------
# Konfigurasi Aplikasi (100% In-Memory, Tanpa Penyimpanan File di Server)
# --------------------------------------------------------------------------
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # Maksimal total upload 100 MB

logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')
logger = logging.getLogger(__name__)

# Ekstensi yang diizinkan per jenis file (whitelist keamanan)
ALLOWED_EXT = {
    'excel': {'.xlsx'},
    'word': {'.docx'},
    'pdf': {'.pdf'},
}

# --------------------------------------------------------------------------
# Konstanta level-modul
# --------------------------------------------------------------------------
NAMA_HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at", "Sabtu", "Minggu"]
NAMA_BULAN = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]
FORMAT_TANGGAL = ["%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y/%m/%d"]

# Regex dikompilasi sekali di level modul
RE_ALASAN = re.compile(
    r'Alasan Kebutuhan Pengadaan Barang Dan Jasa\s*(?:\|\s*)?:\s*(.*)', re.IGNORECASE
)
RE_TANGGAL = re.compile(r'Tanggal Kebutuhan\s*(?:\|\s*)?:\s*([\d-]+)', re.IGNORECASE)
RE_SPESIFIKASI = re.compile(
    r'technical specification\s*[:\-]?\s*(.*?)(?=\n\s*\d+\.|\Z)', re.IGNORECASE | re.DOTALL
)
RE_RUANG_LINGKUP = re.compile(
    r'scope of work\s*[:\-]?\s*(.*?)(?=\n\s*\d+\.|\Z)', re.IGNORECASE | re.DOTALL
)


# 1. Pengaturan Tanggal Indonesia
def format_tanggal_indo(tgl_mentah: str) -> str:
    dt = None
    for fmt in FORMAT_TANGGAL:
        try:
            dt = datetime.strptime(tgl_mentah, fmt)
            break
        except (ValueError, TypeError):
            continue

    if dt is None:
        return tgl_mentah

    try:
        hari = NAMA_HARI[dt.weekday()]
        bulan = NAMA_BULAN[dt.month - 1]
        tanggal_huruf = num2words(dt.day, lang='id').title()
        tahun_huruf = num2words(dt.year, lang='id').title()
        tanggal_angka = dt.strftime("%d-%m-%Y")
        return (
            f"hari ini {hari} tanggal {tanggal_huruf} bulan {bulan} "
            f"tahun {tahun_huruf} ({tanggal_angka})"
        )
    except Exception:
        logger.exception("Gagal memformat tanggal Indonesia untuk nilai: %s", tgl_mentah)
        return tgl_mentah


# 2. Bersihkan Teks
def bersihkan_teks(teks) -> str:
    """Normalisasi teks (hapus spasi berlebih, lowercase) untuk pencocokan kunci."""
    return re.sub(r'\s+', ' ', str(teks).strip().lower())


# 3. Format Rupiah
def format_rupiah(angka) -> str:
    if pd.isna(angka):
        return "Rp. 0"
    try:
        angka_int = int(float(angka))
        return f"Rp. {angka_int:,}".replace(',', '.')
    except (ValueError, TypeError):
        return str(angka)


# 4. Validasi Ekstensi File
def is_allowed(filename: str, kind: str) -> bool:
    if not filename:
        return False
    return Path(filename).suffix.lower() in ALLOWED_EXT[kind]


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process_files():
    excel_file = request.files.get('excel_file')
    word_file = request.files.get('word_template')
    pdf_files = [f for f in request.files.getlist('pdf_files') if f.filename]

    if not excel_file or not is_allowed(excel_file.filename, 'excel'):
        return render_template('index.html', error="File Excel wajib diunggah (.xlsx)."), 400
    if not word_file or not is_allowed(word_file.filename, 'word'):
        return render_template('index.html', error="Template Word wajib diunggah (.docx)."), 400
    if not pdf_files or not all(is_allowed(f.filename, 'pdf') for f in pdf_files):
        return render_template('index.html', error="Semua file pelengkap harus berformat .pdf."), 400

    kolom_form = {
        key: request.form.get(key, '').strip()
        for key in (
            'col_kode', 'val_kode', 'col_no_frbmr', 'col_deskripsi_frbmr',
            'col_desk_pekerjaan', 'col_qtty', 'col_uom', 'col_sl_nonsl', 'col_harga',
        )
    }
    if not all(kolom_form.values()):
        return render_template('index.html', error="Semua kolom pengaturan filter wajib diisi."), 400

    try:
        # Baca seluruh file ke RAM (In-Memory BytesIO) - TIDAK DISIMPAN KE DISK SERVER
        excel_io = io.BytesIO(excel_file.read())
        word_io = io.BytesIO(word_file.read())

        # ==========================================
        # FASE 1: BACA SEMUA PDF SECARA IN-MEMORY
        # ==========================================
        buku_catatan_pdf = []

        for pdf_file in pdf_files:
            pdf_io = io.BytesIO(pdf_file.read())

            try:
                with pdfplumber.open(pdf_io) as pdf:
                    teks_per_halaman = [halaman.extract_text() or '' for halaman in pdf.pages]
            except Exception:
                logger.exception("Gagal membaca file PDF: %s", pdf_file.filename)
                return render_template(
                    'index.html', 
                    error=f"File '{pdf_file.filename}' rusak atau bukan format PDF yang valid."
                ), 400

            teks_pdf_full = "\n".join(teks_per_halaman)
            teks_hal_1_dan_2 = "\n".join(teks_per_halaman[:2])

            cari_alasan = RE_ALASAN.search(teks_pdf_full)
            alasan_kunci = (
                bersihkan_teks(cari_alasan.group(1))
                if cari_alasan else f"tidak_ada_alasan_{pdf_file.filename}"
            )

            cari_tanggal = RE_TANGGAL.search(teks_pdf_full)
            tanggal_pdf = (
                format_tanggal_indo(cari_tanggal.group(1).strip())
                if cari_tanggal else "(Tanggal tidak ditemukan)"
            )

            cari_spesifikasi = RE_SPESIFIKASI.search(teks_hal_1_dan_2)
            spesifikasi_pdf = (
                cari_spesifikasi.group(1).strip()
                if cari_spesifikasi else "(Technical Specification tidak ditemukan)"
            )

            # Pembersihan Scope of Work
            cari_ruang_lingkup = RE_RUANG_LINGKUP.search(teks_hal_1_dan_2)
            if cari_ruang_lingkup:
                rl_mentah = cari_ruang_lingkup.group(1).strip()
                rl_bersih = re.sub(
                    r'(?i)ruang\s+lingkup\s+pekerjaan\s+adalah\s+sebagai\s+berikut\s*[:;]?', 
                    '', 
                    rl_mentah
                ).strip()
                ruang_lingkup_pdf = rl_bersih if rl_bersih else "(Isi Scope of Work kosong setelah dibersihkan)"
            else:
                ruang_lingkup_pdf = "(Scope of Work tidak ditemukan)"

            buku_catatan_pdf.append({
                'teks_full_bersih': bersihkan_teks(teks_pdf_full),
                'alasan': alasan_kunci,
                'tanggal': tanggal_pdf,
                'spesifikasi': spesifikasi_pdf,
                'ruang_lingkup': ruang_lingkup_pdf,
                'nama_file': pdf_file.filename,
            })

        # ==========================================
        # FASE 2: BACA EXCEL SECARA IN-MEMORY
        # ==========================================
        df = pd.read_excel(excel_io, header=1)
        df.columns = df.columns.str.strip()

        kolom_dibutuhkan = [
            kolom_form['col_kode'], kolom_form['col_no_frbmr'], kolom_form['col_deskripsi_frbmr'],
            kolom_form['col_desk_pekerjaan'], kolom_form['col_qtty'], kolom_form['col_uom'],
            kolom_form['col_sl_nonsl'], kolom_form['col_harga'],
        ]
        kolom_hilang = set(kolom_dibutuhkan) - set(df.columns)
        if kolom_hilang:
            return render_template('index.html', error=f"Kolom hilang: {', '.join(sorted(kolom_hilang))}."), 400

        col_kode = kolom_form['col_kode']
        df[col_kode] = df[col_kode].astype(str).str.strip()
        df_filtered = df[df[col_kode].str.lower() == kolom_form['val_kode'].lower()]

        if df_filtered.empty:
            return render_template('index.html', error="Tidak ada data Excel yang cocok dengan filter."), 400

        # ==========================================
        # FASE 3: COCOKKAN EXCEL DENGAN PDF
        # ==========================================
        col_no_frbmr = kolom_form['col_no_frbmr']
        col_deskripsi_frbmr = kolom_form['col_deskripsi_frbmr']
        col_desk_pekerjaan = kolom_form['col_desk_pekerjaan']
        col_qtty = kolom_form['col_qtty']
        col_uom = kolom_form['col_uom']
        col_sl_nonsl = kolom_form['col_sl_nonsl']
        col_harga = kolom_form['col_harga']

        daftar_ba = []
        for no_frbmr, group in df_filtered.groupby(col_no_frbmr):
            deskripsi_frbmr = str(group[col_deskripsi_frbmr].iloc[0])
            deskripsi_pekerjaan = str(group[col_desk_pekerjaan].iloc[0])

            # Ambil kunci dari Excel dan bersihkan
            kunci_no_frbmr = bersihkan_teks(str(no_frbmr))
            kunci_excel_frbmr = bersihkan_teks(deskripsi_frbmr)
            kunci_excel_pekerjaan = bersihkan_teks(deskripsi_pekerjaan)

            tanggal_cocok = ""
            spesifikasi_cocok = ""
            ruang_lingkup_cocok = ""
            pdf_ditemukan = False 

            for data_pdf in buku_catatan_pdf:
                teks_pdf = data_pdf['teks_full_bersih']
                alasan_pdf = data_pdf['alasan']
                match_found = False

                # 1. Prioritas Utama: Cari No. FRBMR
                if kunci_no_frbmr and kunci_no_frbmr != 'nan' and kunci_no_frbmr in teks_pdf:
                    match_found = True
                
                # 2. Prioritas Kedua: Cari Deskripsi FRBMR
                elif kunci_excel_frbmr and len(kunci_excel_frbmr) > 5 and kunci_excel_frbmr in teks_pdf:
                    match_found = True
                
                # 3. Prioritas Ketiga: Cari Deskripsi Pekerjaan
                elif kunci_excel_pekerjaan and len(kunci_excel_pekerjaan) > 5 and kunci_excel_pekerjaan in teks_pdf:
                    match_found = True
                
                # 4. Fallback Terakhir: Cek overlap alasan
                elif alasan_pdf and not alasan_pdf.startswith('tidak_ada_alasan'):
                    if (kunci_excel_pekerjaan in alasan_pdf or alasan_pdf in kunci_excel_pekerjaan or
                        kunci_excel_frbmr in alasan_pdf or alasan_pdf in kunci_excel_frbmr):
                        match_found = True

                if match_found:
                    tanggal_cocok = data_pdf['tanggal']
                    spesifikasi_cocok = data_pdf['spesifikasi']
                    ruang_lingkup_cocok = data_pdf['ruang_lingkup']
                    pdf_ditemukan = True
                    break

            # Lewati jika tidak ada PDF yang cocok
            if not pdf_ditemukan:
                continue

            tabel_pekerjaan = [
                {
                    'deskripsi_pekerjaan': row[col_desk_pekerjaan],
                    'qtty': row[col_qtty],
                    'uom': row[col_uom],
                    'harga': format_rupiah(row[col_harga]),
                    'sl_nonsl': row[col_sl_nonsl],
                }
                for row in group.to_dict('records')
            ]

            daftar_ba.append({
                'deskripsi_frbmr': deskripsi_frbmr,
                'tabel_pekerjaan': tabel_pekerjaan,
                'tanggal': tanggal_cocok,
                'spesifikasi': spesifikasi_cocok,
                'ruang_lingkup': ruang_lingkup_cocok,
            })

        if not daftar_ba:
            return render_template('index.html', error="Proses dihentikan: Tidak ada satupun No. FRBMR di Excel yang cocok dengan teks di dalam file PDF yang Anda unggah."), 400

        # ==========================================
        # FASE 4: RENDER KE TEMPLATE WORD SECARA IN-MEMORY
        # ==========================================
        doc = DocxTemplate(word_io)
        doc.render({'daftar_berita_acara': daftar_ba})

        output_stream = io.BytesIO()
        doc.save(output_stream)
        output_stream.seek(0)

        nama_kode_aman = secure_filename(kolom_form['val_kode']) or 'output'
        download_name = f"Berita_Acara_Otomatis_{nama_kode_aman}.docx"

        response = send_file(
            output_stream,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=download_name,
        )

        # Set cookie di browser pengguna untuk mengingat histori sesi tanpa menyimpan file di server
        waktu_sekarang = datetime.now().strftime("%d/%m/%Y, %H:%M WIB")
        cookie_options = {
            'max_age': 30 * 86400,  # Simpan selama 30 hari
            'samesite': 'Lax',
            'path': '/',
        }
        response.set_cookie('ba_last_excel', urllib.parse.quote(excel_file.filename), **cookie_options)
        response.set_cookie('ba_last_word', urllib.parse.quote(word_file.filename), **cookie_options)
        response.set_cookie('ba_last_pdf_count', str(len(pdf_files)), **cookie_options)
        response.set_cookie('ba_last_val_kode', urllib.parse.quote(kolom_form['val_kode']), **cookie_options)
        response.set_cookie('ba_last_time', urllib.parse.quote(waktu_sekarang), **cookie_options)

        return response

    except KeyError as e:
        logger.exception("Kolom tidak ditemukan saat memproses data")
        return render_template('index.html', error=f"Kolom tidak ditemukan: {e}"), 400
    except Exception:
        logger.exception("Kesalahan tak terduga saat memproses data")
        pesan = "Terjadi kesalahan saat memproses data. Silakan periksa kembali file Anda."
        return render_template('index.html', error=pesan), 500


if __name__ == '__main__':
    debug_mode = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=debug_mode, host='0.0.0.0', port=port)