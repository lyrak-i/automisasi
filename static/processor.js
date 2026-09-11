/**
 * processor.js - 100% Native Client-Side Document Automation Engine
 * Berjalan langsung di peramban web modern tanpa server backend dan tanpa Pyodide/WebAssembly Python.
 * Pustaka yang digunakan:
 * - Mozilla PDF.js (Ekstraksi teks PDF)
 * - SheetJS / XLSX (Parsing file Excel .xlsx)
 * - JSZip (Ekstraksi dan pembentukan arsip .docx)
 * - Mozilla Nunjucks (Engine Jinja2 native JavaScript untuk template Word)
 */

// Konfigurasi Worker Mozilla PDF.js
if (typeof window !== 'undefined' && window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

/**
 * 1. Konversi Angka ke Kata dalam Bahasa Indonesia (Terbilang)
 * Digunakan untuk tanggal dan tahun dokumen resmi (100% deterministik).
 */
function terbilang(n) {
    n = Math.floor(Math.abs(Number(n)));
    if (isNaN(n) || n === 0) return 'Nol';
    const satuan = ['', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh', 'Sebelas'];
    if (n < 12) return satuan[n];
    if (n < 20) return terbilang(n - 10) + ' Belas';
    if (n < 100) return terbilang(Math.floor(n / 10)) + ' Puluh' + (n % 10 !== 0 ? ' ' + terbilang(n % 10) : '');
    if (n < 200) return 'Seratus' + (n % 100 !== 0 ? ' ' + terbilang(n % 100) : '');
    if (n < 1000) return terbilang(Math.floor(n / 100)) + ' Ratus' + (n % 100 !== 0 ? ' ' + terbilang(n % 100) : '');
    if (n < 2000) return 'Seribu' + (n % 1000 !== 0 ? ' ' + terbilang(n % 1000) : '');
    if (n < 1000000) return terbilang(Math.floor(n / 1000)) + ' Ribu' + (n % 1000 !== 0 ? ' ' + terbilang(n % 1000) : '');
    if (n < 1000000000) return terbilang(Math.floor(n / 1000000)) + ' Juta' + (n % 1000000 !== 0 ? ' ' + terbilang(n % 1000000) : '');
    if (n < 1000000000000) return terbilang(Math.floor(n / 1000000000)) + ' Miliar' + (n % 1000000000 !== 0 ? ' ' + terbilang(n % 1000000000) : '');
    return terbilang(Math.floor(n / 1000000000000)) + ' Triliun' + (n % 1000000000000 !== 0 ? ' ' + terbilang(n % 1000000000000) : '');
}

/**
 * 2. Format Tanggal Formal Bahasa Indonesia
 * Contoh: "hari ini Senin tanggal Dua Puluh Tiga bulan Februari tahun Dua Ribu Dua Puluh Enam (23-02-2026)"
 */
const NAMA_HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at", "Sabtu", "Minggu"];
const NAMA_BULAN = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
];

function formatTanggalIndo(tglMentah) {
    if (!tglMentah) return '';
    const str = String(tglMentah).trim();
    let day = null, month = null, year = null;

    // Coba format DD-MM-YYYY atau DD/MM/YYYY
    let m = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (m) {
        day = parseInt(m[1], 10);
        month = parseInt(m[2], 10);
        year = parseInt(m[3], 10);
    } else {
        // Coba format YYYY-MM-DD atau YYYY/MM/DD
        m = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
        if (m) {
            year = parseInt(m[1], 10);
            month = parseInt(m[2], 10);
            day = parseInt(m[3], 10);
        }
    }

    if (!day || !month || !year || month < 1 || month > 12 || day < 1 || day > 31) {
        return str;
    }

    try {
        const dateObj = new Date(year, month - 1, day);
        // Map getDay(): 0(Minggu)->6, 1(Senin)->0, dst.
        const dayIdx = (dateObj.getDay() + 6) % 7;
        const hari = NAMA_HARI[dayIdx];
        const bulan = NAMA_BULAN[month - 1];
        const tanggalHuruf = terbilang(day);
        const tahunHuruf = terbilang(year);
        const tanggalAngka = String(day).padStart(2, '0') + '-' + String(month).padStart(2, '0') + '-' + year;

        return `hari ini ${hari} tanggal ${tanggalHuruf} bulan ${bulan} tahun ${tahunHuruf} (${tanggalAngka})`;
    } catch (e) {
        return str;
    }
}

/**
 * 3. Normalisasi Teks
 */
function bersihkanTeks(teks) {
    if (teks === null || teks === undefined) return '';
    return String(teks).trim().toLowerCase().replace(/\s+/g, ' ');
}

/**
 * 4. Format Rupiah
 */
function formatRupiah(angka) {
    if (angka === null || angka === undefined || angka === '' || (typeof angka === 'number' && isNaN(angka))) {
        return "Rp. 0";
    }
    const cleanStr = String(angka).replace(/[^0-9.-]+/g, '');
    const num = parseFloat(cleanStr);
    if (isNaN(num)) return String(angka);
    const intPart = Math.floor(Math.abs(num));
    const formatted = intPart.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return (num < 0 ? "-Rp. " : "Rp. ") + formatted;
}

/**
 * 5. Ekstraksi Teks dan Metadata Seluruh Berkas PDF Menggunakan Mozilla PDF.js
 */
async function extractAllPdfs(pdfFiles, onProgress) {
    const results = [];
    const RE_ALASAN = /Alasan Kebutuhan Pengadaan Barang Dan Jasa\s*(?:\|\s*)?:\s*(.*)/i;
    const RE_TANGGAL = /Tanggal Kebutuhan\s*(?:\|\s*)?:\s*([\d-]+)/i;
    const RE_SPESIFIKASI = /technical specification\s*[:\-]?\s*([\s\S]*?)(?=\n\s*\d+\.|$)/i;
    const RE_RUANG_LINGKUP = /scope of work\s*[:\-]?\s*([\s\S]*?)(?=\n\s*\d+\.|$)/i;
    const RE_RL_CLEAN = /ruang\s+lingkup\s+pekerjaan\s+adalah\s+sebagai\s+berikut\s*[:;]?/i;

    for (let i = 0; i < pdfFiles.length; i++) {
        const file = pdfFiles[i];
        if (onProgress) {
            onProgress(`Mengekstrak PDF (${i + 1}/${pdfFiles.length}): ${file.name}...`);
        }
        const arrayBuffer = await file.arrayBuffer();
        const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) });
        const pdfDoc = await loadingTask.promise;
        const pagesText = [];

        for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
            const page = await pdfDoc.getPage(pageNum);
            const textContent = await page.getTextContent();
            
            let pageStr = '';
            let lastY = null;
            for (const item of textContent.items) {
                if (!item || !item.str) continue;
                const currentY = item.transform ? item.transform[5] : null;
                if (lastY !== null && currentY !== null && Math.abs(currentY - lastY) > 4) {
                    pageStr += '\n';
                } else if (pageStr.length > 0 && !pageStr.endsWith('\n') && !pageStr.endsWith(' ') && !item.str.startsWith(' ')) {
                    pageStr += ' ';
                }
                pageStr += item.str;
                if (item.hasEOL) {
                    pageStr += '\n';
                }
                lastY = currentY;
            }
            pagesText.push(pageStr);
        }

        const teksPdfFull = pagesText.join('\n');
        const teksHal1Dan2 = pagesText.slice(0, 2).join('\n');

        const cariAlasan = RE_ALASAN.exec(teksPdfFull);
        const alasanKunci = cariAlasan ? bersihkanTeks(cariAlasan[1]) : `tidak_ada_alasan_${file.name}`;

        const cariTanggal = RE_TANGGAL.exec(teksPdfFull);
        const tanggalPdf = cariTanggal ? formatTanggalIndo(cariTanggal[1].trim()) : "(Tanggal tidak ditemukan)";

        const cariSpesifikasi = RE_SPESIFIKASI.exec(teksHal1Dan2);
        const spesifikasiPdf = cariSpesifikasi ? cariSpesifikasi[1].trim() : "(Technical Specification tidak ditemukan)";

        const cariRuangLingkup = RE_RUANG_LINGKUP.exec(teksHal1Dan2);
        let ruangLingkupPdf = "(Scope of Work tidak ditemukan)";
        if (cariRuangLingkup) {
            const rlMentah = cariRuangLingkup[1].trim();
            const rlBersih = rlMentah.replace(RE_RL_CLEAN, '').trim();
            ruangLingkupPdf = rlBersih ? rlBersih : "(Isi Scope of Work kosong setelah dibersihkan)";
        }

        results.push({
            nama_file: file.name,
            teks_full_bersih: bersihkanTeks(teksPdfFull),
            alasan: alasanKunci,
            tanggal: tanggalPdf,
            spesifikasi: spesifikasiPdf,
            ruang_lingkup: ruangLingkupPdf
        });
    }

    return results;
}

/**
 * 6. Parsing dan Pemfilteran Data Excel Menggunakan SheetJS
 */
async function parseExcelData(excelFile, formValues) {
    const arrayBuffer = await excelFile.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        throw new Error("File Excel kosong atau tidak memiliki lembar kerja (worksheet).");
    }

    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (!rawRows || rawRows.length < 2) {
        throw new Error("File Excel tidak memiliki baris data yang cukup.");
    }

    // Tentukan baris header: prioritaskan baris index 1 (standar pandas header=1)
    let headerRowIdx = 1;
    let headers = rawRows[1] ? rawRows[1].map(h => String(h || '').trim()) : [];

    // Fallback ke baris 0 jika kolom filter tidak ditemukan di baris 1
    if (!headers.includes(formValues.col_kode) && rawRows[0]) {
        const headers0 = rawRows[0].map(h => String(h || '').trim());
        if (headers0.includes(formValues.col_kode)) {
            headerRowIdx = 0;
            headers = headers0;
        }
    }

    // Validasi kelengkapan kolom yang dibutuhkan
    const kolomDibutuhkan = [
        formValues.col_kode,
        formValues.col_no_frbmr,
        formValues.col_deskripsi_frbmr,
        formValues.col_desk_pekerjaan,
        formValues.col_qtty,
        formValues.col_uom,
        formValues.col_sl_nonsl,
        formValues.col_harga
    ];

    const kolomHilang = kolomDibutuhkan.filter(col => !headers.includes(col));
    if (kolomHilang.length > 0) {
        throw new Error(`Kolom hilang di file Excel: ${kolomHilang.join(', ')}. Periksa kembali file Excel Anda.`);
    }

    // Konversi baris data ke bentuk array of objects
    const records = [];
    for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!row || row.length === 0) continue;
        const obj = {};
        for (let c = 0; c < headers.length; c++) {
            const h = headers[c];
            if (h) {
                obj[h] = row[c] !== undefined ? row[c] : '';
            }
        }
        records.push(obj);
    }

    // Filter berdasarkan col_kode == val_kode
    const colKode = formValues.col_kode;
    const targetVal = String(formValues.val_kode).trim().toLowerCase();

    const dfFiltered = records.filter(row => {
        const val = String(row[colKode] || '').trim().toLowerCase();
        return val === targetVal;
    });

    if (dfFiltered.length === 0) {
        throw new Error(`Tidak ada data di Excel yang kolom "${colKode}" bernilai "${formValues.val_kode}".`);
    }

    return dfFiltered;
}

/**
 * 7. Pencocokan Data Excel dengan Berkas PDF
 */
function matchAndBuildBAs(dfFiltered, bukuCatatanPdf, formValues) {
    const colNoFrbmr = formValues.col_no_frbmr;
    const colDeskripsiFrbmr = formValues.col_deskripsi_frbmr;
    const colDeskPekerjaan = formValues.col_desk_pekerjaan;
    const colQtty = formValues.col_qtty;
    const colUom = formValues.col_uom;
    const colSlNonsl = formValues.col_sl_nonsl;
    const colHarga = formValues.col_harga;

    // Grouping berdasarkan No. FRBMR
    const groups = new Map();
    for (const row of dfFiltered) {
        const noFrbmr = row[colNoFrbmr] !== undefined ? String(row[colNoFrbmr]).trim() : '';
        if (!groups.has(noFrbmr)) {
            groups.set(noFrbmr, []);
        }
        groups.get(noFrbmr).push(row);
    }

    const daftarBa = [];

    for (const [noFrbmr, groupRows] of groups.entries()) {
        const firstRow = groupRows[0];
        const deskripsiFrbmr = String(firstRow[colDeskripsiFrbmr] || '');
        const deskripsiPekerjaan = String(firstRow[colDeskPekerjaan] || '');

        const kunciNoFrbmr = bersihkanTeks(noFrbmr);
        const kunciExcelFrbmr = bersihkanTeks(deskripsiFrbmr);
        const kunciExcelPekerjaan = bersihkanTeks(deskripsiPekerjaan);

        let tanggalCocok = '';
        let spesifikasiCocok = '';
        let ruangLingkupCocok = '';
        let pdfDitemukan = false;

        for (const dataPdf of bukuCatatanPdf) {
            const teksPdf = dataPdf.teks_full_bersih;
            const alasanPdf = dataPdf.alasan;
            let matchFound = false;

            // Prioritas 1: No. FRBMR terdapat di dalam teks PDF
            if (kunciNoFrbmr && kunciNoFrbmr !== 'nan' && kunciNoFrbmr !== 'null' && teksPdf.includes(kunciNoFrbmr)) {
                matchFound = true;
            }
            // Prioritas 2: Deskripsi FRBMR terdapat di teks PDF
            else if (kunciExcelFrbmr && kunciExcelFrbmr.length > 5 && teksPdf.includes(kunciExcelFrbmr)) {
                matchFound = true;
            }
            // Prioritas 3: Deskripsi Pekerjaan terdapat di teks PDF
            else if (kunciExcelPekerjaan && kunciExcelPekerjaan.length > 5 && teksPdf.includes(kunciExcelPekerjaan)) {
                matchFound = true;
            }
            // Prioritas 4: Fallback overlap teks alasan kebutuhan
            else if (alasanPdf && !alasanPdf.startsWith('tidak_ada_alasan')) {
                if (kunciExcelPekerjaan.includes(alasanPdf) || alasanPdf.includes(kunciExcelPekerjaan) ||
                    kunciExcelFrbmr.includes(alasanPdf) || alasanPdf.includes(kunciExcelFrbmr)) {
                    matchFound = true;
                }
            }

            if (matchFound) {
                tanggalCocok = dataPdf.tanggal;
                spesifikasiCocok = dataPdf.spesifikasi;
                ruangLingkupCocok = dataPdf.ruang_lingkup;
                pdfDitemukan = true;
                break;
            }
        }

        // Lewati jika tidak ada PDF yang cocok
        if (!pdfDitemukan) {
            continue;
        }

        const tabelPekerjaan = groupRows.map(row => ({
            deskripsi_pekerjaan: row[colDeskPekerjaan] !== undefined ? String(row[colDeskPekerjaan]) : '',
            qtty: row[colQtty] !== undefined ? row[colQtty] : '',
            uom: row[colUom] !== undefined ? String(row[colUom]) : '',
            harga: formatRupiah(row[colHarga]),
            sl_nonsl: row[colSlNonsl] !== undefined ? String(row[colSlNonsl]) : ''
        }));

        daftarBa.push({
            deskripsi_frbmr: deskripsiFrbmr,
            tabel_pekerjaan: tabelPekerjaan,
            tanggal: tanggalCocok,
            spesifikasi: spesifikasiCocok,
            ruang_lingkup: ruangLingkupCocok
        });
    }

    if (daftarBa.length === 0) {
        throw new Error("Proses dihentikan: Tidak ada satupun No. FRBMR di Excel yang cocok dengan teks di dalam file PDF yang Anda unggah.");
    }

    return daftarBa;
}

/**
 * 8. Patching XML Word Document untuk Sintaks Nunjucks / Jinja2
 */
function patchDocxXml(srcXml) {
    // 1. Bersihkan pemisahan kurung kurawal oleh tag XML
    srcXml = srcXml.replace(/(?<={)(<[^>]*>)+(?=[{%#])/g, '');
    srcXml = srcXml.replace(/(?<=[%#}])(<[^>]*>)+(?=})/g, '');

    // 2. Hapus tag pemisah w:t di dalam {% ... %}, {# ... #}, {{ ... }}
    srcXml = srcXml.replace(/({%[\s\S]*?%}|{#[\s\S]*?#}|{{[\s\S]*?}})/g, function (match) {
        return match.replace(/<\/w:t>[\s\S]*?(<w:t>|<w:t\b[^>]*>)/g, '');
    });

    // 3. Pastikan xml:space="preserve"
    srcXml = srcXml.replace(/<w:t>((?:(?!<w:t>).)*)({{.*?}}|{%.*?%})/g, '<w:t xml:space="preserve">$1$2');

    // 4. Ubah statement baris/paragraf {%p ... %}, {%tr ... %} menjadi statement template Nunjucks
    for (const y of ['tr', 'tc', 'p', 'r']) {
        const pat = new RegExp(`<w:${y}[ >](?:(?!<w:${y}[ >]).)*?({%|{{)${y} ([^}%]*(?:%}|}})).*?<\\/w:${y}>`, 'g');
        srcXml = srcXml.replace(pat, '$1 $2');
    }

    return srcXml;
}

/**
 * XML Sanitizer untuk memastikan data pengguna aman dan tidak merusak sintaks XML Word
 */
function escapeXml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function escapeDataForXml(val) {
    if (typeof val === 'string') {
        return escapeXml(val);
    }
    if (Array.isArray(val)) {
        return val.map(escapeDataForXml);
    }
    if (val && typeof val === 'object') {
        const res = {};
        for (const [k, v] of Object.entries(val)) {
            res[k] = escapeDataForXml(v);
        }
        return res;
    }
    return val;
}

/**
 * 9. Render Template Word (.docx) Secara Native Menggunakan JSZip & Nunjucks
 */
async function renderWordDocument(wordTemplateFile, daftarBa) {
    const arrayBuffer = await wordTemplateFile.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    const docXmlEntry = zip.file("word/document.xml");
    if (!docXmlEntry) {
        throw new Error("File Word tidak valid: file 'word/document.xml' tidak ditemukan.");
    }

    const rawDocXml = await docXmlEntry.async("text");
    const patchedXml = patchDocxXml(rawDocXml);

    // Konfigurasi Nunjucks dengan autoescape: false (karena template berupa raw XML yang sudah disanitasi)
    const nunjucksEnv = new nunjucks.Environment(null, { autoescape: false });
    const safeData = escapeDataForXml({ daftar_berita_acara: daftarBa });

    let renderedXml;
    try {
        renderedXml = nunjucksEnv.renderString(patchedXml, safeData);
    } catch (err) {
        console.error("Nunjucks rendering error:", err);
        throw new Error("Gagal merender template Word: " + err.message);
    }

    // Perbaiki penomoran docPr ID agar tetap unik
    let docPrId = 0;
    renderedXml = renderedXml.replace(/(<wp:docPr\b[^>]*\bid=")\d+(")/g, function (match, prefix, suffix) {
        docPrId++;
        return `${prefix}${docPrId}${suffix}`;
    });

    // Simpan kembali ke arsip docx
    zip.file("word/document.xml", renderedXml);

    // Hasilkan Uint8Array berkas Word (.docx)
    const outputBytes = await zip.generateAsync({
        type: "uint8array",
        compression: "DEFLATE",
        compressionOptions: { level: 6 }
    });

    return outputBytes;
}

/**
 * 10. Fungsi Eksekusi Utama (100% Client-Side Native Web Automation)
 */
async function runClientSideProcess(excelFile, wordTemplateFile, pdfFiles, formValues, onProgress) {
    // 1. Ekstrak data PDF
    if (onProgress) onProgress(`Mengekstrak teks dari ${pdfFiles.length} file PDF pelengkap...`);
    const bukuCatatanPdf = await extractAllPdfs(pdfFiles, onProgress);

    // 2. Baca dan filter Excel
    if (onProgress) onProgress("Membaca dan memfilter data Excel...");
    const dfFiltered = await parseExcelData(excelFile, formValues);

    // 3. Cocokkan Excel dengan PDF
    if (onProgress) onProgress("Mencocokkan baris Excel dengan dokumen PDF...");
    const daftarBa = matchAndBuildBAs(dfFiltered, bukuCatatanPdf, formValues);

    // 4. Render dokumen Word
    if (onProgress) onProgress("Menyusun Berita Acara ke dalam template Word (.docx)...");
    const outputBytes = await renderWordDocument(wordTemplateFile, daftarBa);

    if (onProgress) onProgress("Dokumen selesai dibuat!");
    return outputBytes;
}

// Ekspor ke window global object
if (typeof window !== 'undefined') {
    window.DocAutomation = {
        terbilang,
        formatTanggalIndo,
        bersihkanTeks,
        formatRupiah,
        extractAllPdfs,
        parseExcelData,
        matchAndBuildBAs,
        patchDocxXml,
        renderWordDocument,
        runClientSideProcess
    };
}
