# Automasi Berita Acara (100% Client-Side Web)

Aplikasi web modern untuk mengotomasi pembuatan dokumen Berita Acara Kesepakatan dari data Excel (`.xlsx`) dan berkas PDF pendukung ke dalam template Word (`.docx`).

Aplikasi ini berjalan **100% di browser pengguna (Client-Side)** menggunakan JavaScript native tanpa server backend dan tanpa runtime Python/Pyodide.

---

## 🚀 Fitur Unggulan

- **100% Native Web Browser**: Menggunakan pustaka client-side terpercaya (SheetJS, JSZip, Mozilla Nunjucks, Mozilla PDF.js).
- **Kinerja Instan**: Dokumen diproses dalam hitungan milidetik langsung di memori komputer Anda.
- **Privasi & Keamanan**: Seluruh dokumen hanya diproses di RAM peramban Anda dan tidak pernah dikirim ke server/internet manapun.
- **Drag and Drop**: Memudahkan pemilihan berkas Excel, template Word, dan banyak berkas PDF sekaligus.
- **Dukungan Format Resmi Indonesia**: Tanggal formal Indonesia terbilang otomatis dan format mata uang Rupiah.
- **Desain Modern & Responsif**: Tampilan minimalis elegan dengan dukungan otomatis Mode Gelap (*Dark Mode*).
- **Siap Hosting Gratis**: Dapat langsung di-host di GitHub Pages, Netlify, Vercel, atau dibuka langsung secara lokal.

---

## 📁 Struktur Berkas

```text
Automasi/
├── index.html                      # Halaman website utama
├── .nojekyll                       # Memastikan aset JavaScript dimuat di GitHub Pages
├── .gitignore                      # Mengabaikan file sistem lokal
├── README.md                       # Dokumentasi aplikasi
├── static/
│   ├── processor.js                # Engine otomasi dokumen 100% native JavaScript
│   └── style.css                   # Desain modern, responsif & dark mode
└── contoh/
    ├── template_berita_acara.docx  # Berkas template Word contoh
    └── contoh_excel.xlsx           # Berkas data Excel contoh
```

---

## 🌐 Cara Penggunaan Lokal

Anda dapat membuka file `index.html` langsung di peramban web modern (Google Chrome, Microsoft Edge, Mozilla Firefox, Safari) atau menjalankan server lokal sederhana:

```powershell
python -m http.server 8000
```
Lalu buka alamat: `http://localhost:8000` di peramban Anda.
