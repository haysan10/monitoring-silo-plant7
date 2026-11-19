# Monitoring Silo Clinker · PLANT 78

Aplikasi **Progressive Web App (PWA)** mobile-first industrial grade untuk monitoring stok clinker di Silo 7.1, 8.1, dan 8.2 (**PLANT 78**).

![PWA Ready](https://img.shields.io/badge/PWA-Ready-orange?style=flat-square)
![GSAP Powered](https://img.shields.io/badge/Animation-GSAP%203-green?style=flat-square)
![Security](https://img.shields.io/badge/Access-NIK%20%26%20PIN%20Admin-red?style=flat-square)
![Offline First](https://img.shields.io/badge/Offline-ServiceWorker-blue?style=flat-square)
![Database](https://img.shields.io/badge/Database-Google%20Sheets-success?style=flat-square)

---

## 🔐 Manajemen Hak Akses & Keamanan
1. **Operator Lapangan**:
   - **Wajib NIK 7 Digit**: Setiap pengisian form wajib mencantumkan 7 angka NIK pekerja untuk tracking siapa yang mengisi.
   - **Jendela Edit 8 Jam**: Operator hanya dapat mengedit isian form selama durasi 8 jam shift berjalan:
     - **Shift A**: `07:00 – 15:00`
     - **Shift B**: `15:01 – 23:00`
     - **Shift C**: `23:01 – 07:00`
   - **Proteksi Penguncian Shift (> 8 Jam)**: Jika shift telah berakhir, data otomatis terkunci (*read-only*) untuk mencegah manipulasi.
   - **Fitur Lapor Salah Input**: Jika terdapat kesalahan ukur pada shift yang telah terkunci, operator dapat menekan tombol **"📢 Lapor Salah Input ke Admin"** untuk mengirim pesan revisi otomatis via WhatsApp atau salin ke clipboard.
   - **Akses Terbatas**: Operator umum hanya dapat melihat **Dasbor & Tren Stok**, serta mengisi form shift berjalan. Riwayat dan Kalibrasi dikunci.

2. **Administrator (PIN Default: `100700`)**:
   - **Buka Riwayat**: Melihat seluruh riwayat shift masa lalu lengkap dengan kolom **NIK Operator**.
   - **Edit & Hapus Bebas**: Admin berwenang merevisi atau menghapus catatan kapan pun tanpa batasan waktu 8 jam.
   - **Ekspor & Impor**: Ekspor CSV (38 kolom) dan impor Excel.
   - **Kalibrasi Fleksibel**:
     - *Kalibrasi Global*: Berlaku untuk semua tanggal secara default.
     - *Kalibrasi Khusus*: Menetapkan bobot dan faktor spesifik hanya pada tanggal & shift tertentu.
   - **Ubah PIN**: PIN admin dapat diganti sewaktu-waktu melalui menu Admin.

---

## ✨ Fitur Utama Lainnya
- **Tampilan Default Light Mode**: Tampilan default terang dan bersih sesuai standar lapangan, dengan opsi beralih ke Dark Mode via tombol `🌓`.
- **Cinematic Industrial Splash Screen (GSAP 3)**: Animasi vektor 3 Silo clinker, balok cairan clinker molten, dan laser sensor piezo-ultrasonik.
- **Anti-Double Entry**: Sistem secara cerdas meng-update baris pada tanggal dan shift yang sama selama shift berjalan tanpa membuat baris duplikat.
- **Dual-Mode Connectivity**:
  - Live PWA di GitHub Pages: [https://haysan10.github.io/monitoring-silo-plant7/](https://haysan10.github.io/monitoring-silo-plant7/)
  - Google Apps Script Web App: [Buka Google Web App](https://script.google.com/macros/s/AKfycbxtrdHL6AJT0Np6dbtktSvVEOisgM3O2iills43YnbVj5oZy7nERun20mAk0jEqo_NcsA/exec)

---

## 📂 Struktur File
- `index.html` : Aplikasi PWA mandiri (HTML, CSS, JavaScript, GSAP Splash, database embedded 2.913 baris, sistem PIN & NIK).
- `Code.gs` : Script backend Google Apps Script (REST API `doGet`/`doPost`, verifikasi PIN Admin, dan pencatatan NIK ke sheet `Rekap Silo 2026`).
- `manifest.webmanifest` / `manifest.json` : Konfigurasi instalasi PWA.
- `sw.js` : Service Worker untuk caching aset dan offline capability.
- `icon-192.png`, `icon-512.png`, `favicon.svg` : Ikon industrial resmi Silo Plant 78.

---
© 2026 PLANT 78 · Unit Operasi & Kendali Mutu Silo Clinker.
