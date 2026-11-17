# Monitoring Silo Clinker · PLANT 78

Aplikasi **Progressive Web App (PWA)** mobile-first industrial grade untuk monitoring stok clinker di Silo 7.1, 8.1, dan 8.2 (**PLANT 78**).

![PWA Ready](https://img.shields.io/badge/PWA-Ready-orange?style=flat-square)
![GSAP Powered](https://img.shields.io/badge/Animation-GSAP%203-green?style=flat-square)
![Offline First](https://img.shields.io/badge/Offline-ServiceWorker-blue?style=flat-square)
![Database](https://img.shields.io/badge/Database-Google%20Sheets-success?style=flat-square)

---

## ✨ Fitur Utama
1. **Cinematic Industrial Splash Screen (GSAP 3)**:
   - Animasi vektor 3 Silo clinker, balok pengisian clinker molten, dan laser sensor piezo-ultrasonik.
   - Tipografi kinetik Plant 78 dengan progress ticker diagnostik.
   - Dapat di-skip seketika dengan ketukan layar dan diulang via tombol `⚡`.
2. **Multi-Granularitas Monitoring**:
   - **⏱️ Setiap Shift**: Real-time shift terakhir (C, A, B), level meter, stok aktual (ton), dan selisih ton vs shift sebelumnya.
   - **📅 Rekap Harian**: Penutupan hari kalender dan selisih konsumsi/pengisian clinker 24 jam.
   - **📊 Rekap Mingguan**: Rolling aggregate 7 hari untuk evaluasi ketersediaan clinker jangka menengah.
3. **Database 1:1 Format Excel `Silo clinker 16C.xlsx`**:
   - Terhubung langsung ke Google Sheets `Rekap Silo 2026` (37 kolom).
   - Seluruh **2.913 catatan riwayat** (2024–2026) telah ditanam (embedded) di dalam aplikasi, sehingga Google Spreadsheet dapat dimulai dari kondisi bersih/kosong tanpa memindahkan ribuan baris masa lalu.
4. **PWA Standalone & Offline-First**:
   - Dapat diinstal ke homescreen Android / iOS / Desktop seperti aplikasi native.
   - Didukung Service Worker (`sw.js`) dan IndexedDB / localStorage untuk operasional tanpa jaringan internet di area silo pabrik.
5. **Konektivitas Dual-Mode**:
   - **Google Apps Script Web App**: [Buka Web App GAS](https://script.google.com/macros/s/AKfycbxtrdHL6AJT0Np6dbtktSvVEOisgM3O2iills43YnbVj5oZy7nERun20mAk0jEqo_NcsA/exec)
   - **GitHub Pages PWA**: Menggunakan REST API CORS ke deployment Google Apps Script di atas.

---

## 🚀 Cara Hosting PWA di GitHub Pages
1. Buat repository baru di GitHub (misalnya: `monitoring-silo-plant7`).
2. Push seluruh file di repositori ini:
   ```bash
   git init
   git add .
   git commit -m "feat: PWA Silo Clinker Plant 78 with GSAP Splash Screen"
   git branch -M main
   git remote add origin https://github.com/<username>/<repo-name>.git
   git push -u origin main
   ```
3. Buka **Settings** repository GitHub > menu **Pages** (di sidebar kiri).
4. Pada bagian **Build and deployment** > **Branch**:
   - Pilih `main` dan folder `/ (root)`.
   - Klik **Save**.
5. Dalam 1-2 menit, URL PWA Anda akan aktif di:
   `https://<username>.github.io/<repo-name>/`
6. Buka URL tersebut di browser smartphone operator (Chrome / Safari), lalu pilih **"Add to Home Screen"** / **"Install App"**.

---

## 📂 Struktur File
- `index.html` : Aplikasi tunggal PWA (HTML, CSS industrial, JavaScript, GSAP Splash, database embedded 2.913 baris).
- `manifest.webmanifest` / `manifest.json` : Konfigurasi instalasi PWA.
- `sw.js` : Service Worker untuk caching aset dan offline capability.
- `Code.gs` : Script backend Google Apps Script (penyedia API `doGet`/`doPost` dan pembentuk sheet `Rekap Silo 2026`).
- `icon-192.png`, `icon-512.png`, `favicon.svg` : Ikon industrial resmi Silo Plant 78.
- `Silo clinker 16C.xlsx` : Data acuan asli master pabrik.

---
© 2026 PLANT 78 · Unit Operasi & Kendali Mutu Silo Clinker Plant 78.
