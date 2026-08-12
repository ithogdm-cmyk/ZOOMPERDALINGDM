# Panduan Pemasangan: Frontend Vercel + Backend API Google Apps Script

Sistem absensi webinar ini sekarang dipisahkan menjadi dua bagian:
1. **Backend API (Google Apps Script)**: Mengelola penulisan/pembacaan data di Google Sheets secara aman.
2. **Frontend Statis (Vercel)**: Situs web absensi & dashboard dengan performa tinggi yang di-host di Vercel secara gratis.

---

## Langkah 1: Persiapan Google Sheets
1. Buka Google Sheets Anda di link: `https://docs.google.com/spreadsheets/d/1fEzzdCOt4Gof-ZUmgW6Px_EjdJ1jkJKH0-8eUsEES9o/edit`
2. Pastikan Anda memiliki sheet bernama **`RAW`** yang bertindak sebagai database pendaftaran.
3. Di dalam sheet **`RAW`**, pastikan **baris pertama (Header)** minimal berisi kolom:
   - **`Email`**
   - **`Nama`**
4. **Tidak perlu membuat sheet Kehadiran secara manual**. Sistem akan otomatis membuat sheet baru bernama **`Kehadiran`** lengkap dengan header terformat rapi (`Email`, `Nama`, `Status Kehadiran`, `Waktu Absen`) saat ada peserta yang pertama kali melakukan absensi.

---

## Langkah 2: Deploy Backend API (Google Apps Script)
1. Di dalam Google Sheets Anda, pada menu atas, klik **Extensions (Ekstensi)** > **Apps Script**.
2. Hapus semua kode bawaan di berkas **`Code.gs`**.
3. Salin seluruh isi berkas dari [`Code.gs`](Code.gs) lokal Anda, lalu tempel (paste) ke editor.
4. Klik ikon **Save (Simpan / Ikon Disket)** di bagian atas editor.
5. Klik tombol **Deploy** di kanan atas > **New deployment (Penerapan baru)**.
6. Klik ikon gerigi (Select type), lalu pilih **Web app (Aplikasi web)**.
7. Isi konfigurasi sebagai berikut:
   - **Description**: `API Absensi Webinar Perdalin GDM 2026`
   - **Execute as (Jalankan sebagai)**: Pilih **`Me (email-anda@gmail.com)`**.
   - **Who has access (Siapa yang memiliki akses)**: Pilih **`Anyone (Siapa saja)`**.
8. Klik tombol **Deploy**.
9. **Persetujuan Akses (Authorization)**:
   - Google akan meminta Anda memberikan izin akses. Klik **Authorize Access**, pilih akun Google Anda, klik **Advanced (Lanjutan)**, klik **Go to Untitled project (unsafe)** atau nama proyek Anda, lalu klik **Allow (Izinkan)**.
10. Setelah selesai, Anda akan mendapatkan **Web App URL** (misalnya: `https://script.google.com/macros/s/AKfycb.../exec`). **Salin URL ini!**

---

## Langkah 3: Konfigurasi Frontend Lokal
1. Buka berkas konfigurasi lokal [`config.js`](config.js).
2. Tempelkan URL Web App yang telah Anda salin tadi ke bagian `API_URL`:
   ```javascript
   const WEBINAR_CONFIG = {
     API_URL: "https://script.google.com/macros/s/AKfycbxxxxxxxxx/exec", // Tempel URL Anda di sini
     WEBINAR_TITLE: "Webinar PERDALIN - GDM 2026",
     ZOOM_LINK: "https://zoom.us/j/your-webinar-id" // Sesuaikan link Zoom jika ada
   };
   ```
3. Simpan berkas `config.js` Anda.

---

## Langkah 4: Deploy Frontend ke Vercel

Ada dua cara mudah untuk mempublikasikan frontend Anda ke Vercel:

### Opsi A: Menggunakan Vercel CLI (Paling Cepat dari Terminal)
1. Buka terminal/command prompt pada komputer Anda.
2. Arahkan ke folder proyek absensi ini (`C:\Users\IT GLOBALDISPOMEDIKA\.gemini\antigravity\scratch\zoom-attendance-gas\`).
3. Install Vercel secara global (jika belum pernah):
   ```bash
   npm install -g vercel
   ```
4. Jalankan perintah deploy:
   ```bash
   vercel
   ```
5. Masuk (*login*) ke akun Vercel Anda jika diminta, dan ikuti instruksi pengaturan proyek (tekan enter/yes untuk konfigurasi bawaan).
6. Setelah proses deploy pertama selesai, jalankan perintah berikut untuk mengunggah ke produksi:
   ```bash
   vercel --prod
   ```
7. Anda akan mendapatkan URL Vercel produksi Anda (misalnya: `https://zoom-attendance-gas.vercel.app`).

### Opsi B: Menggunakan GitHub & Vercel Dashboard (Rekomendasi untuk Integrasi Kontinu)
1. Unggah (*push*) folder lokal `zoom-attendance-gas` ini ke repositori baru di GitHub/GitLab Anda.
2. Buka dashboard Vercel Anda di [vercel.com](https://vercel.com).
3. Klik **Add New** > **Project**.
4. Impor repositori GitHub absensi Anda.
5. Klik **Deploy**. Vercel akan membaca konfigurasi `vercel.json` secara otomatis dan membuat URL produksi untuk Anda.

---

## Langkah 5: Cara Mengakses Halaman di Vercel

Setelah di-deploy ke Vercel, URL Anda akan tampak rapi tanpa ekstensi `.html` berkat fitur `cleanUrls` di `vercel.json`:

1. **Halaman Absensi Peserta & OTS**:
   - Akses langsung URL utama Vercel Anda:
     `https://nama-proyek-anda.vercel.app/`
2. **Dashboard Monitor & Admin Panel**:
   - Cukup tambahkan `/dashboard` di akhir URL:
     `https://nama-proyek-anda.vercel.app/dashboard`
