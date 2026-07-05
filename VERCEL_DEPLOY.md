# Panduan Publikasi ke Vercel (PWA Asisten Koper Elhakim)

Dokumen ini berisi petunjuk lengkap untuk menerbitkan aplikasi **Sistem Manajemen Koper Jemaah Elhakim** ke **Vercel** dengan konfigurasi full-stack (Express API + React Frontend) dan Google Sheets Direct Connection.

---

## 🚀 Langkah Cepat Deploy ke Vercel

### 1. Persiapan Kode di GitHub
1. Hubungkan repositori Anda ke **GitHub**, **GitLab**, atau **Bitbucket**.
2. Push seluruh kode proyek ini ke repositori Anda.

### 2. Impor Proyek di Dashboard Vercel
1. Masuk ke [Vercel Dashboard](https://vercel.com).
2. Klik tombol **"Add New"** -> **"Project"**.
3. Pilih repositori proyek ini dari daftar akun GitHub Anda.

### 3. Konfigurasi Pengaturan Build (Build Settings)
Vercel akan otomatis mendeteksi proyek ini sebagai aplikasi **Vite**. Pastikan pengaturannya sesuai dengan berikut:
* **Framework Preset**: `Vite` (atau `Other` jika tidak terdeteksi)
* **Build Command**: `npm run build`
* **Output Directory**: `dist`
* **Install Command**: `npm install`

### 4. Konfigurasi Environment Variables (PENTING!)
Sebelum mengklik tombol **Deploy**, tambahkan variabel lingkungan (Environment Variables) berikut di bagian **Environment Variables** di Vercel:

| Nama Variabel | Deskripsi | Contoh Nilai |
| :--- | :--- | :--- |
| `GEMINI_API_KEY` | Kunci API Gemini untuk analisis stiker koper | `AIzaSy...` (Bisa Anda dapatkan di Google AI Studio) |
| `NODE_ENV` | Mode server (Wajib diset ke production) | `production` |

> **Catatan Firebase:** Seluruh konfigurasi Firebase Authentication Anda telah tersimpan dengan aman di dalam file `firebase-applet-config.json` di dalam proyek, sehingga Anda tidak perlu mendefinisikan variabel Firebase secara manual lagi!

---

## 🛠️ Cara Kerja Arsitektur Full-Stack di Vercel

Aplikasi ini menggunakan konfigurasi serverless modern di Vercel:
1. **Frontend (React + Vite + Tailwind)**: Semua file statis dikompilasi oleh Vite ke folder `dist` dan dilayani secara instan lewat CDN Vercel berkecepatan tinggi yang mendukung PWA secara penuh.
2. **Backend (Express API)**: Semua permintaan dengan pola `/api/*` secara otomatis diarahkan ke serverless function di `/api/index.ts` (dikonfigurasi melalui file `vercel.json`).

---

## 📲 Pengaturan Progressive Web App (PWA) di Vercel

Setelah Anda berhasil mendeploy ke Vercel, aplikasi Anda akan otomatis dilayani melalui protokol **HTTPS** (Wajib untuk PWA).
* Di Google Chrome, Anda akan melihat tombol **"Instal Aplikasi"** / ikon komputer di bilah alamat URL untuk memasang aplikasi secara instan.
* Di HP, Anda dapat membukanya melalui Chrome atau Safari lalu memilih opsi **"Tambahkan ke Layar Utama"** untuk menikmati asisten lapangan native dengan ikon koper Elhakim!
