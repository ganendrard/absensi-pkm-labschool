# 🏫 Web App Absensi Mahasiswa PKM — SMP Labschool Ciracas

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white)

Aplikasi berbasis web (*Web Application*) berbasis serverless gratis yang dibangun menggunakan **Google Apps Script** untuk mengelola absensi harian dan rekapitulasi kehadiran mahasiswa Praktik Keterampilan Mengajar (PKM) di SMP Labschool Ciracas.

Sistem ini memfasilitasi mahasiswa PKM dari berbagai Program Studi:
* 💻 **Pendidikan Teknik Informatika dan Komputer (PTIK)**
* 🎨 **Pendidikan Seni Rupa (SR)**
* 💬 **Bimbingan dan Konseling (BK)**

---

## ✨ Fitur-Fitur Utama

### 📱 1. Portal Absensi Mahasiswa (`Index.html`)
* **Validasi Swafoto (Camera Capture):** Pengambilan foto kehadiran *real-time* menggunakan kamera HP/Laptop (mencegah unggahan dari galeri foto).
* **Fitur Anti-Spoofing Timestamp:** Validasi jam otomatis dari perangkat untuk mencegah manipulasi waktu pengambilan foto.
* **Geofencing / Radius GPS:** Penguncian titik lokasi berbasis koordinat GPS sekolah (dapat diatur batas jarak maksimumnya).
* **Pencarian Nama & Autocomplete:** Modal pencarian nama interaktif yang dikelompokkan secara otomatis berdasarkan Program Studi (PTIK, SR, BK).
* **Indikator Kehadiran Mingguan:** Panel ringkasan status kehadiran Senin–Jumat *real-time* langsung di layar HP mahasiswa.

### 📊 2. Dashboard Rekap Tata Usaha (`Rekap.html`)
* **Gerbang Keamanan PIN:** Akses khusus staf/TU menggunakan verifikasi PIN berbasis token sesi (`CacheService`).
* **Sistem Periode Fleksibel:** Filter otomatis berbasis periode bulanan (tanggal 21 hingga 20 bulan berikutnya) atau rentang kustom.
* **Kalkulasi Otomatis & Persentase:** Menghitung otomatis status *Hadir*, *Lupa Absen*, *Izin*, *Sakit*, *Dinas Luar*, *Cuti*, dan *Tanpa Keterangan (Alpa)* dilengkapi *progress bar* tingkat kehadiran.
* **Kelola Keterangan / Izin:** Modal interaktif untuk mengubah status alpa menjadi Izin/Sakit/Dinas dengan lampiran catatan.

### ⚡ 3. Backend & Manajemen Data (`Code.gs`)
* **Auto-Setup Spreadsheet & Drive:** Fungsi `setupTemplate()` untuk membuat struktur sheet, pemformatan tabel, dan folder Google Drive secara otomatis.
* **Penyimpanan Foto Google Drive:** Foto swafoto absensi diunggah dan disimpan langsung ke folder Google Drive terkait.
* **Anti-Concurrency Lock:** Menggunakan `LockService` untuk menangani lonjakan absen simultan tanpa menyebabkan bentrok data (*data race*).

---

## 🛠️ Tech Stack & Arsitektur

* **Language:** JavaScript (ES6+), HTML5, CSS3
* **Backend Platform:** Google Apps Script (GAS)
* **Database / Storage:** Google Sheets API & Google Drive API
* **Security & Session:** GAS `CacheService` & `LockService`
* **Web APIs Used:** HTML5 Geolocation API, HTML5 Canvas API (Image Compression), MediaDevices / File API

---

## 🚀 Panduan Instalasi & Deployment

Jika Anda ingin menjalankan atau mereplikasi proyek ini di akun Google Anda sendiri, ikuti langkah-langkah berikut:

### Step 1: Persiapan Google Sheets & Editor Apps Script
1. Buat **Google Sheet Baru** yang masih kosong.
2. Buka menu **Extensions (Ekstensi)** > **Apps Script**.
3. Di dalam editor Apps Script:
   * Ubah file `Code.gs` bawaan dan tempelkan kode dari file [`Code.gs`](./Code.gs).
   * Buat file HTML baru (beri nama `Index`) dan tempelkan kode dari [`Index.html`](./Index.html).
   * Buat file HTML baru (beri nama `Rekap`) dan tempelkan kode dari [`Rekap.html`](./Rekap.html).

### Step 2: Konfigurasi Parameter (`Code.gs`)
Buka file `Code.gs` dan sesuaikan variabel konfigurasi di bagian atas file:
```javascript
const PIN_REKAP        = '123456';               // PIN akses dashboard Rekap TU
const LATITUDE_SEKOLAH  = -6.315921;             // Koordinat Latitude Sekolah
```

### Step 3: Jalankan Inisialisasi Otomatis
1. Pada editor Apps Script, pilih fungsi setupTemplate dari daftar fungsi.
2. Klik tombol Run.
3. Berikan izin otorisasi (Grant Permissions) saat jendela pop-up Google meminta persetujuan.
4. Fungsi ini akan otomatis membuat tab Absensi, tab Keterangan, serta folder baru di Google Drive untuk menyimpan foto.

### Step 4: Deploy Web App
Klik tombol Deploy di pojok kanan atas > New deployment.
1. Pilih jenis deployment: Web app.
2. Atur konfigurasi berikut:
   - Execute as: Me (Akun Google Anda)
   - Who has access: Anyone (Siapa saja)
3. Klik Deploy dan salin Web App URL yang dihasilkan.

### Panduan Penggunaan URL
- Halaman Absensi Mahasiswa: https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec
- Halaman Dashboard Rekap TU: https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?page=rekap

Lisensi
Proyek ini dikembangkan untuk kebutuhan pendidikan dan kolaborasi akademis. Bebas digunakan, dipelajari, atau dikembangkan kembali untuk kebutuhan instansi/sekolah lain.
