# Dasbor Surat LPM

Aplikasi ini adalah web app **Google Apps Script** untuk membantu administrasi surat dan dokumen LPM. Sistem terhubung ke **Google Spreadsheet** sebagai basis data dan **Google Drive** sebagai penyimpanan arsip file.

## Ringkasan Fungsi

Dasbor ini digunakan untuk:

- mengelola **surat keluar**
- mengelola **surat masuk**
- mengelola **dokumen Sistem Penjaminan Mutu / Tata Naskah Dinas**
- mengelola **sertifikat**
- mengelola **peminjaman buku atau barang**
- mencatat **audit log** aktivitas data
- mengirim **notifikasi dan pengingat otomatis** melalui email

## Fitur Utama

### 1. Surat Keluar

- input data surat keluar melalui form web
- pembuatan nomor surat otomatis berdasarkan tipe surat, tanggal, dan urutan
- dukungan berbagai tipe surat seperti peraturan, keputusan, surat undangan, surat tugas, berita acara, laporan, dan lain-lain
- unggah file arsip ke Google Drive
- edit dan hapus data berdasarkan Record ID
- pencatatan audit setiap perubahan data

### 2. Surat Masuk

- input data surat masuk melalui form
- penyimpanan metadata surat ke spreadsheet
- lampiran file arsip jika diperlukan
- daftar data surat masuk yang dapat ditampilkan kembali dari sistem
- fungsi hapus data surat masuk dengan kontrol Record ID

### 3. Dokumen Sistem Penjaminan Mutu

- pengelolaan dokumen SPM / TND
- penomoran dokumen otomatis
- dukungan jenis dokumen seperti Kebijakan, Standar, Manual, Prosedur Operasional Baku, Instruksi Kerja, Formulir, dan lainnya
- riwayat dokumen dapat diperbarui dan diberi status
- pengarsipan file ke folder Drive khusus

### 4. Sertifikat

- pengelolaan data sertifikat dalam bentuk batch
- penomoran dan penyimpanan data sertifikat terstruktur
- dukungan pengolahan beberapa data sekaligus

### 5. Peminjaman

- pencatatan data peminjaman buku atau barang
- pengelolaan tanggal pinjam dan tanggal kembali
- pengiriman email pengingat otomatis untuk data yang melewati jatuh tempo
- diagnosis status reminder untuk mengecek trigger dan data yang belum terkirim

### 6. Dashboard dan Utilitas

- tampilan antarmuka web modern dan responsif
- navigasi modul melalui sidebar
- ringkasan dashboard dan laporan bulanan
- tampilan audit log
- akses cepat ke folder arsip Drive
- pengaturan visibilitas sheet konfigurasi
- fungsi perbaikan dan sinkronisasi data konfigurasi
- pembersihan trigger otomatis dan perbaikan kolom hasil generate

## Struktur File

- [Code.gs](Code.gs) - logika utama aplikasi, pengelolaan surat, dokumen, sertifikat, audit, dashboard, dan utilitas
- [Code-Peminjaman.gs](Code-Peminjaman.gs) - logika khusus reminder dan diagnosis peminjaman
- [index.html](index.html) - antarmuka web aplikasi
- [appsscript.json](appsscript.json) - konfigurasi project dan scope izin

## Alur Kerja Singkat

1. Pengguna membuka web app.
2. Pengguna memilih modul dari sidebar.
3. Data diinput melalui form.
4. Sistem menyimpan data ke Google Spreadsheet.
5. File arsip disimpan ke Google Drive jika ada lampiran.
6. Sistem membuat nomor otomatis, audit log, dan notifikasi sesuai kebutuhan.

## Konfigurasi Penting

Beberapa pengaturan utama ada di [Code.gs](Code.gs), seperti:

- Spreadsheet ID
- Drive Root Folder ID
- domain yang diizinkan
- nama sheet data dan sheet konfigurasi

Pastikan nilai-nilai tersebut sesuai dengan spreadsheet dan folder Drive milik proyek.

## Hak Akses

Web app ini dikonfigurasi untuk akses domain, dan pada kode saat ini domain yang diizinkan adalah:

- `unpar.ac.id`

## Izin yang Dibutuhkan

Dari [appsscript.json](appsscript.json), project ini menggunakan izin untuk:

- Spreadsheet
- Drive
- identitas pengguna
- UI Apps Script
- trigger script
- pengiriman email

## Catatan Penggunaan

- Gunakan data yang valid agar penomoran dan validasi berjalan benar.
- Pastikan trigger otomatis sudah terpasang jika ingin reminder email berjalan.
- Jika struktur sheet berubah, jalankan utilitas perbaikan atau sinkronisasi yang tersedia di script.

## Ringkasan Singkat

Aplikasi ini adalah dasbor administrasi LPM berbasis Google Apps Script yang memusatkan pengelolaan surat, dokumen mutu, sertifikat, dan peminjaman dalam satu sistem dengan nomor otomatis, arsip Drive, audit log, dan notifikasi email.