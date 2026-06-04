# Dasbor Surat LPM

Dokumentasi diperbarui: 2026-06-04.

Dasbor Surat LPM adalah web app berbasis **Google Apps Script** untuk administrasi surat, dokumen mutu, sertifikat, peminjaman, dan kontak LPM Universitas Katolik Parahyangan. Aplikasi menggunakan **Google Spreadsheet** sebagai basis data, **Google Drive** sebagai penyimpanan arsip, dan **Google Contacts / People API** untuk sinkronisasi kontak.

## Ringkasan

Fitur utama aplikasi:

- pengelolaan surat keluar dengan nomor otomatis dan arsip Drive
- pengelolaan surat masuk dengan nomor manual dari dokumen asli
- pengelolaan Dokumen Sistem Penjaminan Mutu
- reservasi nomor sertifikat secara batch
- pencatatan peminjaman buku atau barang beserta reminder email
- manajemen kontak dan sinkronisasi ke Google Contacts
- audit log, notifikasi draft, trigger otomatis, dan utilitas pemeliharaan data

## Modul Aplikasi

### Surat Keluar

Modul ini mengelola surat keluar LPM.

- nomor surat final dibuat otomatis dari tanggal, sequence global, dan kategori Internal/Eksternal
- jenis naskah diambil dari konfigurasi `Config_NoSurat`
- penerima dan tembusan bisa dipilih dari daftar atau ditulis manual
- status data mendukung `DRAFT`, `SUBMITTED`, `RESERVED`, `CANCELLED`, dan `ARCHIVED`
- file arsip dapat diunggah ke Google Drive
- draft aktif muncul di notifikasi dan dapat dikirimkan reminder email
- setiap perubahan penting dicatat ke `Audit_Log`

### Surat Masuk

Modul ini mengelola metadata surat masuk.

- nomor surat diisi manual sesuai nomor pada dokumen asli
- tanggal terima dan tanggal surat disimpan terpisah
- pengirim, penerima, dan tembusan mendukung pilihan konfigurasi atau input manual
- arsip file tersimpan di Google Drive
- catatan atau disposisi tetap dapat diperbarui bersama arsip
- data dapat dicari, diedit, dan dihapus berbasis `Record ID`

### Dokumen Sistem Penjaminan Mutu

Modul ini mengelola dokumen non-surat seperti kebijakan, standar, manual, POB, instruksi kerja, formulir, laporan, dan pedoman.

- kode dokumen dibuat dari unit, jenis dokumen, kelompok kegiatan, tanggal berlaku, nomor dokumen, dan revisi
- nomor dokumen dihitung per kombinasi unit, jenis dokumen, dan kelompok kegiatan
- status dokumen mendukung `DRAFT`, `BERLAKU`, `REVISI`, dan `OBSOLETE`
- dokumen dapat ditandai obsolete tanpa menghapus riwayat
- arsip file tersimpan di folder Drive khusus

### Sertifikat

Modul sertifikat digunakan untuk reservasi nomor dalam jumlah banyak.

- sequence mengikuti nomor global surat keluar pada tahun berjalan
- kategori dapat diatur sebagai Internal atau Eksternal
- satu batch dapat mencadangkan banyak nomor sekaligus
- data sertifikat disimpan sebagai `RESERVED`
- modul ini tidak mewajibkan upload arsip per nomor

### Peminjaman

Modul peminjaman mencatat peminjaman buku atau barang.

- data yang disimpan mencakup email, nama peminjam, unit, nama buku/barang, tanggal pinjam, dan tanggal kembali
- riwayat dapat dicari dari antarmuka web
- sistem menghitung ringkasan total peminjaman dan reminder terkirim
- fungsi reminder email mengirim pengingat untuk peminjaman yang melewati tanggal kembali
- tersedia fungsi diagnosis trigger dan status reminder

### Manajemen Kontak

Modul kontak mengelola data pada sheet `Config_Contacts` dan sinkronisasinya ke Google Contacts.

- tambah, ubah, dan hapus kontak dari antarmuka web
- filter kontak berdasarkan status sync
- sinkronisasi perubahan spreadsheet ke Google Contacts
- import kontak dari Google Contacts ke spreadsheet
- deduplikasi kontak berdasarkan data yang tersedia
- pengelolaan grup kontak dan anggota grup
- trigger harian dapat dipasang untuk sync otomatis pukul 02:00

## Struktur File

- [Code.gs](Code.gs) - backend utama: routing web app, surat keluar, surat masuk, dokumen SPM, sertifikat, audit log, dashboard, konfigurasi, upload arsip, dan utilitas
- [Code-Peminjaman.gs](Code-Peminjaman.gs) - reminder, diagnosis, dan helper email peminjaman
- [Code-Contacts.gs](Code-Contacts.gs) - backend manajemen kontak, Google Contacts, People API, grup kontak, import, sync, deduplikasi, dan trigger sync harian
- [index.html](index.html) - antarmuka web responsif dengan modul utama, form, tabel, preview, toast, loading overlay, dan notifikasi draft
- [appsscript.json](appsscript.json) - manifest Apps Script, runtime V8, scope OAuth, dan konfigurasi People API

## Sheet Utama

Aplikasi membuat atau menggunakan beberapa sheet berikut:

- `Surat_Keluar`
- `Surat_Masuk`
- `Dokumen Sistem Penjaminan Mutu`
- `Data Peminjam Buku atau Barang`
- `Config_Contacts`
- `Config_Signatories`
- `Config_Recipients`
- `Config_NoSurat`
- `Config_Status`
- `Config_DokumenSPM`
- `Audit_Log`

Beberapa sheet legacy masih ditangani oleh kode, misalnya `Uji_Coba` untuk surat keluar lama dan `Tata_Naskah_Dinas` untuk dokumen SPM lama.

## Konfigurasi Penting

Konfigurasi utama berada di [Code.gs](Code.gs):

- `SPREADSHEET_ID`
- `DRIVE_ROOT_FOLDER_ID`
- `ALLOWED_DOMAIN`
- nama sheet utama dan sheet konfigurasi
- batas ukuran upload arsip
- umur minimal draft untuk reminder

Saat memindahkan project ke spreadsheet atau folder Drive baru, perbarui nilai ID tersebut sebelum deploy.

## Izin dan Layanan Google

Manifest [appsscript.json](appsscript.json) menggunakan:

- Google Sheets
- Google Drive
- user email
- Apps Script UI
- ScriptApp trigger
- MailApp / pengiriman email
- Google Contacts
- Google Contacts readonly
- Advanced Service `People` versi `v1`

Untuk modul kontak, pastikan **People API** aktif di Apps Script Advanced Services dan Google Cloud project terkait.

## Akses Pengguna

Web app dikonfigurasi dengan:

- `executeAs`: `USER_DEPLOYING`
- `access`: `DOMAIN`
- domain yang diizinkan di kode: `unpar.ac.id`

Pengguna di luar domain tersebut akan ditolak oleh fungsi pemeriksaan akses.

## Setup Awal

1. Pastikan `SPREADSHEET_ID` dan `DRIVE_ROOT_FOLDER_ID` sudah mengarah ke aset Google yang benar.
2. Aktifkan Advanced Service **People API** jika modul kontak digunakan.
3. Jalankan fungsi `initSpreadsheet()` dari editor Apps Script untuk membuat header sheet dan data konfigurasi awal.
4. Deploy ulang web app setelah mengubah manifest, scope, atau layanan Google.
5. Jalankan atau pasang trigger yang dibutuhkan:
   - `installProjectTriggers()` untuk trigger proyek utama
   - `installContactsDailySyncTrigger()` untuk sync kontak harian
   - trigger reminder peminjaman sesuai kebutuhan operasional

## Upload Arsip

Upload arsip mendukung beberapa file dengan batas:

- maksimal 4 file per unggahan
- ukuran tiap file 50 KB sampai 25 MB
- total unggahan maksimal 40 MB
- jenis file yang umum dipakai: PDF, DOCX, XLSX, JPG, PNG, GIF

File lama yang diganti akan dipindahkan ke trash jika penggantian arsip berhasil.

## Antarmuka

Frontend berada di [index.html](index.html) dan memuat:

- sidebar modul utama
- tab per modul
- form responsif desktop dan mobile
- preview otomatis untuk nomor dan penamaan file
- tabel riwayat dengan pencarian dan filter
- row cards untuk tabel pada layar kecil
- loading overlay dengan progress
- toast notification
- notifikasi draft surat
- ikon Lucide, Tailwind CDN, dan font Google Fonts

## Operasional Harian

Alur penggunaan umum:

1. Buka web app.
2. Pilih modul dari sidebar.
3. Isi form atau cari data pada riwayat.
4. Simpan sebagai draft, submit, cadangkan nomor, atau update data sesuai modul.
5. Unggah arsip bila diperlukan.
6. Pantau notifikasi, audit log, dan trigger otomatis.

## Pemeliharaan

Fungsi utilitas yang tersedia di backend antara lain:

- memperbaiki kolom hasil generate
- membersihkan trigger lama
- menyinkronkan konfigurasi jenis surat
- membuka folder Drive arsip
- menampilkan atau menyembunyikan sheet konfigurasi
- menyegarkan ringkasan dashboard dan laporan bulanan
- mendiagnosis reminder peminjaman
- memasang atau menghapus trigger sync kontak

## Catatan Developer

- Project ini tidak menggunakan build step lokal; file Apps Script langsung berupa `.gs`, `.html`, dan manifest JSON.
- Jika menggunakan `clasp`, pastikan project sudah terhubung ke Apps Script yang benar sebelum `push`.
- Jangan mengubah struktur kolom sheet tanpa menyesuaikan konstanta dan fungsi normalisasi di backend.
- Setelah menambah OAuth scope atau Advanced Service, deploy ulang web app dan minta otorisasi ulang jika diperlukan.

