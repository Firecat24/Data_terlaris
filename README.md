# Dynamic Multi-Year Data Dashboard (Google Apps Script)

Proyek ini merupakan solusi automasi dashboard yang memungkinkan pengguna untuk melakukan filtrasi data lintas tahun secara dinamis. Sistem ini dirancang untuk menangani kumpulan data besar (500+ entri) dengan performa optimal melalui teknik *dependent dropdown* dan pemrosesan data asinkron.

## 🚀 Fitur Utama

* **Multi-Level Dependent Dropdown**: Sistem dropdown berjenjang di mana pilihan "Nama Produk" secara otomatis menyesuaikan dengan "Bulan" yang dipilih, dan pilihan "Satuan" menyesuaikan dengan "Produk" yang spesifik.
* **High-Volume Data Handling**: Menggunakan metode `requireValueInRange` dengan *Hidden Helper Sheet* untuk mengatasi batasan karakter pada dropdown standar Google Sheets, memungkinkan pencarian hingga ribuan item secara lancar.
* **Automatic Data Extraction**: Secara otomatis menarik dan menjumlahkan data kuantitas (QTY) dari berbagai tabel tahunan (misal: 2023, 2024, 2025) berdasarkan kriteria yang dipilih.
* **Smart Cache & UI Cleaning**: Mekanisme pembersihan otomatis pada sel terkait jika ada perubahan pada level filtrasi yang lebih tinggi, mencegah terjadinya ketidaksinkronan data (data mismatch).
* **Event-Driven Automation**: Memanfaatkan trigger `onEdit` untuk memberikan pengalaman pengguna yang interaktif dan responsif tanpa perlu menjalankan skrip secara manual.

## 🛠️ Teknologi yang Digunakan

* **Runtime**: Google Apps Script (V8 Engine)
* **Spreadsheet Service**: Manipulasi DOM spreadsheet tingkat lanjut.
* **Data Validation API**: Untuk membangun kontrol input yang ketat dan *user-friendly*.
* **Logic Optimization**: Implementasi *Set object* untuk pemrosesan array unik yang cepat dan efisien.

## 📈 Arsitektur Logika

1. **Selection Phase**: Pengguna memilih bulan pada dashboard.
2. **Indexing Phase**: Skrip memindai kolom referensi di sheet bulanan yang relevan, mengumpulkan semua entitas unik, dan memindahkannya ke sheet helper tersembunyi.
3. **Validation Injection**: Skrip menyuntikkan aturan validasi data ke sel tujuan berdasarkan indeks yang baru dibuat.
4. **Calculation Phase**: Setelah semua parameter terpenuhi, skrip melakukan kalkulasi lintas kolom untuk menyajikan data QTY secara komparatif antar tahun.

---

> [!IMPORTANT]
> **Disclaimer:** > Kode yang ditampilkan dalam repositori ini merupakan **sampel (logic demo)** untuk keperluan portofolio. Implementasi ini mencerminkan arsitektur sistem manajemen data yang lebih luas dan bertujuan untuk mendemonstrasikan kemampuan dalam pengembangan logika automasi serta manajemen database *flat-file* pada Google Workspace.

---
*Developed by **Muhammad Farhan Putra Pratama, S.H.*** *Sebuah solusi efisiensi administrasi kesehatan berbasis teknologi otomasi.*

<img width="1235" height="890" alt="image" src="https://github.com/user-attachments/assets/a84019fc-df2e-4850-827e-49c2b0934906" />
