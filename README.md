

/// ▄▄▄█████▓ ██▓ ██▀███   ██▀███  ▓█████ ▒██   ██▒
/// ▓  ██▒ ▓▒▓██▒▓██ ▒ ██▒▓██ ▒ ██▒▓█   ▀ ▒▒ █ █ ▒░
/// ▒ ▓██░ ▒░▒██▒▓██ ░▄█ ▒▓██ ░▄█ ▒▒███   ░░  █   ░
/// ░ ▓██▓ ░ ░██░▒██▀▀█▄  ▒██▀▀█▄  ▒▓█  ▄  ░ █ █ ▒ 
///   ▒██▒ ░ ░██░░██▓ ▒██▒░██▓ ▒██▒░▒████▒▒██▒ ▒██▒
///   ▒ ░░   ░▓  ░ ▒▓ ░▒▓░░ ▒▓ ░▒▓░░░ ▒░ ░▒▒ ░ ░▓ ░
///     ░     ▒ ░  ░▒ ░ ▒░  ░▒ ░ ▒░ ░ ░  ░░░   ░▒ ░
///   ░       ▒ ░  ░░   ░   ░░   ░    ░    ░    ░  
///           ░     ░        ░        ░  ░ ░    ░  



Aplikasi desktop dengan GUI yang memudahkan penulisan skripsi di Microsoft Word dengan template dan fitur otomatis.

![Status](https://img.shields.io/badge/status-stable-green)
![Platform](https://img.shields.io/badge/platform-windows-blue)
![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)

# Skripsweet Shortcut 📚✨

Aplikasi desktop dengan GUI yang memudahkan penulisan skripsi di Microsoft Word dengan template dan fitur otomatis.

## 📥 Cara Install dan Penggunaan

1. Download dan extract semua file dalam folder ini
2. Double click pada file `Skripsweet Shortcut.exe`
3. Pastikan Microsoft Word sudah terbuka
4. Tunggu sampai status "Terhubung" muncul di aplikasi
5. Aplikasi siap digunakan!

## 🌟 Fitur

- Pembuatan otomatis untuk:
  - Daftar Isi
  - Daftar Tabel
  - Daftar Gambar
  - BAB I-V dengan template standar
    note: buat Daftar Isi setelah BAB I-V selesai, baru generate Daftar Daftar karena automatis membaca BAB dan Subbabnya.

- 📐 **Page Setup Otomatis**
  - Format A4
  - Margin atas & kiri: 3.5 cm
  - Margin bawah & kanan: 2.5 cm

- 🎨 **Format Standar Skripsi**
  - Font: Times New Roman
  - Ukuran: 12pt
  - Spasi: 1.5
  - Perataan teks yang sesuai standar
  - Penomoran otomatis

- ✨ **Fitur Tambahan**
  - Interface yang user-friendly
  - Status koneksi Word yang real-time
  - Pembuatan subbab otomatis
  - Template placeholder untuk konten

## 💻 Persyaratan Sistem

- Windows 8/10/11
- Microsoft Word 2010 atau lebih baru
- [Microsoft Visual C++ Redistributable](https://learn.microsoft.com/en-us/cpp/windows/latest-supported-vc-redist?view=msvc-170) (biasanya sudah terinstall di Windows)

## 📥 Cara Install

### Metode 1: Langsung Pakai (Recommended)
1. Download file exe dari [Releases](https://github.com/yourusername/skripsweet-shortcut/releases)
2. Extract file zip yang didownload
3. Double click pada file `Skripsweet Shortcut.exe`
4. Aplikasi siap digunakan!

### Metode 2: Dari Source Code
1. Install Python 3.7 atau lebih baru
2. Clone repository ini:
   ```bash
   git clone https://github.com/yourusername/skripsweet-shortcut.git
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Jalankan aplikasi:
   ```bash
   python skripsweet_shortcut_gui.py
   ```

## 🚀 Cara Penggunaan

1. Buka Microsoft Word
2. Jalankan aplikasi Skripsweet Shortcut
3. Tunggu sampai status "Terhubung" muncul di aplikasi
4. Klik tombol sesuai fitur yang diinginkan:
   - "Atur Page Setup" untuk mengatur format halaman
   - "Buat DAFTAR ISI" untuk membuat daftar isi otomatis
   - "Buat DAFTAR TABEL" untuk membuat daftar tabel
   - "Buat DAFTAR GAMBAR" untuk membuat daftar gambar
   - "Buat BAB I - V" untuk membuat bab dengan template lengkap

## ⚠️ Troubleshooting

1. **Aplikasi tidak bisa terhubung ke Word**
   - Pastikan Microsoft Word sudah terinstall
   - Tutup dan buka ulang Word jika not responding
   - Restart aplikasi Skripsweet Shortcut

2. **Tombol tidak bereaksi**
   - Pastikan status "Terhubung" sudah muncul
   - Pastikan Word tidak sedang busy/not responding
   - Coba tutup dan buka ulang Word

3. **Format tidak sesuai**
   - Klik tombol "Atur Page Setup" untuk memperbaiki format
   - Pastikan tidak ada style kustom yang konflik
   - Gunakan fitur "Format Ulang" jika tersedia

## 🛠️ Teknologi yang Digunakan

- Python 3.7+
- Tkinter untuk GUI
- Win32com untuk integrasi dengan Microsoft Word
- Pillow untuk pemrosesan icon
- PyInstaller untuk pembuatan executable

## 📝 Lisensi

MIT License - lihat file [LICENSE](LICENSE) untuk detail lengkap.

## 🤝 Kontribusi

Kontribusi selalu diterima dengan senang hati! Beberapa cara untuk berkontribusi:

1. 🐛 Laporkan bug
2. 💡 Usulkan fitur baru
3. 📖 Perbaiki dokumentasi
4. 🔀 Submit pull request

## 📧 Kontak

Untuk pertanyaan dan saran, silakan:
- Buat issue di GitHub
- Email: [mohammedwinston@yahoo.com]

## ✨ Credits

Dibuat dengan ❤️ untuk memudahkan mahasiswa dalam menulis skripsi.

---
**Note**: Aplikasi ini dibuat untuk membantu format penulisan skripsi secara umum. Sesuaikan dengan pedoman penulisan dari institusi Anda.
