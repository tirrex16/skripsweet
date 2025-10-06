

/// ▄▄▄█████▓ ██▓ ██▀███   ██▀███  ▓█████ ▒██   ██▒
/// ▓  ██▒ ▓▒▓██▒▓██ ▒ ██▒▓██ ▒ ██▒▓█   ▀ ▒▒ █ █ ▒░
/// ▒ ▓██░ ▒░▒██▒▓██ ░▄█ ▒▓██ ░▄█ ▒▒███   ░░  █   ░      # Skripsweet Shortcut - Development Guide
/// ░ ▓██▓ ░ ░██░▒██▀▀█▄  ▒██▀▀█▄  ▒▓█  ▄  ░ █ █ ▒       
///   ▒██▒ ░ ░██░░██▓ ▒██▒░██▓ ▒██▒░▒████▒▒██▒ ▒██▒
///   ▒ ░░   ░▓  ░ ▒▓ ░▒▓░░ ▒▓ ░▒▓░░░ ▒░ ░▒▒ ░ ░▓ ░
///     ░     ▒ ░  ░▒ ░ ▒░  ░▒ ░ ▒░ ░ ░  ░░░   ░▒ ░
///   ░       ▒ ░  ░░   ░   ░░   ░    ░    ░    ░  
///           ░     ░        ░        ░  ░ ░    ░  
🛠️Dokumentasi pengembangan untuk Skripsweet Shortcut, aplikasi GUI yang membantu penulisan skripsi di Microsoft Word.

## 🏗️ Struktur Kode

Aplikasi ini terdiri dari dua kelas utama:

### 1. `WordHelper` Class
Menangani semua interaksi dengan Microsoft Word melalui COM automation.

#### Metode Utama:
- `get_word_app()`: Menghubungkan ke instance Word yang aktif atau membuat baru
- `ensure_doc()`: Memastikan dokumen aktif tersedia
- `apply_default_formatting()`: Mengatur format standar dokumen
- `create_toc()`: Membuat daftar isi
- `create_specific_bab()`: Membuat BAB dengan template
- `create_list_tables()`: Membuat daftar tabel
- `create_list_figures()`: Membuat daftar gambar
- `set_page_setup()`: Mengatur format halaman A4 dan margin

### 2. `Application` Class
Mengatur GUI dan event handling menggunakan Tkinter.

#### Komponen GUI:
- Status bar koneksi Word
- Tombol untuk fitur-fitur utama
- Indikator status aplikasi

## 🔧 Pengembangan

### Persyaratan Sistem
- Python 3.7+
- Windows 8/10/11
- Microsoft Word 2010+
- Dependencies (lihat `requirements.txt`)

### Setup Development Environment
1. Clone repository:
   ```bash
   git clone https://github.com/tirrex16/skripsweet-shortcut.git
   cd skripsweet-shortcut
   ```

2. Buat virtual environment:
   ```bash
   python -m venv venv
   .\venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Struktur Project
```
skripsweet-shortcut/
├── skripsweet_shortcut_gui.py   # File utama aplikasi
├── requirements.txt             # Dependencies
├── images/                      # Assets
│   └── skripsweet_shortcut.ico # Icon aplikasi
└── README.md                   # Dokumentasi
```

## 🔍 Panduan Pengembangan

### 1. Menambah Fitur Word
Untuk menambah fitur manipulasi Word baru:
1. Tambahkan metode baru di class `WordHelper`
2. Ikuti pola yang ada: get Word instance → get document → manipulasi
3. Gunakan try-except untuk error handling
4. Tampilkan feedback ke user via messagebox

Contoh:
```python
def new_feature(self):
    try:
        word = self.get_word_app()
        doc = self.ensure_doc(word)
        # Implementasi fitur
        messagebox.showinfo("Sukses", "Fitur berhasil dijalankan!")
    except Exception as e:
        messagebox.showerror("Error", f"Gagal: {str(e)}")
```

### 2. Menambah Elemen GUI
Untuk menambah elemen interface baru:
1. Tambahkan di method `__init__` class `Application`
2. Gunakan ttk widgets untuk konsistensi tampilan
3. Kelompokkan elemen terkait dalam frame
4. Tambahkan dokumentasi untuk setiap widget

Contoh:
```python
# Add new feature button
ttk.Button(
    main_frame,
    text="Fitur Baru",
    command=self.word_helper.new_feature,
    width=40
).pack(pady=3)
```

### 3. Modifikasi Format
Untuk mengubah format dokumen:
1. Sesuaikan konstanta di `WordHelper.__init__`
2. Ubah parameter di method `apply_default_formatting`
3. Test dengan berbagai versi Word
4. Dokumentasikan perubahan format

### 4. Error Handling
- Selalu gunakan try-except blocks
- Berikan pesan error yang informatif
- Log error untuk debugging
- Handle kasus Word tidak responsif

### 5. Testing
Test cases yang perlu diperhatikan:
- Word belum terbuka
- Dokumen kosong/baru
- Dokumen dengan konten existing
- Koneksi Word terputus
- Format dokumen tidak standar

## 🎯 Area Pengembangan Potensial

1. **Template Kustom**
   - Sistem loading template dari file
   - Editor template visual
   - Multiple template support

2. **Pengaturan Pengguna**
   - Save/load preferensi user
   - Kustomisasi margin dan format
   - Profile untuk institusi berbeda

3. **Integrasi Lanjutan**
   - Backup otomatis
   - Version control dokumen
   - Export ke format lain

4. **Peningkatan UI**
   - Dark mode
   - Tema kustom
   - Keyboard shortcuts kustom

## 🐛 Panduan Debug dan Development

### Persiapan Awal
1. **Install Python**:
   - Download Python 3.7 atau lebih baru dari [python.org](https://www.python.org/downloads/)
   - Saat instalasi, PASTIKAN centang "Add Python to PATH"
   - Verifikasi instalasi dengan membuka Command Prompt dan ketik:
     ```bash
     python --version
     ```

2. **Install Git** (opsional, untuk development):
   - Download dari [git-scm.com](https://git-scm.com/download/win)
   - Install dengan opsi default
   - Verifikasi dengan:
     ```bash
     git --version
     ```

3. **Download Source Code**:
   - Dari GitHub: 
     ```bash
     git clone https://github.com/tirrex16/skripsweet-shortcut.git
     ```
   - Atau download ZIP dari repository dan extract

4. **Setup Project**:
   ```bash
   # Buka Command Prompt sebagai Administrator
   cd "path\to\skripsweet-shortcut"
   
   # Buat virtual environment
   python -m venv venv
   
   # Aktifkan virtual environment
   .\venv\Scripts\activate
   
   # Install dependencies
   pip install -r requirements.txt
   ```

### Running untuk Debug
1. **Buka VS Code**:
   ```bash
   code .
   ```

2. **Install VS Code Extensions**:
   - Python extension
   - Pylance
   - Python Debugger

3. **Setup Debug Configuration**:
   - Tekan F5 atau klik menu Run > Start Debugging
   - Pilih "Python File"
   - VS Code akan membuat file `launch.json`
   - Tambahkan konfigurasi berikut di `launch.json`:
   ```json
   {
       "version": "0.2.0",
       "configurations": [
           {
               "name": "Python: Skripsweet Debug",
               "type": "python",
               "request": "launch",
               "program": "skripsweet_shortcut_gui.py",
               "console": "integratedTerminal",
               "justMyCode": true,
               "env": {
                   "DEBUG": "1"
               }
           }
       ]
   }
   ```

4. **Jalankan dalam Debug Mode**:
   - Tambahkan breakpoints dengan mengklik di sebelah kiri nomor baris
   - Tekan F5 untuk mulai debugging
   - Gunakan Step Over (F10), Step Into (F11), Continue (F5)
   - Watch variables di panel Debug
   - Cek debug output di Debug Console

### Common Issues dan Solusi

1. **ModuleNotFoundError**:
   - Pastikan virtual environment aktif
   - Install ulang dependencies:
     ```bash
     pip install -r requirements.txt
     ```

2. **COM Error dengan Word**:
   - Pastikan Microsoft Word terinstall
   - Jalankan VS Code sebagai Administrator
   - Restart Word jika tidak responsive

3. **Icon tidak muncul**:
   - Pastikan folder `images` ada
   - Periksa path icon di kode
   - Coba absolute path untuk testing

4. **Debug Print**:
   Tambahkan di kode untuk logging:
   ```python
   print("[Debug]", message)  # akan muncul di Debug Console
   ```

### Tips Development
1. Selalu test fitur dengan:
   - Word belum dibuka
   - Word sudah dibuka dengan dokumen
   - Word tidak responsive
   
2. Gunakan breakpoints di:
   - Event handlers
   - Error handling blocks
   - Koneksi Word
   
3. Monitor memory dan CPU usage:
   - Task Manager
   - Process Explorer
   - VS Code Performance tab

## 📝 Panduan Kontribusi

1. Fork repository
2. Buat branch untuk fitur/fix:
   ```bash
   git checkout -b feature/nama-fitur
   ```
3. Commit changes dengan pesan deskriptif
4. Push ke fork Anda
5. Submit Pull Request

### Coding Standards
- Ikuti PEP 8
- Dokumentasikan fungsi dan kelas
- Gunakan type hints
- Berikan komentar untuk logika kompleks

## 📚 Resources

- [Python Win32COM Documentation](https://docs.microsoft.com/en-us/office/client-developer/word/word-home)
- [Tkinter Documentation](https://docs.python.org/3/library/tkinter.html)
- [Word VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/word)

## 📧 Kontak Development

Untuk pertanyaan pengembangan:
- GitHub Issues
- Pull Requests
- Email: [mohammedwinston@yahoo.com]

## 📄 Lisensi

MIT License - Lihat file [LICENSE](LICENSE)

---

**Note**: Dokumentasi ini akan terus diupdate sesuai perkembangan project. Kontribusi untuk memperbaiki dokumentasi sangat diterima!
