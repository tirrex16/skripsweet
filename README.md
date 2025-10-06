# Skripsweet Shortcut - Development Guide 🛠️

Dokumentasi pengembangan untuk Skripsweet Shortcut, aplikasi GUI yang membantu penulisan skripsi di Microsoft Word.

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

## 🐛 Debug Mode

Untuk mengaktifkan debug mode:
1. Set environment variable:
   ```bash
   set DEBUG=1
   ```
2. Jalankan aplikasi:
   ```bash
   python skripsweet_shortcut_gui.py
   ```
3. Cek output di console untuk log detail

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
