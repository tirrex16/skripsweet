import tkinter as tk
from tkinter import ttk, messagebox
import win32com.client
import pythoncom
from datetime import datetime
import threading
import time
import os

class WordHelper:
    def __init__(self):
        self._word = None
        self.chapter_status = {i: False for i in range(1, 6)}
        
        # Constants for formatting
        self.FONT_NAME = "Times New Roman"
        self.FONT_SIZE = 12
        self.LINE_SPACING = 1.5  # 1.5 line spacing
        
        # Margin settings (in centimeters)
        self.MARGIN_TOP = 3.5
        self.MARGIN_LEFT = 3.5
        self.MARGIN_BOTTOM = 2.5
        self.MARGIN_RIGHT = 2.5
        
        # Color settings
        self.TEXT_COLOR = 0x000000  # Pure black RGB(0,0,0)
        
    def get_word_app(self):
        """Get active Word application or create new one"""
        if self._word is not None:
            try:
                # Test if Word is still accessible
                _ = self._word.Version
                print("[Debug] Using existing Word connection")
                return self._word
            except:
                self._word = None
                print("[Debug] Previous Word connection lost")
        
        try:
            print("[Debug] Trying to connect to Word...")
            # Try to get running instance first
            try:
                self._word = win32com.client.GetObject(Class="Word.Application")
                print("[Debug] Connected to running Word instance")
            except:
                print("[Debug] No running Word instance found, creating new...")
                self._word = win32com.client.Dispatch("Word.Application")
                self._word.Visible = True
                print("[Debug] Created new Word instance")
            
            # Test the connection
            version = self._word.Version
            print(f"[Debug] Connected to Word version: {version}")
            return self._word
            
        except Exception as e:
            error_msg = f"Tidak dapat terhubung ke Microsoft Word: {str(e)}\n\nPastikan:\n1. Microsoft Word sudah terinstall\n2. Word tidak sedang busy/not responding\n3. Coba tutup dan buka ulang Word"
            messagebox.showerror("Error", error_msg)
            raise
            
    def ensure_doc(self, word):
        """Get active document or create new one"""
        try:
            print("[Debug] Checking for active document...")
            if word.Documents.Count > 0:
                doc = word.ActiveDocument
                print(f"[Debug] Found active document: {doc.Name}")
                self.apply_default_formatting(doc)
                return doc
            else:
                print("[Debug] No documents open")
                messagebox.showinfo("Info", "Membuat dokumen baru...")
                doc = word.Documents.Add()
                self.apply_default_formatting(doc)
                print("[Debug] Created new document")
                return doc
        except Exception as e:
            print(f"[Debug] Error accessing document: {str(e)}")
            messagebox.showinfo("Info", "Membuat dokumen baru karena error...")
            try:
                doc = word.Documents.Add()
                self.apply_default_formatting(doc)
                return doc
            except Exception as e2:
                error_msg = f"Gagal membuat dokumen: {str(e2)}\n\nPastikan:\n1. Word tidak sedang busy\n2. Coba tutup dan buka ulang Word"
                messagebox.showerror("Error", error_msg)
                raise
                
    def apply_default_formatting(self, doc):
        """Apply default formatting to the document"""
        try:
            # Set default font for the whole document
            doc.Content.Font.Name = self.FONT_NAME
            doc.Content.Font.Size = self.FONT_SIZE
            doc.Content.Font.Color = 0x000000  # Pure black
            
            # Set line spacing
            doc.Content.ParagraphFormat.LineSpacingRule = 1  # wdLineSpace1pt5
            doc.Content.ParagraphFormat.LineSpacing = self.LINE_SPACING * 12  # 12 points = 1 line
            
            # Setup styles
            # Heading 1 style (for chapter titles)
            try:
                h1 = doc.Styles("Heading 1")
                h1.Font.Name = self.FONT_NAME
                h1.Font.Size = 14
                h1.Font.Bold = True  # Always bold for BAB and titles
                h1.Font.AllCaps = True
                h1.Font.Color = 0x000000  # Pure black
                h1.ParagraphFormat.Alignment = 1  # Center
                h1.ParagraphFormat.LineSpacingRule = 1
                h1.ParagraphFormat.LineSpacing = self.LINE_SPACING * 12
                h1.ParagraphFormat.SpaceBefore = 24  # 2 lines before
                h1.ParagraphFormat.SpaceAfter = 24   # 2 lines after
            except: pass
            
            # Heading 2 style (for sections)
            try:
                h2 = doc.Styles("Heading 2")
                h2.Font.Name = self.FONT_NAME
                h2.Font.Size = 12
                h2.Font.Bold = True  # Always bold for sections
                h2.Font.Color = 0x000000  # Pure black
                h2.ParagraphFormat.Alignment = 0  # Left align
                h2.ParagraphFormat.LineSpacingRule = 1
                h2.ParagraphFormat.LineSpacing = self.LINE_SPACING * 12
                h2.ParagraphFormat.SpaceBefore = 12  # 1 line before
                h2.ParagraphFormat.SpaceAfter = 12   # 1 line after
                h2.ParagraphFormat.LeftIndent = 0    # No indent for section headings
            except: pass
            
            # Heading 3 style (for subsections)
            try:
                h3 = doc.Styles("Heading 3")
                h3.Font.Name = self.FONT_NAME
                h3.Font.Size = 12
                h3.Font.Bold = True
                h3.ParagraphFormat.LineSpacingRule = 1
                h3.ParagraphFormat.LineSpacing = self.LINE_SPACING * 12
                h3.ParagraphFormat.SpaceBefore = 12
                h3.ParagraphFormat.SpaceAfter = 12
                h3.ParagraphFormat.LeftIndent = 36   # Indent subsections
            except: pass
            
            # Normal style
            try:
                normal = doc.Styles("Normal")
                normal.Font.Name = self.FONT_NAME
                normal.Font.Size = 12
                normal.ParagraphFormat.LineSpacingRule = 1
                normal.ParagraphFormat.LineSpacing = self.LINE_SPACING * 12
                normal.ParagraphFormat.FirstLineIndent = 36  # 0.5 inch indent
            except: pass
            
        except Exception as e:
            print(f"[Debug] Warning: Could not apply some formatting: {str(e)}")

    def create_toc(self):
        """Buat daftar isi"""
        try:
            print("\n[Debug] === Creating Table of Contents ===")
            word = self.get_word_app()
            word.Visible = True
            doc = self.ensure_doc(word)
            sel = word.Selection
            
            print("[Debug] Adding DAFTAR ISI text...")
            sel.TypeText("DAFTAR ISI")
            sel.TypeParagraph()
            sel.Style = doc.Styles("Heading 1")
            sel.ParagraphFormat.Alignment = 1  # Center
            sel.TypeParagraph()
            
            print("[Debug] Creating Table of Contents...")
            toc_range = sel.Range
            
            try:
                toc = doc.TablesOfContents.Add(
                    toc_range, 
                    True,   # UseHeadingStyles
                    1,      # UpperLevel
                    3,      # LowerLevel
                    True,   # RightAlignPageNumbers
                    True,   # IncludePageNumbers
                    True    # UseHyperlinks
                )
                
                # Format TOC
                toc_range = toc.Range
                toc_range.Font.Name = self.FONT_NAME
                toc_range.Font.Size = self.FONT_SIZE
                
                print("[Debug] Updating TOC...")
                toc.Update()
                
            except Exception as e:
                print(f"[Debug] Error in TOC creation: {str(e)}")
                raise
            
            print("[Debug] TOC creation successful!")
            messagebox.showinfo("Sukses", "Daftar isi berhasil dibuat!")
            
        except Exception as e:
            error_msg = f"Gagal membuat daftar isi:\n{str(e)}\n\nCoba:\n1. Pastikan Word tidak sedang busy\n2. Simpan dokumen terlebih dahulu\n3. Tutup dan buka ulang Word"
            print(f"[Debug] Error: {str(e)}")
            messagebox.showerror("Error", error_msg)

    def create_specific_bab(self, bab_number):
        """Buat BAB dengan nomor tertentu"""
        try:
            word = self.get_word_app()
            doc = self.ensure_doc(word)
            sel = word.Selection
            
            rn = self.to_roman(bab_number)
            
            # Add centered BAB heading and title together
            sel.TypeText(f"BAB {rn}")
            sel.TypeParagraph()
            
            # Get chapter title
            titles = {
                1: "PENDAHULUAN",
                2: "TINJAUAN PUSTAKA",
                3: "METODE PENELITIAN",
                4: "HASIL DAN PEMBAHASAN",
                5: "PENUTUP"
            }
            title = titles.get(bab_number, "JUDUL BAB")
                
            # Format both BAB and title with same style
            sel.TypeText(title)
            sel.TypeParagraph()
            
            # Select both BAB and title together
            sel.MoveUp(Unit=4, Count=2)  # Move up to start of BAB
            sel.MoveDown(Unit=4, Count=2, Extend=1)  # Select down to end of title
            
            # Apply formatting to the whole selection
            sel.Style = "Heading 1"
            sel.ParagraphFormat.Alignment = 1  # Center
            sel.Font.Bold = True  # Make both bold
            sel.Font.Color = 0x000000  # Pure black
            sel.Font.Size = 14  # Same size
            
            # Move to end and add spacing
            sel.MoveDown(Unit=4, Count=1)
            sel.TypeParagraph()  # Extra line after title
            sel.TypeParagraph()
            
            # Add sections and subsections
            sections_map = {
                1: {
                    "Latar Belakang": [
                        "Latar Belakang Masalah",
                        "Identifikasi Masalah"
                    ],
                    "Rumusan Masalah": [],
                    "Tujuan Penelitian": [],
                    "Batasan Masalah": [],
                    "Manfaat Penelitian": [
                        "Manfaat Teoritis",
                        "Manfaat Praktis"
                    ]
                },
                2: {
                    "Tinjauan Pustaka": [
                        "Penelitian Terdahulu",
                        "State of the Art"
                    ],
                    "Landasan Teori": [],
                    "Kerangka Pemikiran": []
                },
                3: {
                    "Metode Penelitian": [
                        "Jenis Penelitian",
                        "Pendekatan Penelitian"
                    ],
                    "Prosedur Penelitian": [
                        "Tahap Persiapan",
                        "Tahap Pelaksanaan",
                        "Tahap Analisis"
                    ],
                    "Instrumen Penelitian": [],
                    "Teknik Pengumpulan Data": [],
                    "Teknik Analisis Data": []
                },
                4: {
                    "Hasil Penelitian": [
                        "Deskripsi Data",
                        "Analisis Data"
                    ],
                    "Pembahasan": [
                        "Interpretasi Hasil",
                        "Diskusi Temuan"
                    ],
                    "Keterbatasan Penelitian": []
                },
                5: {
                    "Kesimpulan": [],
                    "Saran": [
                        "Saran Teoretis",
                        "Saran Praktis"
                    ]
                }
            }

            sections = sections_map.get(bab_number, {"Subbab 1": [], "Subbab 2": []})
                
            # Add sections and subsections
            for i, (section, subsections) in enumerate(sections.items(), 1):
                # Add section heading (e.g., 1.1, 2.1, etc.)
                sel.TypeText(f"{bab_number}.{i}. {section}")
                sel.TypeParagraph()
                sel.Style = "Heading 2"
                sel.ParagraphFormat.Alignment = 0  # Left align
                sel.Font.Bold = True  # Make subsection bold
                sel.Font.Color = 0x000000  # Pure black
                
                if not subsections:  # If no subsections, add placeholder
                    sel.TypeText("[[tulis isi di sini]]")
                    sel.Style = "Normal"  # Reset to normal style for content
                    sel.ParagraphFormat.Alignment = 0  # Left align content
                    sel.TypeParagraph()
                    sel.TypeParagraph()
                else:
                    # Add subsections
                    for j, subsection in enumerate(subsections, 1):
                        sel.TypeText(f"{bab_number}.{i}.{j}. {subsection}")
                        sel.TypeParagraph()
                        sel.Style = "Heading 3"
                        sel.ParagraphFormat.Alignment = 0  # Left align
                        sel.Font.Bold = True  # Make subsection bold
                        sel.Font.Color = 0x000000  # Pure black
                        
                        sel.TypeText("[[tulis isi di sini]]")
                        sel.Style = "Normal"  # Reset to normal style for content
                        sel.ParagraphFormat.Alignment = 0  # Left align content
                        sel.TypeParagraph()
                        sel.TypeParagraph()
                
            # Mark chapter as created and check if all chapters exist
            self.chapter_status[bab_number] = True
            if all(self.chapter_status.values()):
                self.enable_list_buttons()
            
            messagebox.showinfo("Sukses", f"BAB {rn} berhasil dibuat!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuat BAB {rn}: {str(e)}")

    def create_bab(self):
        """For backwards compatibility"""
        self.create_specific_bab(1)

    def create_list_tables(self):
        """Buat daftar tabel"""
        try:
            word = self.get_word_app()
            doc = self.ensure_doc(word)
            sel = word.Selection
            
            # Add empty line before
            sel.TypeParagraph()
            
            # Add heading
            sel.TypeText("DAFTAR TABEL")
            sel.TypeParagraph()
            sel.Style = "Heading 1"
            sel.ParagraphFormat.Alignment = 1  # Center
            sel.Font.Color = 0x000000  # Pure black
            sel.TypeParagraph()
            
            # Add table of figures
            rng = sel.Range
            doc.TablesOfFigures.Add(
                Range=rng,
                Caption="Tabel",
                IncludeLabel=True,
                RightAlignPageNumbers=True,
                UseHeadingStyles=False,
                UseHyperlinks=True
            )
            
            messagebox.showinfo("Sukses", "Daftar tabel berhasil dibuat!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuat daftar tabel: {str(e)}")

    def create_list_figures(self):
        """Buat daftar gambar"""
        try:
            word = self.get_word_app()
            doc = self.ensure_doc(word)
            sel = word.Selection
            
            # Add empty line before
            sel.TypeParagraph()
            
            # Add heading
            sel.TypeText("DAFTAR GAMBAR")
            sel.TypeParagraph()
            sel.Style = "Heading 1"
            sel.ParagraphFormat.Alignment = 1  # Center
            sel.Font.Color = 0x000000  # Pure black
            sel.TypeParagraph()
            
            # Add table of figures
            rng = sel.Range
            doc.TablesOfFigures.Add(
                Range=rng,
                Caption="Gambar",
                IncludeLabel=True,
                RightAlignPageNumbers=True,
                UseHeadingStyles=False,
                UseHyperlinks=True
            )
            
            messagebox.showinfo("Sukses", "Daftar gambar berhasil dibuat!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuat daftar gambar: {str(e)}")

    def create_template(self):
        """Buat template skripsi baru"""
        try:
            word = self.get_word_app()
            word.Visible = True
            doc = word.Documents.Add()
            
            # Set page setup
            print("[Debug] Setting up page format...")
            self.set_page_setup()
            
            # Apply default formatting
            self.apply_default_formatting(doc)
            
            # Add cover page
            sel = word.Selection
            sel.TypeText("JUDUL SKRIPSI")
            sel.TypeParagraph()
            sel.ParagraphFormat.Alignment = 1  # Center
            sel.Font.Bold = True
            sel.Font.Size = 14
            sel.TypeParagraph()
            sel.TypeParagraph()
            
            sel.TypeText("Oleh:")
            sel.TypeParagraph()
            sel.TypeText("[NAMA MAHASISWA]")
            sel.TypeParagraph()
            sel.TypeText("[NIM]")
            sel.TypeParagraph()
            sel.TypeParagraph()
            sel.TypeParagraph()
            
            sel.TypeText("PROGRAM STUDI [NAMA PRODI]")
            sel.TypeParagraph()
            sel.TypeText("FAKULTAS [NAMA FAKULTAS]")
            sel.TypeParagraph()
            sel.TypeText("[NAMA UNIVERSITAS]")
            sel.TypeParagraph()
            sel.TypeText("[TAHUN]")
            
            # Add sections
            sections = [
                "LEMBAR PERSETUJUAN",
                "LEMBAR PENGESAHAN",
                "PERNYATAAN KEASLIAN",
                "KATA PENGANTAR",
                "DAFTAR ISI",
                "DAFTAR TABEL",
                "DAFTAR GAMBAR",
                "DAFTAR LAMPIRAN",
                "ABSTRAK",
                "ABSTRACT"
            ]
            
            for section in sections:
                sel.InsertBreak(2)  # Page break
                sel.TypeText(section)
                sel.TypeParagraph()
                sel.Style = "Heading 1"
                sel.ParagraphFormat.Alignment = 1  # Center
                sel.TypeParagraph()
                sel.TypeText("[[Isi " + section.lower() + " di sini]]")
                sel.TypeParagraph()
            
            messagebox.showinfo("Sukses", "Template skripsi berhasil dibuat!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuat template: {str(e)}")
            
    def format_selection(self):
        """Format teks yang dipilih dengan format standar"""
        try:
            word = self.get_word_app()
            doc = self.ensure_doc(word)
            sel = word.Selection
            
            # Apply standard formatting
            sel.Font.Name = self.FONT_NAME
            sel.Font.Size = self.FONT_SIZE
            sel.ParagraphFormat.LineSpacingRule = 1
            sel.ParagraphFormat.LineSpacing = self.LINE_SPACING * 12
            
            messagebox.showinfo("Sukses", "Format teks berhasil diubah!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengubah format: {str(e)}")
            
    def set_page_setup(self):
        """Set page setup untuk skripsi"""
        try:
            word = self.get_word_app()
            doc = self.ensure_doc(word)
            
            # Constants
            CM_TO_PT = 28.3464567
            WD_PAPER_A4 = 7
            
            # Set page setup
            ps = doc.PageSetup
            ps.PaperSize = WD_PAPER_A4
            ps.TopMargin = self.MARGIN_TOP * CM_TO_PT
            ps.BottomMargin = self.MARGIN_BOTTOM * CM_TO_PT
            ps.LeftMargin = self.MARGIN_LEFT * CM_TO_PT
            ps.RightMargin = self.MARGIN_RIGHT * CM_TO_PT
            
            # Force Portrait orientation (0 = Portrait, 1 = Landscape)
            ps.Orientation = 0
            
            # Set page numbers
            ps.DifferentFirstPageHeaderFooter = True  # Different first page
            ps.FooterDistance = 1.27 * CM_TO_PT  # 1.27 cm from bottom
            
            # Set default text color to black
            doc.Content.Font.Color = self.TEXT_COLOR
            
            messagebox.showinfo("Sukses", "Page setup berhasil diatur!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengatur page setup: {str(e)}")

    @staticmethod
    def to_roman(num):
        """Convert number to roman numeral"""
        vals = [
            ("M",1000),("CM",900),("D",500),("CD",400),
            ("C",100),("XC",90),("L",50),("XL",40),
            ("X",10),("IX",9),("V",5),("IV",4),("I",1)
        ]
        result = []
        for sym, val in vals:
            while num >= val:
                result.append(sym)
                num -= val
        return "".join(result)

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("Skripsweet Shortcut untuk Skripsi")
        self.word_helper = WordHelper()
        
        # Load application icon if it exists
        try:
            icon_path = "images/skripsweet_shortcut.ico"
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except:
            pass  # If icon loading fails, use default icon
            
        # Set window size and position
        window_width = 500
        window_height = 650
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # Track chapter completion
        self.chapter_status = {i: False for i in range(1, 6)}  # Track which chapters are created
        
        # Create main frame with padding
        main_frame = ttk.Frame(self, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create Word status frame
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=5)
        
        # Add status indicator
        self.word_status_label = ttk.Label(
            status_frame,
            text="Status Word: ",
            font=("Helvetica", 10)
        )
        self.word_status_label.pack(side=tk.LEFT, padx=5)
        
        self.word_status_indicator = ttk.Label(
            status_frame,
            text="Mengecek...",
            font=("Helvetica", 10, "bold"),
            foreground="gray"
        )
        self.word_status_indicator.pack(side=tk.LEFT)
        
        # Add refresh button
        ttk.Button(
            status_frame,
            text="↻",
            width=3,
            command=self.check_word_status
        ).pack(side=tk.RIGHT, padx=5)
        
        # Start the status check timer
        self.after(1000, self.check_word_status)  # Start checking after 1 second
        
        # Add description
        ttk.Label(
            main_frame,
            text="Klik tombol di bawah untuk membuat bagian skripsi:",
            wraplength=600,
            font=("Helvetica", 10)
        ).pack(side=tk.TOP, pady=5)
        
        # Add description label
        ttk.Label(
            main_frame,
            text="Skripsweet Shortcut untuk Skripsi",
            font=("Helvetica", 16, "bold")
        ).pack(pady=10)

        # Add label for page setup
        ttk.Label(
            main_frame,
            text="Pengaturan Halaman",
            font=("Helvetica", 12, "bold")
        ).pack(pady=10)
        
        ttk.Button(
            main_frame,
            text="Atur Page Setup (A4, Margin: T/L=3.5cm, B/R=2.5cm)",
            command=self.word_helper.set_page_setup,
            width=50
        ).pack(pady=5)
        
        # Add separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Add label for preliminary pages
        ttk.Label(
            main_frame,
            text="Halaman Awal",
            font=("Helvetica", 12, "bold")
        ).pack(pady=10)
        
        # Create list buttons with direct commands
        prelim_buttons = [
            ("Buat DAFTAR ISI", self.word_helper.create_toc),
            ("Buat DAFTAR TABEL", self.word_helper.create_list_tables),
            ("Buat DAFTAR GAMBAR", self.word_helper.create_list_figures)
        ]
        
        for text, command in prelim_buttons:
            ttk.Button(
                main_frame,
                text=text,
                command=command,
                width=40
            ).pack(pady=3)
            
        # Add separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Add label for chapters
        ttk.Label(
            main_frame,
            text="Bab-bab Skripsi",
            font=("Helvetica", 12, "bold")
        ).pack(pady=10)
        
        # Add buttons for each chapter
        chapter_titles = {
            1: "PENDAHULUAN",
            2: "TINJAUAN PUSTAKA",
            3: "METODE PENELITIAN",
            4: "HASIL DAN PEMBAHASAN",
            5: "PENUTUP"
        }
        
        for i in range(1, 6):
            ttk.Button(
                main_frame,
                text=f"Buat BAB {self.word_helper.to_roman(i)} - {chapter_titles[i]}",
                command=lambda x=i: self.word_helper.create_specific_bab(x),
                width=50
            ).pack(pady=3)
            
        # Add status bar
        self.status = ttk.Label(
            main_frame,
            text="Siap digunakan",
            font=("Helvetica", 10),
            foreground="green"
        )
        self.status.pack(side=tk.BOTTOM, pady=10)
    
    def check_word_status(self):
        """Check if Word is running and accessible"""
        if hasattr(self, 'word_status_indicator'):
            try:
                # Try to get Word instance without showing error dialog
                try:
                    word = win32com.client.GetObject(Class="Word.Application")
                    self.word_status_indicator.config(
                        text="Terhubung",
                        foreground="green"
                    )
                except pythoncom.com_error:
                    self.word_status_indicator.config(
                        text="Tidak Terhubung",
                        foreground="red"
                    )
            except Exception as e:
                print(f"Error checking Word status: {str(e)}")
                self.word_status_indicator.config(
                    text="Error",
                    foreground="red"
                )
            
            # Schedule next check in 5 seconds
            self.after(5000, self.check_word_status)

if __name__ == "__main__":
    app = Application()
    app.mainloop()