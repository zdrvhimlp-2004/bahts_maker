"""Antarmuka grafis (satu berkas) untuk mesin bahts: semua perintah `python mesin/bahts.py ...` jadi tombol.

Jalankan:  python bahts_gui.py      (atau klik dua kali di Windows)
Butuh Python 3 dengan tkinter (sudah ada di pemasangan Python resmi untuk Windows).
Berkas ini harus ada di folder utama repo (sejajar dengan folder `mesin`).
"""
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

ROOT = Path(__file__).resolve().parent
BAHTS = ROOT / "mesin" / "bahts.py"
SAMPUL = ROOT / "sampul_saya.txt"
NAMA_FOLDER = re.compile(r"^[A-Za-z0-9_-]+$")

# Isi tab «Panduan & konsep»: (gaya, teks). h1 = judul bagian, h2 = judul tombol/konsep, p = paragraf.
PANDUAN = [
    ("h1", "1. Gambaran besar"),
    ("p", "Mesin ini membantu menulis bahts (riset ilmiah berbahasa Arab) dengan satu prinsip: setiap kutipan harus bisa "
          "dibuktikan ke halaman kitabnya. Anda (atau agen AI) hanya menulis satu berkas, naskah.txt. Sisanya dikerjakan mesin:\n"
          "   naskah.txt  →  cek (buktikan semua kutipan)  →  rakit (susun Word + PDF)\n"
          "Satu bahts = satu folder. Folder kitab/ berisi kitab yang sudah diunduh dan dipakai bersama oleh semua bahts."),
    ("p", "Tiap tombol di jendela ini hanya menjalankan perintah  python mesin/bahts.py <perintah>  lalu menampilkan hasilnya "
          "di kotak «Hasil». Jadi apa yang bisa dilakukan GUI ini bisa juga dilakukan dari terminal."),

    ("h1", "2. Tab Persiapan"),
    ("h2", "Periksa"),
    ("p", "Hanya memeriksa, tidak mengubah apa pun. Mengecek: versi Python (3.12+), pustaka Python (pandas, pyarrow, pypdf, "
          "fontTools), katalog kitab, sistem Windows, Microsoft Word desktop, font mushaf, huruf Traditional Arabic, dan huruf "
          "Sakkal Majalla. Tiap baris BELUM disertai perintah perbaikannya. Di akhir ia menyimpulkan dua hal: «Menulis dan cek» "
          "(butuh Python, pustaka, katalog) dan «rakit» (butuh juga Windows + Word + font)."),
    ("h2", "Pasang pustaka Python"),
    ("p", "Menjalankan  pip install -r requirements.txt. Pustakanya: pandas dan pyarrow (membaca katalog kitab yang berformat "
          "parquet), pypdf dan fontTools (memperbaiki teks di dalam PDF hasil Word dan membaca berkas huruf)."),
    ("h2", "Siapkan"),
    ("p", "Dua pekerjaan sekali saja: (1) mengunduh katalog kitab Shamela (berkas _meta/book_metadata.parquet dari dataset "
          "Hugging Face AuthenticIlm/Shamela4_Full_DB, sekitar 8.500 judul) ke folder kitab/; (2) mengunduh font mushaf "
          "«KFGQPC HAFS Uthmanic Script» (milik Mujamma' Raja Fahd) dari static-cdn.tarteel.ai atau cadangan di GitHub, memeriksa "
          "sidik SHA-256-nya, lalu memasangnya untuk akun Windows Anda saja. Font ini dipakai untuk mencetak ayat."),
    ("h2", "Daftar tanda naskah"),
    ("p", "Mencetak semua tanda yang dipahami mesin di naskah.txt: @nama:, #K, #F, #B, #M, #S, #R, {Q:surat:ayat}, [^catatan kaki], "
          "<<id kitab|juz|halaman>>, dan lain-lain."),

    ("h1", "3. Tab Bahts"),
    ("h2", "Bahts baru"),
    ("p", "Membuat folder dan kerangka naskah.txt dari mesin/naskah_baru.txt. Data sampul (nama, NIM, jenis, semester) diambil "
          "otomatis dari sampul_saya.txt. Nama folder berhuruf Latin tanpa spasi; judul berhuruf Arab."),
    ("h2", "Cek naskah"),
    ("p", "Inti keamanan mesin. Langkahnya: (1) membaca naskah dan memeriksa tanda-tandanya; (2) mencocokkan tiap {Q:...} dengan "
          "teks mushaf Hafs yang tersimpan di mesin/quran_hafs_v22.json; (3) untuk tiap bukti <<id|juz|halaman>>, membuka halaman "
          "kitab itu di folder kitab/ dan memastikan kutipan bertanda petik di depan catatan kaki ada di halaman tersebut, utuh "
          "per kata. Sebelum membandingkan, kedua teks disamakan dulu (mesin/norma.py): harakat dibuang, bentuk hamzah/ya/ta "
          "marbuthah disamakan, kalimat pujian seperti «صلى الله عليه وسلم» diabaikan. Hasil: GAGAL (harus 0), «perhatian» (lolos "
          "tetapi harus dilihat sendiri), dan perkiraan jumlah halaman."),
    ("p", "Yang TIDAK bisa dinilai cek: apakah yang berkata memang orang yang Anda sebut, apakah nama kitab benar, dan kutipan "
          "yang tidak diberi catatan kaki. Itu tugas pembacaan manusia."),
    ("h2", "Rakit Word + PDF"),
    ("p", "Pertama menjalankan cek. Kalau lolos: (1) olah.py mengubah naskah menjadi berkas bahan JSON di mesin/_tmp/; "
          "(2) rakit.ps1 (PowerShell) mengendalikan Microsoft Word lewat COM Automation: membuka mesin/sampul.doc, mengisi data "
          "sampul, menyusun isi, judul bertingkat, catatan kaki, nomor halaman, daftar isi, فهرس الآيات dan فهرس الأحاديث, lalu "
          "المصادر والمراجع; bentuk halaman dan huruf mengikuti profil di mesin/format.json (bawaan mengikuti pedoman Jami'at "
          "al-Imam); (3) Word menyimpan .docx (SaveAs2) dan mengekspor PDF (ExportAsFixedFormat); (4) teks_pdf.py memperbaiki "
          "peta huruf di PDF memakai pypdf dan fontTools supaya teks Arab bisa disalin dan dicari dengan benar. Tanpa Word "
          "langkah ini mustahil, dan mesin sengaja tidak memakai pengganti Word."),
    ("h2", "Buka naskah / Buka folder"),
    ("p", "Membuka naskah.txt di editor bawaan Windows atau membuka folder bahts di Explorer. Perubahan isi selalu lewat "
          "naskah.txt; .docx hanya hasil jadi."),

    ("h1", "4. Tab Cari kitab: konsep dan sumber datanya"),
    ("p", "Ada dua lapis. Lapis lokal (offline) dipakai oleh cek dan rakit, sehingga hasil akhirnya tidak bergantung pada "
          "layanan orang lain. Lapis online hanya penunjuk jalan untuk menemukan kitab mana yang perlu dibaca."),
    ("h2", "Lapis lokal: perpustakaan Shamela di komputer Anda"),
    ("p", "Sumbernya dataset Hugging Face AuthenticIlm/Shamela4_Full_DB (salinan data Maktabah Shamela, tiap kitab terpisah). "
          "Perintah Ambil mengunduh satu kitab (book_metadata.json, toc.jsonl, pages.jsonl) ke kitab/<nomor>/ dalam bentuk "
          "terkompresi (pages.jsonl.gz) lewat HTTPS biasa dengan modul urllib bawaan Python, dengan percobaan ulang bila putus. "
          "Setelah terunduh, Cari, Teks, Halaman, Bab, dan Kartu bekerja sepenuhnya offline: teks tiap halaman dinormalisasi "
          "(norma.py) lalu disimpan sebagai cache supaya pencarian cepat. Nomor kitab = shamela_id."),
    ("h2", "Katalog"),
    ("p", "Mencari judul atau nama pengarang di berkas katalog lokal (parquet) memakai pandas; ejaan disamakan dan pola boleh "
          "regex. Menampilkan nomor kitab dan tanda ADA bila kitabnya sudah diunduh. Maksimal 30 baris."),
    ("h2", "Kitab terunduh"),
    ("p", "Mendaftar kitab yang sudah ada di folder kitab/."),
    ("h2", "Ambil (unduh)"),
    ("p", "Mengunduh kitab bernomor itu sekali saja dari dataset di atas. Kitab besar bisa makan waktu beberapa menit."),
    ("h2", "Kartu (data cetak)"),
    ("p", "Menampilkan judul, pengarang, muhaqqiq, penerbit, cetakan, dan keterangan penomoran dari kitab lokal. Dipakai untuk "
          "catatan kaki rujukan pertama dan daftar pustaka. Perhatikan baris «ترقيم الكتاب موافق للمطبوع»: bila tertulis nomor "
          "halamannya otomatis, nomornya bukan nomor cetakan."),
    ("h2", "Cari frasa di kitab ini"),
    ("p", "Mencari frasa di satu kitab lokal dan menampilkan juz dan halaman cetaknya. Pencarian tidak peka harakat dan "
          "variasi hamzah, jadi tulis frasa pendek seperti di kitab."),
    ("h2", "Teks sekitar frasa"),
    ("p", "Menampilkan teks asli (lengkap dengan harakatnya) di sekitar frasa, untuk disalin persis ke naskah."),
    ("h2", "Tampilkan halaman"),
    ("p", "Menampilkan isi satu halaman: nomor kitab, juz (isi - bila satu jilid), dan halaman. Ini cara membaca konteks "
          "kutipan sebelum dipakai."),
    ("h2", "Temukan di semua kitab (online)"),
    ("p", "Mencari frasa di isi SELURUH kitab Shamela lewat API pencarian turath.io (api.turath.io/search). Mengembalikan "
          "kitab, juz, halaman, dan cuplikan (maksimal 20 hasil), sehingga Anda tahu kitab mana yang perlu di-Ambil. Butuh "
          "internet, tidak resmi, dan berjatah: jangan dipanggil ratusan kali berturut-turut. Hasil kosong bukan bukti frasa "
          "tidak ada. Kitabnya tetap harus diambil dan dibaca; kutipannya tetap harus lolos cek."),
    ("h2", "Penunjuk jalan lain (hanya di terminal)"),
    ("p", "hadits «طرف»: mencari tempat hadits di 12 kitab hadits lewat المنصة الحديثية (alminasa.ai). dirasat «kata kunci»: "
          "mencari penelitian terdahulu di katalog OpenAlex (menurut kata dan menurut makna) dan DOAJ. dirasat-pdf: mengunduh "
          "PDF penelitian dari OpenAlex (butuh kunci API gratis di kunci_openalex.txt). scan: memberi alamat PDF scan cetakan "
          "halaman di files.turath.io. Semua ini layanan pihak ketiga yang bisa berubah atau mati tanpa kabar."),
    ("h2", "Connector (khusus agen AI)"),
    ("p", "Berkas .mcp.json mendaftarkan connector seperti shamela, turath, hadithunlocked, tafsir, dan quran yang bisa dipakai "
          "agen AI untuk mencari bahan. Jendela ini tidak memakainya."),

    ("h1", "5. Tab Data sampul"),
    ("p", "Mengedit sampul_saya.txt (@nama, @nim, @jenis, @fasl, dan @thalibah: ya bila mahasiswi). Berkas ini dibaca perintah "
          "«Bahts baru» untuk mengisi sampul tiap bahts. Pembimbing (@musyrif, @jabatan) ditulis di naskah.txt masing-masing "
          "bahts. Sampul bawaan mencetak identitas kampus tertentu (mesin/sampul.doc); kampus lain perlu sampul sendiri."),

    ("h1", "6. Batasan yang perlu diingat"),
    ("p", "• Teks Shamela adalah ketikan manusia dan bisa salah ketik; bandingkan dengan edisi lain atau scan bila ragu.\n"
          "• Lolos cek berarti kutipan ada di halaman itu, bukan berarti penisbatannya benar. Baca konteksnya.\n"
          "• Jumlah halaman pasti baru diketahui setelah Rakit; angka di cek hanya perkiraan kasar.\n"
          "• Nama di sampul adalah tanggung jawab Anda. Tanyakan ke dosen apakah bantuan AI dibolehkan.\n"
          "• Rakit hanya jalan di Windows dengan Microsoft Word desktop (bukan Word Online, Google Docs, atau LibreOffice)."),

    ("h1", "7. Peta berkas"),
    ("p", "mesin/            kode mesin (jangan diisi bahan bahts)\n"
          "kitab/            kitab Shamela yang sudah diunduh, dipakai semua bahts\n"
          "<folder bahts>/   naskah.txt, lalu .docx dan .pdf hasil Rakit\n"
          "sampul_saya.txt   data sampul pribadi (tidak ikut repo)\n"
          "kunci_openalex.txt  kunci API OpenAlex (opsional)\n"
          "bahts_gui.py      jendela ini    |    Mesin Bahts.bat  peluncurnya"),
]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mesin Bahts")
        self.geometry("980x720")
        self.minsize(760, 560)
        self.antre = queue.Queue()
        self.sibuk = False
        self.proses = None

        if not BAHTS.exists():
            messagebox.showerror("Mesin Bahts", f"Tidak menemukan {BAHTS}.\nLetakkan berkas ini di folder utama repo.")
            self.destroy()
            return

        atas = ttk.Frame(self, padding=(8, 8, 8, 0))
        atas.pack(fill="x")
        ttk.Label(atas, text="Folder repo:").pack(side="left")
        ttk.Label(atas, text=str(ROOT)).pack(side="left", padx=6)

        # atas = tab-tab, bawah = hasil perintah; garis pemisahnya bisa digeser
        belah = ttk.PanedWindow(self, orient="vertical")
        belah.pack(fill="both", expand=True, padx=8, pady=8)
        self.buku = ttk.Notebook(belah)
        belah.add(self.buku, weight=3)
        self._tab_persiapan()
        self._tab_bahts()
        self._tab_kitab()
        self._tab_sampul()
        self._tab_panduan()

        bawah = ttk.Frame(belah)
        belah.add(bawah, weight=2)
        bar = ttk.Frame(bawah)
        bar.pack(fill="x")
        ttk.Label(bar, text="Hasil").pack(side="left")
        self.status = ttk.Label(bar, text="siap")
        self.status.pack(side="right")
        ttk.Button(bar, text="Hentikan", command=self.hentikan).pack(side="right", padx=6)
        ttk.Button(bar, text="Bersihkan", command=self.bersihkan).pack(side="right")
        isi = ttk.Frame(bawah)
        isi.pack(fill="both", expand=True, pady=(4, 0))
        self.keluaran = tk.Text(isi, wrap="word", font=("Consolas", 10), state="disabled")
        gulung = ttk.Scrollbar(isi, command=self.keluaran.yview)
        self.keluaran.configure(yscrollcommand=gulung.set)
        gulung.pack(side="right", fill="y")
        self.keluaran.pack(side="left", fill="both", expand=True)

        self.segarkan_daftar()
        self.after(100, self._kuras)

    # ---------- tab-tab ----------
    def _tab_persiapan(self):
        t = ttk.Frame(self.buku, padding=10)
        self.buku.add(t, text="Persiapan")
        ttk.Label(t, wraplength=880, justify="left", text=(
            "Urutan pertama kali: 1) Periksa  2) Pasang pustaka Python (kalau ada yang BELUM)  3) Siapkan "
            "(unduh katalog kitab dan pasang font mushaf).  Lalu Periksa lagi sampai semua SIAP.")).pack(anchor="w")
        b = ttk.Frame(t)
        b.pack(anchor="w", pady=8)
        ttk.Button(b, text="1. Periksa", command=lambda: self.jalankan(["periksa"])).pack(side="left", padx=4)
        ttk.Button(b, text="2. Pasang pustaka Python", command=self.pasang_pustaka).pack(side="left", padx=4)
        ttk.Button(b, text="3. Siapkan", command=lambda: self.jalankan(["siapkan"])).pack(side="left", padx=4)
        ttk.Button(b, text="Daftar tanda naskah", command=lambda: self.jalankan(["tanda"])).pack(side="left", padx=4)

    def _tab_bahts(self):
        t = ttk.Frame(self.buku, padding=10)
        self.buku.add(t, text="Bahts")
        kiri = ttk.Frame(t)
        kiri.pack(side="left", fill="y")
        ttk.Label(kiri, text="Bahts yang ada:").pack(anchor="w")
        self.daftar = tk.Listbox(kiri, height=7, width=28, exportselection=False)
        self.daftar.pack()
        ttk.Button(kiri, text="Segarkan daftar", command=self.segarkan_daftar).pack(fill="x", pady=(4, 0))

        tengah = ttk.Frame(t, padding=(14, 0))
        tengah.pack(side="left", fill="y")
        for teks, perintah in [
            ("Buka naskah.txt", self.buka_naskah),
            ("Cek naskah", lambda: self.untuk_terpilih("cek")),
            ("Rakit Word + PDF (butuh Word)", lambda: self.untuk_terpilih("rakit")),
            ("Buka folder bahts", self.buka_folder),
        ]:
            ttk.Button(tengah, text=teks, command=perintah).pack(fill="x", pady=2)

        kanan = ttk.LabelFrame(t, text="Bahts baru", padding=8)
        kanan.pack(side="left", fill="both", expand=True)
        ttk.Label(kanan, text="Nama folder (huruf Latin, tanpa spasi):").pack(anchor="w")
        self.e_folder = ttk.Entry(kanan, width=30)
        self.e_folder.pack(anchor="w")
        ttk.Label(kanan, text="Judul (huruf Arab):").pack(anchor="w", pady=(6, 0))
        self.e_judul = ttk.Entry(kanan, width=60, justify="right")
        self.e_judul.pack(anchor="w", fill="x")
        ttk.Button(kanan, text="Buat bahts baru", command=self.buat_baru).pack(anchor="w", pady=8)

    def _tab_kitab(self):
        t = ttk.Frame(self.buku, padding=10)
        self.buku.add(t, text="Cari kitab")
        baris1 = ttk.Frame(t)
        baris1.pack(fill="x")
        ttk.Label(baris1, text="Frasa / judul:").pack(side="left")
        self.e_frasa = ttk.Entry(baris1, width=40, justify="right")
        self.e_frasa.pack(side="left", padx=6)
        ttk.Button(baris1, text="Temukan di semua kitab", command=self.temukan).pack(side="left", padx=3)
        ttk.Button(baris1, text="Katalog (judul/pengarang)", command=self.katalog).pack(side="left", padx=3)
        ttk.Button(baris1, text="Kitab terunduh", command=lambda: self.jalankan(["kitab"])).pack(side="left", padx=3)

        baris2 = ttk.Frame(t)
        baris2.pack(fill="x", pady=8)
        ttk.Label(baris2, text="Nomor kitab:").pack(side="left")
        self.e_id = ttk.Entry(baris2, width=9)
        self.e_id.pack(side="left", padx=6)
        ttk.Button(baris2, text="Ambil (unduh)", command=lambda: self.dengan_id(["ambil"])).pack(side="left", padx=3)
        ttk.Button(baris2, text="Kartu (data cetak)", command=lambda: self.dengan_id(["kartu"])).pack(side="left", padx=3)
        ttk.Button(baris2, text="Cari frasa di kitab ini", command=self.cari_id).pack(side="left", padx=3)
        ttk.Button(baris2, text="Teks sekitar frasa", command=self.teks_id).pack(side="left", padx=3)

        baris3 = ttk.Frame(t)
        baris3.pack(fill="x")
        ttk.Label(baris3, text="Juz (- bila satu jilid):").pack(side="left")
        self.e_juz = ttk.Entry(baris3, width=5)
        self.e_juz.insert(0, "-")
        self.e_juz.pack(side="left", padx=6)
        ttk.Label(baris3, text="Halaman:").pack(side="left")
        self.e_hal = ttk.Entry(baris3, width=7)
        self.e_hal.pack(side="left", padx=6)
        ttk.Button(baris3, text="Tampilkan halaman", command=self.halaman).pack(side="left", padx=3)

    def _tab_sampul(self):
        t = ttk.Frame(self.buku, padding=10)
        self.buku.add(t, text="Data sampul")
        ttk.Label(t, wraplength=880, justify="left", text=(
            "Disimpan di sampul_saya.txt dan dipakai tiap bahts baru. Nama dan jenis ditulis huruf Arab. "
            "Pembimbing diisi di naskah.txt tiap bahts.")).pack(anchor="w")
        form = ttk.Frame(t)
        form.pack(anchor="w", pady=8)
        self.sampul = {}
        for baris, (kunci, label) in enumerate([
            ("nama", "Nama (Arab)"), ("nim", "Nomor mahasiswa"),
            ("jenis", "Jenis tugas (Arab)"), ("fasl", "Semester (Arab)")]):
            ttk.Label(form, text=label + ":").grid(row=baris, column=0, sticky="w", pady=2)
            e = ttk.Entry(form, width=44, justify="right" if kunci != "nim" else "left")
            e.grid(row=baris, column=1, padx=6)
            self.sampul[kunci] = e
        self.v_thalibah = tk.BooleanVar()
        ttk.Checkbutton(form, text="Mahasiswi (sampul mencetak «اسم الطالبة»)",
                        variable=self.v_thalibah).grid(row=4, column=1, sticky="w", pady=4)
        ttk.Button(form, text="Simpan", command=self.simpan_sampul).grid(row=5, column=1, sticky="w")
        self.muat_sampul()

    def _tab_panduan(self):
        t = ttk.Frame(self.buku, padding=6)
        self.buku.add(t, text="Panduan & konsep")
        isi = ttk.Frame(t)
        isi.pack(fill="both", expand=True)
        teks = tk.Text(isi, wrap="word", font=("Segoe UI", 10), padx=10, pady=8, height=12)
        gulung = ttk.Scrollbar(isi, command=teks.yview)
        teks.configure(yscrollcommand=gulung.set)
        gulung.pack(side="right", fill="y")
        teks.pack(side="left", fill="both", expand=True)
        teks.tag_configure("h1", font=("Segoe UI", 13, "bold"), spacing1=14, spacing3=4, foreground="#1f4e79")
        teks.tag_configure("h2", font=("Segoe UI", 10, "bold"), spacing1=8, spacing3=1)
        teks.tag_configure("p", lmargin1=0, lmargin2=0, spacing3=2)
        for jenis, isi_teks in PANDUAN:
            teks.insert("end", isi_teks + "\n", jenis)
        teks.configure(state="disabled")

    # ---------- menjalankan perintah ----------
    def tulis(self, teks, tebal=False):
        self.keluaran.configure(state="normal")
        self.keluaran.insert("end", teks)
        self.keluaran.see("end")
        self.keluaran.configure(state="disabled")

    def bersihkan(self):
        self.keluaran.configure(state="normal")
        self.keluaran.delete("1.0", "end")
        self.keluaran.configure(state="disabled")

    def jalankan(self, argumen, lewat_mesin=True):
        if self.sibuk:
            messagebox.showinfo("Mesin Bahts", "Masih ada perintah yang berjalan. Tunggu selesai atau tekan Hentikan.")
            return
        perintah = [sys.executable, "-X", "utf8", str(BAHTS)] + argumen if lewat_mesin else argumen
        tampil = " ".join(argumen) if lewat_mesin else " ".join(argumen[2:])
        self.tulis(f"\n> {tampil}\n")
        self.sibuk = True
        self.status.configure(text="berjalan…")
        threading.Thread(target=self._kerja, args=(perintah,), daemon=True).start()

    def _kerja(self, perintah):
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        bendera = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        try:
            self.proses = subprocess.Popen(
                perintah, cwd=ROOT, env=env, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                stdin=subprocess.DEVNULL, text=True, encoding="utf-8", errors="replace", creationflags=bendera)
            for baris in self.proses.stdout:
                self.antre.put(baris)
            kode = self.proses.wait()
            self.antre.put(f"[selesai, kode {kode}]\n")
        except Exception as galat:
            self.antre.put(f"[gagal menjalankan: {galat}]\n")
        self.antre.put(None)

    def _kuras(self):
        try:
            while True:
                x = self.antre.get_nowait()
                if x is None:
                    self.sibuk = False
                    self.proses = None
                    self.status.configure(text="siap")
                else:
                    self.tulis(x)
        except queue.Empty:
            pass
        self.after(100, self._kuras)

    def hentikan(self):
        if self.proses and self.proses.poll() is None:
            self.proses.terminate()
            self.tulis("[dihentikan]\n")

    # ---------- aksi ----------
    def pasang_pustaka(self):
        self.jalankan([sys.executable, "-m", "pip", "install", "-r", str(ROOT / "requirements.txt")], lewat_mesin=False)

    def segarkan_daftar(self):
        self.daftar.delete(0, "end")
        for p in sorted(ROOT.iterdir()):
            if p.is_dir() and (p / "naskah.txt").exists():
                self.daftar.insert("end", p.name)

    def terpilih(self):
        pilihan = self.daftar.curselection()
        if not pilihan:
            messagebox.showinfo("Mesin Bahts", "Pilih dulu satu bahts di daftar.")
            return None
        return self.daftar.get(pilihan[0])

    def untuk_terpilih(self, perintah):
        nama = self.terpilih()
        if nama:
            self.jalankan([perintah, nama])

    def buka_naskah(self):
        nama = self.terpilih()
        if nama and os.name == "nt":
            os.startfile(ROOT / nama / "naskah.txt")
        elif nama:
            self.tulis(f"Buka sendiri: {ROOT / nama / 'naskah.txt'}\n")

    def buka_folder(self):
        nama = self.terpilih()
        if nama and os.name == "nt":
            os.startfile(ROOT / nama)
        elif nama:
            self.tulis(f"Folder: {ROOT / nama}\n")

    def buat_baru(self):
        folder = self.e_folder.get().strip()
        judul = self.e_judul.get().strip()
        if not NAMA_FOLDER.match(folder):
            messagebox.showwarning("Mesin Bahts", "Nama folder harus huruf Latin/angka/_/- tanpa spasi.")
            return
        if not judul:
            messagebox.showwarning("Mesin Bahts", "Isi judul bahts.")
            return
        self.jalankan(["baru", folder, judul])
        self.after(1500, self.segarkan_daftar)

    def temukan(self):
        frasa = self.e_frasa.get().strip()
        if frasa:
            self.jalankan(["temukan", frasa])

    def katalog(self):
        frasa = self.e_frasa.get().strip()
        if frasa:
            self.jalankan(["katalog", frasa])

    def id_kitab(self):
        nomor = self.e_id.get().strip()
        if not nomor.isdigit():
            messagebox.showwarning("Mesin Bahts", "Isi nomor kitab (angka).")
            return None
        return nomor

    def dengan_id(self, awal):
        nomor = self.id_kitab()
        if nomor:
            self.jalankan(awal + [nomor])

    def cari_id(self):
        nomor, frasa = self.id_kitab(), self.e_frasa.get().strip()
        if nomor and frasa:
            self.jalankan(["cari", nomor, frasa])

    def teks_id(self):
        nomor, frasa = self.id_kitab(), self.e_frasa.get().strip()
        if nomor and frasa:
            self.jalankan(["teks", nomor, frasa])

    def halaman(self):
        nomor, juz, hal = self.id_kitab(), self.e_juz.get().strip() or "-", self.e_hal.get().strip()
        if nomor and hal.isdigit():
            self.jalankan(["halaman", nomor, juz, hal])
        elif nomor:
            messagebox.showwarning("Mesin Bahts", "Isi nomor halaman (angka).")

    def muat_sampul(self):
        nilai = {}
        if SAMPUL.exists():
            for baris in SAMPUL.read_text(encoding="utf-8-sig").splitlines():
                m = re.match(r"^@(\w+):\s*(.*)$", baris)
                if m:
                    nilai[m.group(1)] = m.group(2).strip()
        for kunci, e in self.sampul.items():
            e.delete(0, "end")
            e.insert(0, nilai.get(kunci, ""))
        self.v_thalibah.set(nilai.get("thalibah") == "ya")

    def simpan_sampul(self):
        baris = [f"@{k}: {e.get().strip()}" for k, e in self.sampul.items() if e.get().strip()]
        if self.v_thalibah.get():
            baris.append("@thalibah: ya")
        SAMPUL.write_text("\n".join(baris) + "\n", encoding="utf-8")
        self.tulis(f"Tersimpan: {SAMPUL}\n")


if __name__ == "__main__":
    app = App()
    if os.environ.get("BAHTS_GUI_TES"):
        app.after(1500, app.destroy)
    app.mainloop()
