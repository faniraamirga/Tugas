"""Susun Tugas-3 (Word) dari teks + diagram, siap diunggah ke SIAKAD."""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HARGA_AKHIR = 65000

def rp(n):
    return "Rp" + f"{n:,.0f}".replace(",", ".")

def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def compact(cell):
    for p in cell.paragraphs:
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(0)

doc = Document()
for s in doc.sections:
    s.top_margin = s.bottom_margin = Cm(2.5)
    s.left_margin = s.right_margin = Cm(2.5)
st = doc.styles["Normal"]
st.font.name = "Times New Roman"; st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
st.paragraph_format.line_spacing = 1.5

def para(text="", bold=False, italic=False, align="justify", size=None, space_after=6):
    p = doc.add_paragraph()
    r = p.add_run(text); r.bold = bold; r.italic = italic
    if size: r.font.size = Pt(size)
    if align == "center": p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "justify": p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(space_after)
    return p

def heading(text):
    p = doc.add_paragraph()
    r = p.add_run(text); r.bold = True; r.font.size = Pt(13)
    p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    return p

def cell_text(cell, text, bold=False, italic=False, size=10.5, center=False):
    cell.text = ""
    r = cell.paragraphs[0].add_run(text); r.bold = bold; r.italic = italic; r.font.size = Pt(size)
    if center: cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

# ---------- Kepala ----------
para("Tugas 3", bold=True, align="center", size=14, space_after=0)
para("Diagram rantai nilai satu produk kreatif lokal", bold=True, align="center", size=13, space_after=0)
para("Kopi arabika Latimojong kemasan 200 gram yang dijual di Kota Palopo",
     align="center", size=12, space_after=12)

identitas = [("Nama", "Fanira Amirga"), ("NIM", "……………………"),
             ("Mata kuliah", "Ekonomi Kreatif (MBS12523), pertemuan 3: Ekosistem dan Rantai Nilai"),
             ("Program studi", "Manajemen Bisnis Syariah, FEBI UIN Palopo, semester V"),
             ("Dosen", "Qisty Amalina Rusmana Putri, M.E.")]
t = doc.add_table(rows=len(identitas), cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
for row, (k, v) in zip(t.rows, identitas):
    row.cells[0].text = k; row.cells[1].text = ": " + v
    row.cells[0].width = Cm(3.5); row.cells[1].width = Cm(12.5)
    compact(row.cells[0]); compact(row.cells[1])
para("", space_after=4)

# ---------- 1. Produk ----------
heading("1. Produk yang dipilih dan alasannya")
para("Produk yang saya petakan adalah kopi arabika Latimojong dalam kemasan biji atau bubuk 200 gram, "
     "yang dijual di toko oleh-oleh dan kedai kopi di Kota Palopo. Kopinya ditanam petani di Kecamatan "
     "Latimojong dan Bastem, Kabupaten Luwu, pada ketinggian 1.500 meter ke atas. Roastery lokal seperti "
     "Kopi Wanua dan Kopi Sangtra di Belopa atau Solaku Roastery di Suli menyangrai dan mengemasnya, lalu "
     "kopi itu dikirim ke Palopo. Pada 31 Maret 2026 Pemkab Luwu menyerahkan sertifikat Indikasi Geografis "
     "\"Kopi Arabika Latimojong\" kepada lembaga pelindungnya.")
para("Produk ini masuk subsektor kuliner. Biji kopi yang dijual curah hanya dihargai sebagai bahan mentah. "
     "Harganya baru naik dua sampai tiga kali lipat setelah roastery memilih profil sangrai (medium, full "
     "wash, natural), memberi nama dan merek, membuat kemasan, dan menuliskan asal kebunnya di label. Di "
     "situlah letak nilai kreatifnya. Saya memilih kopi karena rantainya melewati semua aktor ekosistem yang "
     "dibahas di kelas: petani dan roaster sebagai pencipta, koperasi dan bantuan alat dari Disperindag "
     "Sulsel sebagai penyedia modal, kedai dan toko oleh-oleh sebagai pasar, Pemkab Luwu yang mengurus "
     "sertifikat IG dan bantuan bibit, serta perguruan tinggi yang meneliti dan mendampingi petani.")

# ---------- 2. Diagram ----------
heading("2. Diagram rantai nilai dari bahan mentah sampai ke tangan pembeli")
para("Satuan hitungnya adalah bahan untuk satu bungkus 200 gram kopi sangrai. Biji kopi menyusut sekitar "
     "15 sampai 17 persen saat disangrai, jadi satu bungkus 200 gram membutuhkan kira-kira 240 gram biji "
     "kering (green bean). Semua harga di bawah dihitung untuk porsi 240 gram itu.")
doc.add_picture("tugas-3/rantai_nilai_kopi_latimojong.png", width=Cm(16))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.paragraphs[-1].paragraph_format.keep_with_next = True
para("Gambar 1. Rantai nilai satu bungkus kopi arabika Latimojong 200 gram (angka perkiraan)",
     italic=True, align="center", size=10, space_after=10)

para("Alur rantainya sebagai berikut.", space_after=2)
tahapan = [
    "Kreasi oleh petani kopi di Latimojong dan Bastem. Petani menanam, memetik buah merah, mengupas, "
    "memfermentasi, dan menjemur biji sampai kering. Bahan dari luar hampir tidak ada karena bibit diambil "
    "dari kebun sendiri dan pupuknya dari kulit buah kopi, sehingga seluruh harga jual dicatat sebagai "
    "nilai tambah. Petani menjual biji kering sekitar Rp80.000 per kg ke pengepul.",
    "Distribusi pertama oleh pengepul kecamatan. Pengepul mengumpulkan biji dari banyak petani, menyortir, "
    "dan mengangkutnya dari pegunungan ke Belopa atau Palopo. Ia menjual sekitar Rp100.000 per kg ke roastery.",
    "Produksi oleh roastery. Roastery menyangrai, mengistirahatkan biji beberapa hari, menggiling bila "
    "diminta, mengemas dalam kantong aluminium foil bermerek, dan mengurus izin PIRT serta sertifikat halal. "
    "Satu bungkus 200 gram dijual ke toko atau kedai sekitar Rp45.000.",
    "Distribusi kedua oleh toko oleh-oleh dan kedai kopi di Palopo. Toko memajang, menyimpan stok, "
    "menanggung barang yang tidak laku, dan menjual ke pembeli sekitar Rp65.000 per bungkus. Kedai kopi "
    "bisa menjual kopi yang sama sebagai 12 sampai 15 cangkir seharga Rp16.500 sampai Rp21.500 per cangkir.",
    "Konsumsi oleh pembeli, yang membawa pulang kopi sebagai oleh-oleh khas Luwu atau meminumnya di kedai.",
    "Konservasi. Kulit buah kopi kembali ke kebun sebagai pupuk dan ampas seduhan bisa menjadi kompos. "
    "Kemasan aluminium foil belum didaur ulang dan tertinggal sebagai sampah.",
]
for isi in tahapan:
    p = doc.add_paragraph(isi, style="List Number")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(3)

# ---------- 3. Tabel ----------
heading("3. Perkiraan harga dan nilai tambah tiap pelaku")
para("Rumusnya sama dengan yang dipakai di kelas: nilai tambah = harga jual dikurangi bahan yang dibeli. "
     "Kemasan Rp3.000 saya hitung sebagai bahan yang dibeli roastery supaya nilai tambah roastery tidak "
     "menggelembung.")

baris = [
    ("Petani kopi Latimojong/Bastem", "tidak ada", "Rp19.200\n(240 g × Rp80.000/kg)", 0, 19200),
    ("Pengepul kecamatan", "Rp19.200", "Rp24.000\n(240 g × Rp100.000/kg)", 19200, 24000),
    ("Roastery", "Rp27.000\n(biji Rp24.000 + kemasan Rp3.000)", "Rp45.000", 27000, 45000),
    ("Toko oleh-oleh / kedai kopi Palopo", "Rp45.000", "Rp65.000", 45000, 65000),
]
tbl = doc.add_table(rows=1, cols=5); tbl.style = "Table Grid"; tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl.autofit = False
LEBAR = [Cm(3.9), Cm(3.5), Cm(3.5), Cm(2.6), Cm(2.5)]
for c, h in zip(tbl.rows[0].cells, ["Pelaku", "Bahan dibeli", "Dijual seharga", "Nilai tambah", "Bagian dari harga"]):
    cell_text(c, h, bold=True, center=True); shade(c, "D9D9D9")
total_nt = 0
for pelaku, beli_txt, jual_txt, beli, jual in baris:
    nt = jual - beli; total_nt += nt
    cells = tbl.add_row().cells
    for i, v in enumerate([pelaku, beli_txt, jual_txt, rp(nt), f"{nt / HARGA_AKHIR * 100:.0f}%"]):
        cell_text(cells[i], v, bold=(i == 3), center=(i > 0))
cells = tbl.add_row().cells
for i, v in enumerate(["Pemasok kemasan (di luar rantai kreatif)", "tidak ada", "Rp3.000", "Rp3.000", "5%"]):
    cell_text(cells[i], v, italic=True, center=(i > 0))
cells = tbl.add_row().cells
for i, v in enumerate(["Harga di tangan pembeli", "", rp(HARGA_AKHIR), rp(total_nt + 3000), "100%"]):
    cell_text(cells[i], v, bold=True, center=(i > 0)); shade(cells[i], "F2F2F2")
for row in tbl.rows:
    for c, w in zip(row.cells, LEBAR):
        c.width = w; compact(c)
    row.cells[0].paragraphs[0].runs[0].font.size = Pt(10)
# LibreOffice membaca lebar dari tblGrid, bukan dari tiap sel
for gc, w in zip(tbl._tbl.tblGrid.findall(qn("w:gridCol")), LEBAR):
    gc.set(qn("w:w"), str(int(w.twips)))
assert total_nt + 3000 == HARGA_AKHIR
para("Tabel 1. Perkiraan nilai tambah per bungkus kopi 200 gram", italic=True, align="center", size=10, space_after=6)
para("Sumber angka: harga green bean Rp80.000/kg dan biji sangrai Rp100.000/kg dari Kopi Sangtra Belopa "
     "(Kompas.com, 2018); kisaran harga green bean arabika nasional 2025 Rp105.000 sampai Rp135.000/kg "
     "(Kerta Koffie, 2025); harga kopi di kedai Palopo Rp16.500 sampai Rp21.500 per cangkir (menu Cousin "
     "Coffee dan Kedai Kopi Kulo). Harga eceran kemasan 200 gram saya perkirakan dari kisaran harga kopi "
     "single origin Sulawesi di pasar daring. Semua angka dibulatkan dan masih perlu dikonfirmasi ke pelaku.",
     italic=True, size=10)

# ---------- 4. Tiga ukuran ----------
heading("4. Menilai rantai dengan tiga ukuran")
para("Dari sisi laba, setiap pelaku masih untung dan rantai ini bertahan. Permintaan bahkan melebihi "
     "pasokan: Kopi Wanua mengaku kewalahan memenuhi pesanan, dan Bupati Luwu pernah diminta menunggu "
     "sampai bulan Mei karena stok kopi kosong.")
para("Dari sisi manusia, petani memegang 30 persen harga akhir. Angka ini lebih baik daripada perajin bakul "
     "rotan di kelas yang hanya memegang 23 persen, tetapi petani menanggung kerja paling lama. Ia merawat "
     "pohon setahun untuk satu musim panen tiga bulan, memetik di lereng 1.500 meter, dan menjemur di daerah "
     "yang sering hujan.")
para("Dari sisi lingkungan, kopi di Latimojong ditanam di bawah naungan cengkeh dan pisang, sehingga tutupan "
     "hutan di hulu sungai tetap terjaga, dan kulit buah kembali ke kebun. Sisa yang belum tertangani adalah "
     "kemasan foil sekali pakai.")

# ---------- 5. Paragraf penutup ----------
heading("5. Siapa yang bagiannya paling kecil, dan apakah itu sepadan?")
para("Kalau dilihat dari angka, bagian paling kecil diambil pengepul, Rp4.800 dari Rp65.000. Tetapi pengepul "
     "memegang biji itu hanya beberapa hari dan memutar modalnya berkali-kali dalam satu musim, jadi bagiannya "
     "sepadan dengan risiko dan tenaga yang ia keluarkan. Yang menurut saya tidak sepadan adalah bagian petani. "
     "Petani tercatat menerima Rp19.200, tertinggi ketiga di rantai, tetapi angka itu adalah upah untuk "
     "pekerjaan paling panjang: merawat pohon selama setahun, memetik satu per satu buah yang merah, mengupas, "
     "memfermentasi, dan menjemur, dengan risiko gagal panen karena cuaca yang ditanggung sendiri. Roastery dan "
     "toko menyentuh kopi itu hanya beberapa jam sampai beberapa minggu, dan bersama-sama mengambil Rp38.000 "
     "atau 58 persen dari harga akhir. Polanya sama dengan bakul rotan di kelas. Semakin dekat ke konsumen, "
     "semakin besar nilai yang diambil, karena di sanalah informasi pasar berada. Roastery tahu profil sangrai "
     "apa yang sedang disukai, toko tahu siapa pembelinya dan berapa mereka sanggup bayar, sedangkan petani di "
     "Latimojong tidak tahu keduanya. Inilah ketegangan antara value creation dan value capture yang dibahas "
     "Dai dkk. (2024). Rasa khas kopi Latimojong diciptakan di kebun, tetapi nilainya paling banyak ditangkap "
     "di Belopa dan Palopo. Sertifikat Indikasi Geografis dan cara Kopi Wanua menjual langsung dari kebun ke "
     "warkop Palopo tanpa pengepul menunjukkan bahwa jalur untuk menaikkan bagian petani sudah ada: memangkas "
     "perantara, membangun merek sendiri, dan menjual langsung ke konsumen.")

# ---------- Referensi ----------
heading("Referensi")
refs = [
    "Amir, A. (2018, 28 September). Sangtra, kopi asli Pegunungan Latimojong. Kompas.com. "
    "https://travel.kompas.com/read/2018/09/28/201000627/sangtra-kopi-asli-pegunungan-latimojong",
    "Dai, G., Zhang, L., Zhang, Q., & Mao, M. (2024). Navigating tensions between value creation and value "
    "capture in ecosystems. Journal of Business Research, 170, 114333. https://doi.org/10.1016/j.jbusres.2023.114333",
    "Hasnida, Nuraeni, & Hasan, I. (2021). Analisis sistem agribisnis kopi arabika di Desa Tolajuk, Kecamatan "
    "Latimojong, Kabupaten Luwu. Wiratani: Jurnal Ilmiah Agribisnis, 4(1). "
    "https://jurnal.agribisnis.umi.ac.id/index.php/wiratani/article/view/132",
    "Kerta Koffie. (2025, 12 September). Harga green bean arabika terbaru 2025: Panduan lengkap untuk roastery "
    "dan cafe. https://kertakoffie.com/harga-green-bean-arabika-terbaru-2025-panduan-lengkap-untuk-roastery-dan-cafe/",
    "Putri, Q. A. R. (2026). Pertemuan 3: Ekosistem dan rantai nilai [Slide perkuliahan Ekonomi Kreatif MBS12523]. "
    "Manajemen Bisnis Syariah, FEBI UIN Palopo.",
    "Redaksi Luwu. (2025, 25 November). Bupati dorong peningkatan produksi dan kualitas Kopi Arabika Latimojong: "
    "Setiap produk kopi harus melekatkan nama Luwu. Tekape.co. "
    "https://tekape.co/bupati-dorong-peningkatan-produksi-dan-kualitas-kopi-arabika-latimojong-setiap-produk-kopi-harus-melekatkan-nama-luwu/",
    "Redaksi Palopopos. (2022, 19 Oktober). Masmindo perkenalkan Kopi Latimojong ke pasar Eropa. Palopo Pos. "
    "https://palopopos.fajar.co.id/2022/10/19/masmindo-perkenalkan-kopi-latimojong-ke-pasar-eropa/",
    "SINDOSulsel. (2025, 24 November). Kopi Wanua Belopa, aroma Latimojong yang mulai menembus pasar nasional. "
    "https://sindosulsel.com/kopi-wanua-belopa-aroma-latimojong-yang-mulai-menembus-pasar-nasional/",
    "SINDOSulsel. (2025, 25 November). Kopi Solaku, kebanggaan baru dari Arabika Latimojong. "
    "https://sindosulsel.com/kopi-solaku-kebanggaan-baru-dari-arabika-latimojong/",
    "SulselNow. (2026, 1 April). Kopi Arabika Latimojong resmi kantongi sertifikat Indikasi Geografis, siap "
    "bersaing di pasar nasional. https://sulselnow.com/11835/kopi-arabika-latimojong-resmi-kantongi-sertifikat-"
    "indikasi-geografis-siap-bersaing-di-pasar-nasional/",
]
for r_ in refs:
    p = para(r_, size=11, space_after=4)
    p.paragraph_format.left_indent = Cm(1); p.paragraph_format.first_line_indent = Cm(-1)
    p.paragraph_format.line_spacing = 1.15

doc.save("tugas-3/Tugas-3_Rantai_Nilai_Kopi_Latimojong.docx")
print("saved")
