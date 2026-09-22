"""Susun Tugas-3 (Word) dari teks + diagram, siap diunggah ke SIAKAD."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def rp(n):
    return "Rp" + f"{n:,.0f}".replace(",", ".")

def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

doc = Document()
for s in doc.sections:
    s.top_margin = s.bottom_margin = Cm(2.5)
    s.left_margin = s.right_margin = Cm(2.5)
st = doc.styles["Normal"]
st.font.name = "Times New Roman"; st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
st.paragraph_format.line_spacing = 1.5

def para(text="", bold=False, italic=False, align=None, size=None, space_after=6):
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
    return p

# ---------- Kepala ----------
para("TUGAS-3", bold=True, align="center", size=14, space_after=0)
para("Diagram Rantai Nilai Satu Produk Kreatif Lokal", bold=True, align="center", size=13, space_after=0)
para("Kopi Arabika Latimojong Kemasan 200 g yang Dijual di Kota Palopo", align="center", size=12, space_after=12)

t = doc.add_table(rows=5, cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
identitas = [("Nama", "Fanira Amirga"), ("NIM", "……………………"),
             ("Mata Kuliah", "Ekonomi Kreatif (MBS12523) · Pertemuan 3: Ekosistem dan Rantai Nilai"),
             ("Program Studi", "Manajemen Bisnis Syariah, FEBI UIN Palopo · Semester V"),
             ("Dosen", "Qisty Amalina Rusmana Putri, M.E.")]
for row, (k, v) in zip(t.rows, identitas):
    row.cells[0].text = k; row.cells[1].text = ": " + v
    row.cells[0].width = Cm(3.5); row.cells[1].width = Cm(12.5)
    for c in row.cells:
        for p in c.paragraphs:
            p.paragraph_format.line_spacing = 1.15; p.paragraph_format.space_after = Pt(0)
para("", space_after=4)

# ---------- 1. Produk ----------
heading("1. Produk yang Dipilih dan Alasannya")
para("Produk yang dipetakan adalah kopi arabika Latimojong dalam kemasan biji atau bubuk 200 gram, "
     "yang dijual di toko oleh-oleh dan kedai kopi di Kota Palopo. Kopinya ditanam petani di Kecamatan "
     "Latimojong dan Bastem, Kabupaten Luwu, pada ketinggian 1.500 meter ke atas, lalu diolah oleh "
     "roastery lokal seperti Kopi Wanua dan Kopi Sangtra di Belopa atau Solaku Roastery di Suli, dan "
     "dipasarkan ke Palopo. Pada Maret 2026 kopi ini resmi menerima sertifikat Indikasi Geografis (IG) "
     "“Kopi Arabika Latimojong”.", align="justify")
para("Produk ini masuk subsektor kuliner ekonomi kreatif, tetapi nilai kreatifnya tidak terletak pada "
     "biji kopinya. Biji yang sama, jika dijual sebagai komoditas curah, hanya dihargai sebagai bahan "
     "mentah. Nilai kreatif muncul ketika roastery menentukan profil sangrai (medium, full wash, natural), "
     "memberi nama dan merek, membuat kemasan, dan menceritakan asal kebunnya. Karena itu satu bungkus "
     "kopi bermerek dapat dijual dua sampai tiga kali lipat harga bijinya. Rantai kopi juga melibatkan "
     "seluruh aktor ekosistem yang dibahas di kelas: pencipta (petani dan roaster), penyedia modal "
     "(koperasi dan bantuan alat dari Disperindag Sulsel), pasar (kedai dan toko oleh-oleh), pemerintah "
     "(sertifikat IG dan bantuan bibit dari Pemkab Luwu), serta perguruan tinggi yang meneliti dan "
     "mendampingi petani.", align="justify")

# ---------- 2. Diagram ----------
heading("2. Diagram Rantai Nilai dari Bahan Mentah sampai ke Tangan Pembeli")
para("Satuan perhitungan adalah bahan untuk satu bungkus 200 gram kopi sangrai. Karena biji kopi menyusut "
     "sekitar 15–17% saat disangrai, satu bungkus 200 gram membutuhkan kira-kira 240 gram biji kopi "
     "kering (green bean). Semua harga di bawah dihitung untuk porsi 240 gram tersebut.", align="justify")
doc.add_picture("tugas-3/rantai_nilai_kopi_latimojong.png", width=Cm(16))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
cap = para("Gambar 1. Rantai nilai satu bungkus kopi arabika Latimojong 200 g (angka perkiraan)",
           italic=True, align="center", size=10, space_after=10)

para("Alur rantainya sebagai berikut:", space_after=2)
tahapan = [
    ("Kreasi — petani kopi di Latimojong/Bastem.", " Petani menanam, memetik buah merah, mengupas, "
     "memfermentasi, dan menjemur biji sampai kering. Bahan yang dibeli dari luar hampir tidak ada "
     "(bibit dari kebun sendiri, pupuk dari kulit buah kopi), sehingga seluruh harga jualnya dicatat "
     "sebagai nilai tambah. Petani menjual biji kering sekitar Rp80.000 per kg ke pengepul."),
    ("Distribusi 1 — pengepul kecamatan.", " Pengepul mengumpulkan biji dari banyak petani, menyortir, "
     "dan mengangkut dari pegunungan ke Belopa atau Palopo. Ia menjual sekitar Rp100.000 per kg ke roastery."),
    ("Produksi — roastery/pengolah.", " Roastery menyangrai, mengendapkan (resting) biji, menggiling bila "
     "diminta, mengemas dalam kantong aluminium foil bermerek, dan mengurus izin PIRT serta sertifikat halal. "
     "Satu bungkus 200 g dijual ke toko atau kedai sekitar Rp45.000."),
    ("Distribusi 2 — toko oleh-oleh dan kedai kopi di Palopo.", " Toko memajang, menyimpan stok, "
     "menanggung barang yang tidak laku, dan menjual ke pembeli sekitar Rp65.000 per bungkus. Kedai kopi "
     "bahkan bisa menjual kopi yang sama sebagai 12–15 cangkir seharga Rp16.500–Rp21.500 per cangkir."),
    ("Konsumsi — pembeli.", " Pembeli membawa pulang kopi sebagai oleh-oleh khas Luwu atau meminumnya di kedai."),
    ("Konservasi.", " Kulit buah kopi dikembalikan ke kebun sebagai pupuk dan ampas seduhan bisa menjadi "
     "kompos. Kemasan aluminium foil belum didaur ulang dan menjadi sampah yang ditinggalkan rantai ini."),
]
for judul, isi in tahapan:
    p = doc.add_paragraph(style="List Number")
    r = p.add_run(judul); r.bold = True
    p.add_run(isi)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(3)

# ---------- 3. Tabel ----------
heading("3. Perkiraan Harga dan Nilai Tambah Tiap Pelaku")
para("Rumus yang dipakai sama dengan di kelas: nilai tambah = harga jual − bahan yang dibeli. "
     "Kemasan Rp3.000 dihitung sebagai bahan yang dibeli roastery, supaya nilai tambah roastery tidak "
     "menggelembung.", align="justify")

baris = [
    ("Petani kopi Latimojong/Bastem", "—", "Rp19.200\n(240 g × Rp80.000/kg)", 19200),
    ("Pengepul kecamatan", "Rp19.200", "Rp24.000\n(240 g × Rp100.000/kg)", 24000 - 19200),
    ("Roastery / pengolah", "Rp27.000\n(biji Rp24.000 + kemasan Rp3.000)", "Rp45.000", 45000 - 27000),
    ("Toko oleh-oleh / kedai kopi Palopo", "Rp45.000", "Rp65.000", 65000 - 45000),
]
tbl = doc.add_table(rows=1, cols=5); tbl.style = "Table Grid"; tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = ["Pelaku", "Bahan dibeli", "Dijual seharga", "Nilai tambah", "Bagian dari harga akhir"]
for c, h in zip(tbl.rows[0].cells, hdr):
    c.text = ""; r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(10.5)
    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER; shade(c, "D9D9D9")
total_nt = 0
for pelaku, beli, jual, nt in baris:
    total_nt += nt
    cells = tbl.add_row().cells
    vals = [pelaku, beli, jual, rp(nt), f"{nt/65000*100:.0f}%"]
    for i, (c, v) in enumerate(zip(cells, vals)):
        c.text = ""; r = c.paragraphs[0].add_run(v); r.font.size = Pt(10.5)
        if i == 3: r.bold = True
        if i > 0: c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
cells = tbl.add_row().cells
for i, v in enumerate(["Pemasok kemasan (di luar rantai kreatif)", "—", "Rp3.000", "Rp3.000", "5%"]):
    c = cells[i]; c.text = ""; r = c.paragraphs[0].add_run(v); r.font.size = Pt(10.5); r.italic = True
    if i > 0: c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
cells = tbl.add_row().cells
for i, v in enumerate(["Harga di tangan pembeli", "", "Rp65.000", rp(total_nt + 3000), "100%"]):
    c = cells[i]; c.text = ""; r = c.paragraphs[0].add_run(v); r.bold = True; r.font.size = Pt(10.5)
    if i > 0: c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    shade(c, "F2F2F2")
for row in tbl.rows:
    for c in row.cells:
        for p in c.paragraphs:
            p.paragraph_format.line_spacing = 1.15; p.paragraph_format.space_after = Pt(0)
para("", space_after=2)
para("Sumber angka: harga green bean dan biji siap giling Kopi Sangtra Belopa (Kompas.com, 2018: green bean "
     "Rp80.000/kg, biji sangrai Rp100.000/kg); kisaran harga green bean arabika nasional 2025 Rp105.000–"
     "Rp135.000/kg (Kerta Koffie, 2025); harga menu kedai kopi di Palopo Rp16.500–Rp21.500 per cangkir "
     "(Cousin Coffee, Kedai Kopi Kulo). Harga eceran kemasan 200 g adalah perkiraan dari kisaran harga "
     "kopi single origin Sulawesi. Semua angka dibulatkan dan perlu dikonfirmasi ke pelaku.",
     italic=True, size=10, align="justify")

# ---------- 4. Tiga ukuran ----------
heading("4. Menilai Rantai dengan Tiga Ukuran")
ukuran = [
    ("Laba.", " Setiap pelaku masih untung dan rantai ini bertahan; permintaan bahkan melebihi pasokan, "
     "Kopi Wanua menyebut kewalahan memenuhi pesanan dan Bupati Luwu pernah diminta menunggu sampai bulan "
     "lima karena stok kosong."),
    ("Manusia.", " Petani memegang 30% harga akhir. Angka ini lebih baik daripada perajin bakul rotan di kelas "
     "(23%), tetapi petani menanggung kerja paling lama: merawat pohon setahun untuk satu musim panen tiga "
     "bulan, memetik di lereng 1.500 meter, dan menjemur di daerah yang sering hujan."),
    ("Lingkungan.", " Kopi di Latimojong ditanam di bawah naungan cengkeh dan pisang, sehingga menjaga tutupan "
     "hutan di hulu sungai. Kulit buah kembali ke kebun. Sisa yang belum tertangani adalah kemasan foil "
     "sekali pakai."),
]
for judul, isi in ukuran:
    p = doc.add_paragraph(style="List Bullet")
    r = p.add_run(judul); r.bold = True; p.add_run(isi)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.space_after = Pt(3)

# ---------- 5. Paragraf penutup ----------
heading("5. Siapa yang Bagiannya Paling Kecil, dan Apakah Itu Sepadan?")
para("Kalau dilihat dari angka, bagian paling kecil diambil pengepul, Rp4.800 dari Rp65.000. Tetapi "
     "pengepul memegang biji itu hanya beberapa hari dan memutar modalnya berkali-kali dalam satu musim, "
     "sehingga bagiannya sebenarnya sepadan dengan risiko dan tenaga yang ia keluarkan. Yang menurut saya "
     "tidak sepadan adalah bagian petani. Petani memang tercatat menerima Rp19.200, tertinggi ketiga di "
     "rantai, tetapi angka itu adalah upah untuk pekerjaan paling panjang: merawat pohon selama setahun, "
     "memetik satu per satu buah yang merah, mengupas, memfermentasi, dan menjemur, dengan risiko gagal "
     "panen karena cuaca yang ditanggung sendiri. Sementara roastery dan toko, yang menyentuh kopi itu hanya "
     "beberapa jam sampai beberapa minggu, bersama-sama mengambil Rp38.000 atau 58% dari harga akhir. Pola "
     "ini sama persis dengan bakul rotan di kelas: semakin dekat ke konsumen, semakin besar nilai yang "
     "diambil, karena di sanalah informasi pasar berada. Roastery tahu profil sangrai apa yang sedang "
     "disukai, toko tahu siapa pembelinya dan berapa mereka sanggup bayar; petani di Latimojong tidak tahu "
     "keduanya. Inilah ketegangan antara value creation dan value capture yang dibahas Dai dkk. (2024): rasa "
     "khas kopi Latimojong diciptakan di kebun, tetapi nilainya paling banyak ditangkap di Belopa dan Palopo. "
     "Kabar baiknya, sertifikat Indikasi Geografis dan koperasi petani seperti Kopi Wanua, yang menjual "
     "langsung dari kebun ke warkop Palopo tanpa pengepul, membuktikan bahwa jalur untuk menaikkan bagian "
     "petani sudah ada: memangkas perantara, membangun merek sendiri, dan menjual langsung ke konsumen.",
     align="justify")

# ---------- Referensi ----------
heading("Referensi")
refs = [
    "Dai, Y., Zhang, W., Zhang, X., & Mao, J. (2024). Navigating tensions between value creation and value "
    "capture in ecosystems. Journal of Business Research, 170, 114333. https://doi.org/10.1016/j.jbusres.2023.114333",
    "Amir, A. (2018, 28 September). Sangtra, kopi asli Pegunungan Latimojong. Kompas.com. "
    "https://travel.kompas.com/read/2018/09/28/201000627/sangtra-kopi-asli-pegunungan-latimojong",
    "Kerta Koffie. (2025, 12 September). Harga green bean arabika terbaru 2025: Panduan lengkap untuk roastery dan cafe. "
    "https://kertakoffie.com/harga-green-bean-arabika-terbaru-2025-panduan-lengkap-untuk-roastery-dan-cafe/",
    "SINDOSulsel. (2025, 24 November). Kopi Wanua Belopa, aroma Latimojong yang mulai menembus pasar nasional. "
    "https://sindosulsel.com/kopi-wanua-belopa-aroma-latimojong-yang-mulai-menembus-pasar-nasional/",
    "SINDOSulsel. (2025, 25 November). Kopi Solaku, kebanggaan baru dari Arabika Latimojong. "
    "https://sindosulsel.com/kopi-solaku-kebanggaan-baru-dari-arabika-latimojong/",
    "SulselNow. (2026, 1 April). Kopi Arabika Latimojong resmi kantongi sertifikat Indikasi Geografis, siap bersaing "
    "di pasar nasional. https://sulselnow.com/11835/kopi-arabika-latimojong-resmi-kantongi-sertifikat-indikasi-geografis-siap-bersaing-di-pasar-nasional/",
    "Tekape.co. (2025, 25 November). Bupati dorong peningkatan produksi dan kualitas Kopi Arabika Latimojong: "
    "Setiap produk kopi harus melekatkan nama Luwu. https://tekape.co/",
    "Putri, Q. A. R. (2025). Pertemuan 3: Ekosistem dan rantai nilai [Slide perkuliahan]. Manajemen Bisnis Syariah, FEBI UIN Palopo.",
]
for r_ in refs:
    p = para(r_, size=11, space_after=4, align="justify")
    p.paragraph_format.left_indent = Cm(1); p.paragraph_format.first_line_indent = Cm(-1)
    p.paragraph_format.line_spacing = 1.15

doc.save("tugas-3/Tugas-3_Rantai_Nilai_Kopi_Latimojong.docx")
print("saved")
