"""Susun Tugas-3 (Word): rantai nilai dompo pisang berlin Dompis Berlin Thary, Palopo."""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HARGA_AKHIR = 20000
BAHAN_LAIN = 1700
KEMASAN = 1800
GAMBAR = "tugas-3/rantai_nilai_dompo_pisang_berlin.png"

def rp(n):
    return "Rp" + f"{n:,.0f}".replace(",", ".")

def pct(n, total=HARGA_AKHIR):
    v = n / total * 100
    return (f"{v:.1f}".replace(".", ",") if v != int(v) else f"{int(v)}") + "%"

def shade(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)

def compact(cell):
    for p in cell.paragraphs:
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(0)

def set_grid(tbl, widths):
    tbl.autofit = False
    for row in tbl.rows:
        for c, w in zip(row.cells, widths):
            c.width = w; compact(c)
    # LibreOffice membaca lebar dari tblGrid, bukan dari tiap sel
    for gc, w in zip(tbl._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(w.twips)))
    # Jaga tabel tetap satu halaman: baris tidak dipecah, tiap baris menempel ke baris berikutnya
    for i, row in enumerate(tbl.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        if i < len(tbl.rows) - 1:
            for c in row.cells:
                for p in c.paragraphs:
                    p.paragraph_format.keep_with_next = True

def restart_numbering(paragraph):
    """Mulai daftar bernomor baru dari 1 dengan membuat instance w:num baru."""
    numbering = doc.part.numbering_part.numbering_definitions._numbering
    # Gaya "List Number" menyimpan numId di definisi gaya, bukan di paragraf
    style_numPr = paragraph.style.element.pPr.numPr
    old_num_id = style_numPr.numId.val
    old_num = next(n for n in numbering.findall(qn("w:num")) if n.get(qn("w:numId")) == str(old_num_id))
    abstract_id = old_num.find(qn("w:abstractNumId")).get(qn("w:val"))
    new_id = max(int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))) + 1
    num = OxmlElement("w:num"); num.set(qn("w:numId"), str(new_id))
    abs_ref = OxmlElement("w:abstractNumId"); abs_ref.set(qn("w:val"), abstract_id); num.append(abs_ref)
    ovr = OxmlElement("w:lvlOverride"); ovr.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:startOverride"); start.set(qn("w:val"), "1"); ovr.append(start); num.append(ovr)
    numbering.append(num)
    numPr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
    numPr.get_or_add_ilvl().val = 0
    numPr.get_or_add_numId().val = new_id
    return new_id

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

def numbered(text, num_id=None):
    p = doc.add_paragraph(text, style="List Number")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(3)
    if num_id is not None:
        numPr = p._p.get_or_add_pPr().get_or_add_numPr()
        numPr.get_or_add_ilvl().val = 0
        numPr.get_or_add_numId().val = num_id
    return p

# ---------- Kepala ----------
para("Tugas 3", bold=True, align="center", size=14, space_after=0)
para("Diagram rantai nilai satu produk kreatif lokal", bold=True, align="center", size=13, space_after=0)
para("Dompo pisang berlin Dompis Berlin Thary, Kota Palopo", align="center", size=12, space_after=12)

identitas = [("Nama", "Fanira Amirga"), ("NIM", "……………………"),
             ("Mata kuliah", "Ekonomi Kreatif (MBS12523), pertemuan 3: Ekosistem dan Rantai Nilai"),
             ("Program studi", "Manajemen Bisnis Syariah, FEBI UIN Palopo, semester V"),
             ("Dosen", "Qisty Amalina Rusmana Putri, M.E.")]
t = doc.add_table(rows=len(identitas), cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
for row, (k, v) in zip(t.rows, identitas):
    row.cells[0].text = k; row.cells[1].text = ": " + v
set_grid(t, [Cm(3.5), Cm(12.5)])
para("", space_after=4)

# ---------- 1. Produk ----------
heading("1. Produk yang dipilih dan alasannya")
para("Produk yang saya petakan adalah dompo pisang berlin merek Dompis Berlin Thary, camilan pisang yang "
     "dijemur sampai kering lalu digoreng, dalam kemasan 200 gram. Usaha ini dijalankan Eka Indah Lestari "
     "(Thary) di Jalan Dr. Ratulangi, Kelurahan Salobulo, Kecamatan Wara Utara, dan menjadi UMKM binaan "
     "Dinas Koperasi dan UMKM Kota Palopo. Thary memulainya pada 2016 untuk menambah biaya kuliah, setelah "
     "melihat banyak pisang berlin di lapak neneknya di pasar terbuang karena terlalu matang dan tidak laku "
     "(Tribun Timur, 2024).")
para("Dompo pisang dibuat banyak orang di Palopo, dan itu justru alasan saya memilih merek ini. Hampir semua "
     "pembuat memakai pisang kepok atau raja; Thary memakai pisang berlin yang kecil, manis, dan sedikit asam, "
     "sehingga rasa dompo-nya berbeda dari yang lain. Slide pertemuan ini menyebut galon merah muda sebagai "
     "contoh bahwa pembeda kecil pun menciptakan nilai, dan pisang berlin adalah pembeda kecil semacam itu. "
     "Nilai kreatifnya ada pada pilihan bahan, resep, nama Dompis yang mudah diingat, dan kemasan yang "
     "membuat camilan pasar menjadi oleh-oleh. Rantainya juga pendek dan berada di dalam kota, jadi semua "
     "pelakunya dapat ditanyai langsung.")
para("Ekosistem yang menopang produk ini lengkap. Pencipta adalah petani pisang dan Thary sendiri. Modal awal "
     "datang dari keluarga, lalu Dinas Koperasi dan UMKM masuk sebagai pemerintah yang membina, mengurus "
     "izin PIRT, dan membawa produk ke pameran. Pasarnya adalah toko oleh-oleh, pembeli lewat WhatsApp dan "
     "Instagram, serta pesanan dari luar kota. Perguruan tinggi hadir lewat program pengabdian dan mahasiswa "
     "yang meneliti usaha semacam ini, termasuk tugas ini.")

# ---------- 2. Diagram ----------
heading("2. Diagram rantai nilai dari bahan mentah sampai ke tangan pembeli")
para("Satuan hitungnya adalah satu sisir pisang berlin, kira-kira 1 kilogram atau 12 sampai 15 buah, yang "
     "setelah dikupas, dijemur dua sampai tiga hari, dan digoreng menyusut menjadi sekitar 200 gram dompo. "
     "Penyusutan sekitar 70 persen ini sesuai dengan kadar air sale pisang kering 15 sampai 20 persen "
     "(Santoso dalam UMM, 2003). Jadi satu sisir pisang menjadi satu bungkus dompo, dan semua harga di bawah "
     "dihitung untuk satuan itu.")
doc.add_picture(GAMBAR, width=Cm(16))
doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
doc.paragraphs[-1].paragraph_format.keep_with_next = True
para("Gambar 1. Rantai nilai satu bungkus dompo pisang berlin 200 gram (angka perkiraan). Panah putus-putus "
     "menunjukkan jalur penjualan langsung tanpa toko.", italic=True, align="center", size=10, space_after=10)

para("Alur rantainya sebagai berikut.", space_after=2)
numbered("Kreasi oleh petani pisang di kebun campuran sekitar Luwu dan pinggiran Palopo. Pisang berlin ditanam "
         "di sela kakao atau kelapa, hampir tanpa biaya: anakan dari rumpun sendiri, pupuk dari daun dan kulit "
         "yang membusuk. Petani menjual satu tandan berisi empat sampai lima sisir sekitar Rp15.000 sampai "
         "Rp20.000 ke pedagang, atau Rp4.000 per sisir. Karena bahan yang dibeli hampir nol, seluruh harga jual "
         "dicatat sebagai nilai tambah petani.")
numbered("Distribusi pertama oleh pedagang pisang di pasar Palopo. Pedagang membeli per tandan, mengangkut ke "
         "kota, memeram, memajang, dan menanggung pisang yang tidak laku sampai lewat matang. Ia menjual "
         "Rp8.000 per sisir ke pembeli, termasuk ke Thary. Pisang yang sudah terlalu matang biasanya dilepas "
         "lebih murah, dan dari pisang seperti itulah Thary memulai usahanya.")
numbered("Produksi oleh Dompis Berlin Thary. Pisang dikupas, dibelah, dijemur dua sampai tiga hari di bawah "
         "matahari, dicelup adonan tepung tapioka, digoreng, ditiriskan, lalu dikemas dalam standing pouch "
         "berstiker merek. Selain pisang Rp8.000, Thary membeli tepung, minyak goreng, dan gas sekitar Rp1.700 "
         "serta kemasan dan label Rp1.800 per bungkus. Ia menjual ke toko oleh-oleh sekitar Rp15.000 per bungkus.")
numbered("Distribusi kedua oleh toko oleh-oleh di Palopo. Toko memajang, menyimpan stok, menanggung barang yang "
         "melewati masa simpan, dan menjual ke pembeli Rp20.000 per bungkus. Pada jalur langsung, Thary sendiri "
         "yang menjual Rp20.000 lewat WhatsApp, Instagram, dan rumah produksi, sehingga bagian toko masuk ke Thary.")
numbered("Konsumsi oleh pembeli, kebanyakan perantau dan pendatang yang membawa dompo sebagai oleh-oleh khas "
         "Palopo atau memesannya dari luar kota.")
numbered("Konservasi. Rantai ini dimulai dari pisang yang tadinya dibuang, jadi ia mengurangi sampah pasar. Kulit "
         "pisang menjadi pakan ternak atau kompos, dan penjemuran memakai matahari tanpa bahan bakar. Yang "
         "tertinggal sebagai sampah adalah minyak jelantah dan kemasan plastik sekali pakai.")

# ---------- 3. Tabel ----------
heading("3. Perkiraan harga dan nilai tambah tiap pelaku")
para("Rumusnya sama dengan yang dipakai di kelas: nilai tambah = harga jual dikurangi bahan yang dibeli. "
     "Tepung, minyak, gas, kemasan, dan label saya hitung sebagai bahan yang dibeli Thary supaya nilai "
     "tambahnya tidak menggelembung. Tabel 1 memakai jalur yang paling umum, yaitu lewat toko oleh-oleh.")

baris = [
    ("Petani pisang berlin", "tidak ada", "Rp4.000\n(1 sisir)", 0, 4000),
    ("Pedagang pisang di pasar", "Rp4.000", "Rp8.000\n(1 sisir)", 4000, 8000),
    ("Dompis Berlin Thary", "Rp11.500\n(pisang Rp8.000, tepung, minyak, gas Rp1.700, kemasan dan label Rp1.800)",
     "Rp15.000\n(1 bungkus 200 g)", 8000 + BAHAN_LAIN + KEMASAN, 15000),
    ("Toko oleh-oleh Palopo", "Rp15.000", "Rp20.000", 15000, HARGA_AKHIR),
]
tbl = doc.add_table(rows=1, cols=5); tbl.style = "Table Grid"; tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for c, h in zip(tbl.rows[0].cells, ["Pelaku", "Bahan dibeli", "Dijual seharga", "Nilai tambah", "Bagian dari harga"]):
    cell_text(c, h, bold=True, center=True); shade(c, "D9D9D9")
total_nt = 0
for pelaku, beli_txt, jual_txt, beli, jual in baris:
    nt = jual - beli; total_nt += nt
    cells = tbl.add_row().cells
    for i, v in enumerate([pelaku, beli_txt, jual_txt, rp(nt), pct(nt)]):
        cell_text(cells[i], v, bold=(i == 3), center=(i > 0), size=10)
lain = BAHAN_LAIN + KEMASAN
cells = tbl.add_row().cells
for i, v in enumerate(["Pemasok tepung, minyak, gas, kemasan (di luar rantai kreatif)", "tidak ada", rp(lain), rp(lain), pct(lain)]):
    cell_text(cells[i], v, italic=True, center=(i > 0), size=10)
cells = tbl.add_row().cells
for i, v in enumerate(["Harga di tangan pembeli", "", rp(HARGA_AKHIR), rp(total_nt + lain), "100%"]):
    cell_text(cells[i], v, bold=True, center=(i > 0), size=10); shade(cells[i], "F2F2F2")
set_grid(tbl, [Cm(3.9), Cm(4.0), Cm(3.0), Cm(2.6), Cm(2.5)])
assert total_nt + lain == HARGA_AKHIR
para("Tabel 1. Perkiraan nilai tambah per bungkus dompo 200 gram, jalur lewat toko oleh-oleh",
     italic=True, align="center", size=10, space_after=8)

para("Tabel 2 membandingkan dua jalur yang sama-sama dipakai Thary. Pada jalur langsung, toko hilang dari "
     "rantai dan Thary menjual sendiri dengan harga yang sama, Rp20.000.")
tbl2 = doc.add_table(rows=1, cols=3); tbl2.style = "Table Grid"; tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
for c, h in zip(tbl2.rows[0].cells, ["Pelaku", "Jalur lewat toko", "Jalur langsung (WA, Instagram, rumah produksi)"]):
    cell_text(c, h, bold=True, center=True); shade(c, "D9D9D9")
skenario = [
    ("Petani pisang berlin", 4000, 4000),
    ("Pedagang pisang di pasar", 4000, 4000),
    ("Dompis Berlin Thary", 3500, 8500),
    ("Toko oleh-oleh", 5000, 0),
    ("Pemasok bahan lain dan kemasan", lain, lain),
]
for nama, a, b in skenario:
    cells = tbl2.add_row().cells
    cell_text(cells[0], nama, size=10)
    cell_text(cells[1], f"{rp(a)} ({pct(a)})", center=True, size=10)
    cell_text(cells[2], (f"{rp(b)} ({pct(b)})" if b else "tidak ada"), center=True, size=10,
              bold=(nama == "Dompis Berlin Thary"))
cells = tbl2.add_row().cells
for i, v in enumerate(["Harga di tangan pembeli", rp(HARGA_AKHIR), rp(HARGA_AKHIR)]):
    cell_text(cells[i], v, bold=True, center=(i > 0), size=10); shade(cells[i], "F2F2F2")
set_grid(tbl2, [Cm(5.0), Cm(4.5), Cm(6.5)])
assert sum(a for _, a, _ in skenario) == HARGA_AKHIR and sum(b for _, _, b in skenario) == HARGA_AKHIR
para("Tabel 2. Nilai tambah per bungkus pada dua jalur penjualan", italic=True, align="center", size=10, space_after=8)

para("Sumber angka: harga pisang berlin di tingkat petani Rp10.000 dan di pedagang Rp15.000 per tandan (Suara "
     "Indonesia, 2022) serta Rp20.000 sampai Rp35.000 per tandan di pasar induk (Jatimnow, 2020), saya "
     "sesuaikan ke harga Palopo sekarang dan dibagi per sisir; harga dompo Rp10.000 sampai Rp20.000 per bungkus "
     "dari kios oleh-oleh di Luwu Timur (Jadesta Kemenpar) dan toko daring Makassar; komposisi biaya tepung, "
     "minyak, dan kemasan dari studi nilai tambah sale pisang skala rumah tangga (Yandra dkk., 2025; Dyah, "
     "2025). Semua angka dibulatkan dan masih perlu dikonfirmasi ke Thary, pedagang pasar, dan toko.",
     italic=True, size=10)

# ---------- 4. Analisis per pelaku ----------
heading("4. Membaca angka: kerja, risiko, dan informasi tiap pelaku")
para("Angka di Tabel 1 baru bermakna kalau disandingkan dengan apa yang ditanggung tiap pelaku. Slide "
     "pertemuan ini mengingatkan bahwa yang perlu ditanyakan adalah apakah pembagiannya sepadan dengan yang "
     "ditanggung masing-masing.")
para("Petani menerima Rp4.000 per sisir, 20 persen dari harga akhir, untuk menanam dan menunggu sekitar "
     "sepuluh bulan sampai satu tandan siap panen. Biaya tunainya memang kecil, tetapi risikonya nyata: pisang "
     "yang matang serentak harus dijual dalam hitungan hari dengan harga berapa pun yang ditawarkan pedagang, "
     "karena petani tidak punya gudang, alat pengering, dan pembeli lain.")
para("Pedagang pasar juga menerima Rp4.000 per sisir, tetapi ia memutar uangnya dalam dua sampai tiga hari, "
     "sedangkan petani menunggu berbulan-bulan. Ia menanggung ongkos angkut dan pisang yang busuk. Sebagai "
     "orang pertama di rantai yang tahu jenis pisang apa yang dicari pembeli kota, pedagang bisa menahan margin "
     "sebesar margin petani untuk kerja yang jauh lebih singkat.")
para("Thary menerima Rp3.500 per bungkus kalau menitip di toko. Untuk itu ia membeli pisang, mengupas dan "
     "membelah 12 sampai 15 buah, menjemur dua sampai tiga hari sambil menjaga dari hujan dan debu, menggoreng, "
     "mengemas, dan mengurus izin PIRT serta label. Risiko produk yang ditolak toko atau kedaluwarsa di rak juga "
     "ada di pundaknya. Dari sisi jam kerja, ia jelas yang paling banyak bekerja per bungkus, tetapi "
     "bagiannya justru paling kecil di antara empat pelaku pada jalur toko.")
para("Toko oleh-oleh menerima Rp5.000, bagian terbesar, untuk memajang dan menjual. Toko membayar sewa dan "
     "menanggung barang yang tidak laku, tetapi tidak membuat apa pun. Kekuatannya ada pada lokasi dan "
     "informasi: toko tahu siapa perantau yang datang, kapan musim mudik, dan berapa harga yang masih dianggap "
     "wajar untuk oleh-oleh. Ini pola yang sama dengan toko oleh-oleh pada contoh bakul rotan di kelas.")
para("Tabel 2 menunjukkan apa yang terjadi kalau Thary memakai informasi pasar itu sendiri. Dengan menjual "
     "langsung lewat WhatsApp, Instagram, dan rumah produksi, bagiannya naik dari Rp3.500 menjadi Rp8.500, "
     "dari 17,5 persen menjadi 42,5 persen harga akhir, tanpa menaikkan harga ke pembeli. Yang ia bayar adalah "
     "waktu untuk melayani pesanan dan biaya kirim yang ditanggung pembeli. Inilah yang dimaksud Dai dkk. "
     "(2024) dengan ketegangan antara menciptakan nilai dan menangkap nilai: pada jalur toko, nilai diciptakan "
     "di dapur Thary tetapi paling banyak ditangkap di rak toko; pada jalur langsung, pencipta nilai dan "
     "penangkap nilainya orang yang sama.")

# ---------- 5. Tiga ukuran ----------
heading("5. Menilai rantai dengan tiga ukuran")
para("Dari sisi laba, setiap pelaku masih untung dan usaha ini bertahan sejak 2016, sudah hampir sepuluh tahun. "
     "Pesanan datang dari luar kota, dan Dinas Koperasi dan UMKM menjadikannya UMKM binaan. Titik lemahnya "
     "adalah produksi yang bergantung pada matahari: pada musim hujan, penjemuran bisa gagal dan pasokan ke "
     "toko terhenti.")
para("Dari sisi manusia, petani dan Thary adalah dua pelaku yang paling banyak bekerja dan sama-sama menerima "
     "bagian kecil pada jalur toko, 20 persen dan 17,5 persen. Pedagang dan toko, yang bekerja lebih singkat, "
     "menerima 45 persen. Bedanya, Thary sudah punya jalan keluar lewat penjualan langsung, sedangkan petani "
     "belum. Usaha ini juga memberi upah bagi ibu-ibu sekitar yang membantu mengupas dan mengemas saat pesanan ramai.")
para("Dari sisi lingkungan, rantai ini menyelamatkan pisang yang tadinya terbuang dan memakai energi matahari "
     "untuk mengeringkan. Kulit pisang kembali ke tanah atau ternak. Sisa yang belum tertangani adalah minyak "
     "jelantah dan kemasan plastik, dua hal yang bisa jadi bahan proposal akhir: minyak jelantah bisa disetor ke "
     "pengumpul, dan kemasan bisa diganti kertas untuk pembeli lokal.")

# ---------- 6. Paragraf penutup ----------
heading("6. Siapa yang bagiannya paling kecil, dan apakah itu sepadan?")
para("Dari angka di Tabel 1, bagian paling kecil pada jalur toko diambil Thary, Rp3.500 dari Rp20.000, padahal "
     "ialah yang membeli pisang, mengupas, menjemur tiga hari, menggoreng, mengemas, dan mengurus izin. Toko yang "
     "hanya memajang mendapat Rp5.000, dan pedagang pasar yang memegang pisang dua hari mendapat Rp4.000. Menurut "
     "saya pembagian ini tidak sepadan dengan kerja dan risiko yang ditanggung. Penyebabnya adalah informasi "
     "pasar yang dipegang toko: ia tahu siapa pembeli oleh-oleh dan berapa mereka sanggup bayar, sementara Thary "
     "di dapur tidak. Itulah pola yang berulang dari bakul rotan di kelas sampai "
     "dompo di Palopo: semakin dekat ke konsumen, semakin besar nilai yang diambil. Thary sudah menemukan "
     "jawabannya sendiri. Dengan menjual langsung lewat WhatsApp dan "
     "Instagram, ia memangkas perantara, memakai merek Dompis sebagai jaminan mutu, dan berbicara langsung dengan "
     "pembeli, sehingga bagiannya naik menjadi Rp8.500, yang terbesar di rantai, tanpa membuat pembeli membayar "
     "lebih. Bagian yang paling kecil dan paling sulit diperbaiki justru milik petani: Rp4.000 untuk sepuluh "
     "bulan menunggu, tanpa pilihan pembeli lain. Kalau Thary membeli pisang berlin langsung dari petani dengan "
     "harga di antara Rp4.000 dan Rp8.000, petani naik bagiannya, Thary tetap menghemat, dan pedagang pasar "
     "kehilangan pisang yang sebetulnya sering ia buang. Dari sudut pandang itu, rantai pendek yang dibangun "
     "Thary lebih menguntungkan baginya dan lebih adil bagi orang yang paling lama bekerja di rantai ini.")

# ---------- 7. Pertanyaan konfirmasi ----------
heading("7. Angka yang perlu dikonfirmasi ke pelaku")
para("Semua angka di atas adalah perkiraan dari sumber daring. Sebelum pertemuan keempat saya akan menanyakan "
     "hal berikut kepada Thary, seorang pedagang pisang di pasar, dan satu toko oleh-oleh, lalu memperbarui "
     "Tabel 1 dan Gambar 1.", space_after=2)
pertanyaan = [
    "Kepada Thary: harga beli pisang berlin per sisir atau per tandan, berapa sisir untuk satu bungkus 200 gram, "
    "harga jual ke toko dan harga jual langsung, biaya kemasan dan label per bungkus, serta berapa persen "
    "penjualan yang lewat toko.",
    "Kepada pedagang pasar: harga beli dari petani per tandan, asal pisang berlin, dan berapa banyak pisang "
    "yang terbuang atau dilepas murah tiap minggu.",
    "Kepada toko oleh-oleh: harga ambil dan harga jual dompo, sistem titip atau beli putus, dan berapa lama "
    "dompo bertahan di rak.",
]
first = numbered(pertanyaan[0])
new_id = restart_numbering(first)
for q in pertanyaan[1:]:
    numbered(q, num_id=new_id)

# ---------- Referensi ----------
heading("Referensi")
refs = [
    "Astaginy, N., & Sudarnice. (2019). Strategi pengembangan usaha dempo pisang sebagai produk keunggulan daerah "
    "(studi kasus pembuatan dempo pisang di Kabupaten Bombana). BISEI: Jurnal Bisnis dan Ekonomi Islam, 4(2), 68–74. "
    "https://doi.org/10.33752/bisei.v4i02.606",
    "Dai, G., Zhang, L., Zhang, Q., & Mao, M. (2024). Navigating tensions between value creation and value "
    "capture in ecosystems. Journal of Business Research, 170, 114333. https://doi.org/10.1016/j.jbusres.2023.114333",
    "Dyah, M. (2025). Analisis nilai tambah usaha pengolahan pisang sale skala rumah tangga di Desa Banyu Urip, "
    "Kecamatan Tanjung Lago, Kabupaten Banyuasin [Skripsi, Universitas Tridinanti]. "
    "http://repository.univ-tridinanti.ac.id/10528/",
    "Hairon, I. (2022, 22 Juni). Di Jember, pisang berlin dijual sangat murah. Suara Indonesia. "
    "https://suaraindonesia.co.id/news/features/62b2aebc1a9e4/di-jember-pisang-berlin-dijual-sangat-murah",
    "Jatimnow. (2020, 29 Oktober). Imut dan rasa sedikit masam, pisang barlin di PIOS diburu pengunjung. "
    "https://jatimnow.com/baca-30937-imut-dan-rasa-sedikit-masam-pisang-barlin-di-pios-diburu-pengunjung",
    "Kementerian Pariwisata. (t.t.). Produk wisata dompo pisang, Desa Wisata Dermaga Desa Pasi-pasi, Luwu Timur. "
    "Jadesta. https://jadesta.kemenpar.go.id/paket/dompo_pisang",
    "Nandini, A. B. (2024, 12 November). Dompis Berlin Thary, olahan pisang khas Kota Palopo Sulsel yang bikin "
    "nagih. Tribun Timur. https://makassar.tribunnews.com/2024/11/12/dompis-berlin-thary-olahan-pisang-khas-kota-"
    "palopo-sulsel-yang-bikin-nagih",
    "Putri, Q. A. R. (2026). Pertemuan 3: Ekosistem dan rantai nilai [Slide perkuliahan Ekonomi Kreatif MBS12523]. "
    "Manajemen Bisnis Syariah, FEBI UIN Palopo.",
    "Universitas Muhammadiyah Malang. (2003). Pengaruh suhu dan lama pengeringan pisang terhadap mutu sale pisang. "
    "Department of Agribisnis Student Research. http://student-research.umm.ac.id/index.php/dept_of_agribisnis/article/view/3164",
    "Yandra, H. M., Hakimi, R., & Raesi, S. (2025). Analisis nilai tambah pengolahan pisang (Musa paradisiaca L.) "
    "pada UMKM Heni Pisang Sale di Nagari IV Koto, Kecamatan Pulau Punjung, Kabupaten Dharmasraya. JAS (Jurnal "
    "Agri Sains), 9(1). https://ojs.umb-bungo.ac.id/index.php/JAS/article/download/1725/1384",
]
for r_ in refs:
    p = para(r_, size=11, space_after=4)
    p.paragraph_format.left_indent = Cm(1); p.paragraph_format.first_line_indent = Cm(-1)
    p.paragraph_format.line_spacing = 1.15

doc.save("tugas-3/Tugas-3_Rantai_Nilai_Dompo_Pisang_Berlin.docx")
print("saved")
