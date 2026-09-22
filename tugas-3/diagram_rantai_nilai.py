"""Diagram rantai nilai 1 bungkus (200 g) kopi arabika Latimojong yang dijual di Palopo."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

def rp(n):
    return "Rp" + f"{n:,.0f}".replace(",", ".")

# (pelaku, tahap, beli, jual). Angka perkiraan, satuan: bahan untuk 1 bungkus 200 g.
aktor = [
    ("Petani kopi\nLatimojong / Bastem", "KREASI", 0, 19200),
    ("Pengepul\nkecamatan", "DISTRIBUSI 1", 19200, 24000),
    ("Roastery / pengolah\n(Belopa - Palopo)", "PRODUKSI", 27000, 45000),
    ("Toko oleh-oleh /\nkedai kopi Palopo", "DISTRIBUSI 2", 45000, 65000),
    ("Pembeli\n(konsumen)", "KONSUMSI", 65000, None),
]
warna = ["#c8e6c9", "#fff3c4", "#ffd8b8", "#f8bbd0", "#dbe5ff"]

fig = plt.figure(figsize=(16, 9.2), dpi=150)
fig.patch.set_facecolor("white")
ax = fig.add_axes([0.02, 0.42, 0.96, 0.53])
ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis("off")

w, h, gap = 16.5, 62, 3.3
x0 = (100 - (5 * w + 4 * gap)) / 2
nilai_tambah = []
for i, (nama, tahap, beli, jual) in enumerate(aktor):
    x = x0 + i * (w + gap)
    box = FancyBboxPatch((x, 18), w, h, boxstyle="round,pad=0.4,rounding_size=2",
                         fc=warna[i], ec="#444", lw=1.3)
    ax.add_patch(box)
    ax.text(x + w / 2, 87, tahap, ha="center", va="center", fontsize=10.5,
            fontweight="bold", color="#333")
    ax.text(x + w / 2, 72, nama, ha="center", va="center", fontsize=11.5, fontweight="bold")
    if jual is None:
        ax.text(x + w / 2, 50, "Membayar\n" + rp(beli), ha="center", va="center", fontsize=11)
        ax.text(x + w / 2, 30, "harga akhir\n1 bungkus 200 g", ha="center", va="center",
                fontsize=9.5, style="italic", color="#333")
    else:
        nt = jual - beli
        nilai_tambah.append((nama.replace("\n", " "), nt))
        beli_txt = "Bahan dibeli: tidak ada" if beli == 0 else "Bahan dibeli: " + rp(beli)
        if i == 2:
            beli_txt += "\n(biji 240 g Rp24.000\n+ kemasan Rp3.000)"
        ax.text(x + w / 2, 52, beli_txt, ha="center", va="center", fontsize=9.8)
        ax.text(x + w / 2, 38, "Dijual: " + rp(jual), ha="center", va="center", fontsize=10.5)
        ax.text(x + w / 2, 25, "Nilai tambah\n" + rp(nt), ha="center", va="center",
                fontsize=11.5, fontweight="bold", color="#b71c1c")
    if i < 4:
        ax.add_patch(FancyArrowPatch((x + w + 0.3, 49), (x + w + gap - 0.3, 49),
                                     arrowstyle="-|>", mutation_scale=22, lw=2, color="#333"))

# Tahap kelima: konservasi (garis putus-putus di bawah rantai)
ax.add_patch(FancyBboxPatch((x0, 2), 5 * w + 4 * gap, 11, boxstyle="round,pad=0.3,rounding_size=1.5",
                            fc="#f1f8e9", ec="#666", lw=1, ls="--"))
ax.text(50, 7.5, "KONSERVASI: kulit buah kopi kembali ke kebun sebagai pupuk, ampas seduhan bisa jadi kompos, "
        "kemasan aluminium foil belum didaur ulang dan tertinggal sebagai sampah",
        ha="center", va="center", fontsize=9.6, color="#333")

fig.text(0.5, 0.965, "Rantai nilai 1 bungkus kopi arabika Latimojong (200 g) yang dijual di Kota Palopo",
         ha="center", fontsize=15, fontweight="bold")
fig.text(0.5, 0.935, "Alur dari bahan mentah sampai ke tangan pembeli. Angka perkiraan; "
         "nilai tambah = harga jual dikurangi bahan yang dibeli", ha="center", fontsize=10.5, color="#444")

# Grafik batang nilai tambah
ax2 = fig.add_axes([0.10, 0.06, 0.60, 0.30])
nama = [n for n, _ in nilai_tambah][::-1]
nt = [v for _, v in nilai_tambah][::-1]
bars = ax2.barh(nama, nt, color=warna[:4][::-1], edgecolor="#444")
for b, v in zip(bars, nt):
    ax2.text(v + 400, b.get_y() + b.get_height() / 2, rp(v), va="center", fontsize=10.5, fontweight="bold")
ax2.set_xlim(0, 26000)
ax2.set_xlabel("Nilai tambah per bungkus (Rp)", fontsize=10)
ax2.set_title("Di titik mana uang paling banyak berhenti?", fontsize=12, fontweight="bold", loc="left")
ax2.spines[["top", "right"]].set_visible(False)
ax2.tick_params(axis="y", labelsize=10)
ax2.set_xticks(range(0, 26001, 5000)); ax2.set_xticklabels([rp(t) for t in range(0, 26001, 5000)], fontsize=9)

fig.text(0.735, 0.30, "Bagian petani dan pengepul:\n" + rp(19200 + 4800) + " dari " + rp(65000) + " (37%)",
         fontsize=11, va="top", fontweight="bold")
fig.text(0.735, 0.22, "Bagian roastery dan toko/kedai:\n" + rp(18000 + 20000) + " dari " + rp(65000) + " (58%)\n"
         "sisanya Rp3.000 ke pemasok kemasan",
         fontsize=11, va="top")
fig.text(0.735, 0.10, "Pola yang sama dengan bakul rotan di kelas:\nsemakin dekat ke konsumen,\n"
         "semakin besar nilai tambah yang diambil.", fontsize=10, va="top", style="italic", color="#444")

fig.savefig("tugas-3/rantai_nilai_kopi_latimojong.png", bbox_inches="tight", facecolor="white")
print("ok", nilai_tambah, sum(v for _, v in nilai_tambah) + 3000)
