"""Diagram rantai nilai 1 bungkus (200 g) dompo pisang berlin Dompis Berlin Thary, Palopo."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

def rp(n):
    return "Rp" + f"{n:,.0f}".replace(",", ".")

HARGA_AKHIR = 20000
# (pelaku, tahap, bahan dibeli, harga jual). Satuan: satu sisir pisang berlin -> satu bungkus dompo 200 g.
aktor = [
    ("Petani pisang berlin", "KREASI", 0, 3500),
    ("Pedagang pisang\ndi pasar Palopo", "DISTRIBUSI", 3500, 7500),
    ("Dompis Berlin Thary\n(kupas, jemur, goreng, kemas)", "PRODUKSI", 7500, 14000),
    ("Toko oleh-oleh\nPalopo", "DISTRIBUSI", 14000, HARGA_AKHIR),
    ("Pembeli", "KONSUMSI", HARGA_AKHIR, None),
]
warna = ["#c8e6c9", "#fff3c4", "#ffd8b8", "#f8bbd0", "#dbe5ff"]

fig = plt.figure(figsize=(16, 6.4), dpi=150)
fig.patch.set_facecolor("white")
ax = fig.add_axes([0.02, 0.03, 0.96, 0.82])
ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis("off")

w, h, gap = 16.5, 80, 3.3
x0 = (100 - (5 * w + 4 * gap)) / 2
nilai_tambah = []
for i, (nama, tahap, beli, jual) in enumerate(aktor):
    x = x0 + i * (w + gap)
    ax.add_patch(FancyBboxPatch((x, 8), w, h, boxstyle="round,pad=0.4,rounding_size=2",
                                fc=warna[i], ec="#444", lw=1.3))
    ax.text(x + w / 2, 96, tahap, ha="center", va="center", fontsize=10.5, fontweight="bold", color="#333")
    ax.text(x + w / 2, 76, nama, ha="center", va="center", fontsize=10.8, fontweight="bold")
    if jual is None:
        ax.text(x + w / 2, 48, "Membayar\n" + rp(beli), ha="center", va="center", fontsize=11.5)
        ax.text(x + w / 2, 24, "harga di tangan pembeli,\n1 bungkus 200 g", ha="center", va="center",
                fontsize=9.5, style="italic", color="#333")
    else:
        nt = jual - beli
        nilai_tambah.append((nama.split("\n")[0], nt))
        if beli == 0:
            beli_txt = "Bahan dibeli: tidak ada\n(anakan dan pupuk\ndari kebun sendiri)"
        elif i in (1, 2):
            beli_txt = "Bahan dibeli: " + rp(beli) + "\n(1 sisir pisang)"
        else:
            beli_txt = "Bahan dibeli: " + rp(beli) + "\n(1 bungkus dompo)"
        ax.text(x + w / 2, 54, beli_txt, ha="center", va="center", fontsize=9.3)
        ax.text(x + w / 2, 35, "Dijual: " + rp(jual), ha="center", va="center", fontsize=10.8)
        ax.text(x + w / 2, 20, "Nilai tambah\n" + rp(nt), ha="center", va="center",
                fontsize=12, fontweight="bold", color="#b71c1c")
    if i < 4:
        ax.add_patch(FancyArrowPatch((x + w + 0.3, 48), (x + w + gap - 0.3, 48),
                                     arrowstyle="-|>", mutation_scale=22, lw=2, color="#333"))

fig.text(0.5, 0.955, "Rantai nilai satu bungkus dompo pisang berlin (200 g) Dompis Berlin Thary, Kota Palopo",
         ha="center", fontsize=15, fontweight="bold")
fig.text(0.5, 0.905, "Dari bahan mentah sampai ke tangan pembeli. Angka perkiraan; satu sisir pisang berlin menjadi "
         "satu bungkus dompo; nilai tambah = harga jual dikurangi bahan yang dibeli",
         ha="center", fontsize=10.2, color="#444")

fig.savefig("tugas-3/rantai_nilai_dompo_pisang_berlin.png", bbox_inches="tight", facecolor="white")
total = sum(v for _, v in nilai_tambah)
assert total == HARGA_AKHIR, total
assert len({v for _, v in nilai_tambah}) == 4, "nilai tambah harus berbeda tiap pelaku"
print("ok", nilai_tambah, total)
