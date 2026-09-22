"""Diagram rantai nilai 1 bungkus (200 g) dompo pisang berlin Dompis Berlin Thary, Palopo."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

def rp(n):
    return "Rp" + f"{n:,.0f}".replace(",", ".")

HARGA_AKHIR = 20000
BAHAN_LAIN = 1700      # tepung tapioka, minyak goreng, gas
KEMASAN = 1800         # standing pouch dan stiker label

# (pelaku, tahap, bahan dibeli, harga jual). Satuan: satu sisir pisang berlin -> satu bungkus 200 g.
aktor = [
    ("Petani pisang berlin\n(kebun campuran Luwu)", "KREASI", 0, 4000),
    ("Pedagang pisang\ndi pasar Palopo", "DISTRIBUSI 1", 4000, 8000),
    ("Dompis Berlin Thary\n(kupas, jemur, goreng, kemas)", "PRODUKSI", 8000 + BAHAN_LAIN + KEMASAN, 15000),
    ("Toko oleh-oleh\nPalopo", "DISTRIBUSI 2", 15000, HARGA_AKHIR),
    ("Pembeli\n(konsumen)", "KONSUMSI", HARGA_AKHIR, None),
]
warna = ["#c8e6c9", "#fff3c4", "#ffd8b8", "#f8bbd0", "#dbe5ff"]

fig = plt.figure(figsize=(16, 9.6), dpi=150)
fig.patch.set_facecolor("white")
ax = fig.add_axes([0.02, 0.40, 0.96, 0.55])
ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis("off")

w, h, gap = 16.5, 58, 3.3
x0 = (100 - (5 * w + 4 * gap)) / 2
nilai_tambah = []
xs = []
for i, (nama, tahap, beli, jual) in enumerate(aktor):
    x = x0 + i * (w + gap); xs.append(x)
    ax.add_patch(FancyBboxPatch((x, 22), w, h, boxstyle="round,pad=0.4,rounding_size=2",
                                fc=warna[i], ec="#444", lw=1.3))
    ax.text(x + w / 2, 87, tahap, ha="center", va="center", fontsize=10.5, fontweight="bold", color="#333")
    ax.text(x + w / 2, 73, nama, ha="center", va="center", fontsize=10.8, fontweight="bold")
    if jual is None:
        ax.text(x + w / 2, 52, "Membayar\n" + rp(beli), ha="center", va="center", fontsize=11)
        ax.text(x + w / 2, 33, "harga akhir\n1 bungkus 200 g", ha="center", va="center",
                fontsize=9.5, style="italic", color="#333")
    else:
        nt = jual - beli
        nilai_tambah.append((nama.split("\n")[0], nt))
        if beli == 0:
            beli_txt = "Bahan dibeli: tidak ada\n(anakan dan pupuk\ndari kebun sendiri)"
        elif i == 2:
            beli_txt = ("Bahan dibeli: " + rp(beli) + "\n(pisang 1 sisir Rp8.000,\ntepung, minyak, gas Rp1.700,\nkemasan dan label Rp1.800)")
        else:
            beli_txt = "Bahan dibeli: " + rp(beli) + "\n(1 sisir pisang)" if i == 1 else "Bahan dibeli: " + rp(beli) + "\n(1 bungkus dompo)"
        ax.text(x + w / 2, 55, beli_txt, ha="center", va="center", fontsize=9.3)
        ax.text(x + w / 2, 40, "Dijual: " + rp(jual), ha="center", va="center", fontsize=10.5)
        ax.text(x + w / 2, 29, "Nilai tambah\n" + rp(nt), ha="center", va="center",
                fontsize=11.5, fontweight="bold", color="#b71c1c")
    if i < 4:
        ax.add_patch(FancyArrowPatch((x + w + 0.3, 51), (x + w + gap - 0.3, 51),
                                     arrowstyle="-|>", mutation_scale=22, lw=2, color="#333"))

# Jalur langsung: Thary menjual sendiri lewat WA, Instagram, dan rumah produksi
xa = xs[2] + w / 2; xb = xs[4] + w / 2
ax.add_patch(FancyArrowPatch((xa, 22), (xb, 22), connectionstyle="arc3,rad=0.16",
                             arrowstyle="-|>", mutation_scale=22, lw=1.8, ls="--", color="#1565c0"))
ax.text((xa + xb) / 2, 3.0, "Jalur langsung: Thary menjual sendiri Rp20.000 lewat WA, Instagram, dan rumah produksi.\n"
        "Nilai tambah Thary naik dari Rp3.500 menjadi Rp8.500; bagian toko hilang.",
        ha="center", va="center", fontsize=9.6, color="#1565c0")

fig.text(0.5, 0.965, "Rantai nilai satu bungkus dompo pisang berlin (200 g) Dompis Berlin Thary, Kota Palopo",
         ha="center", fontsize=15, fontweight="bold")
fig.text(0.5, 0.937, "Alur dari bahan mentah sampai ke tangan pembeli. Angka perkiraan; satuan: satu sisir pisang berlin "
         "menjadi satu bungkus dompo; nilai tambah = harga jual dikurangi bahan yang dibeli",
         ha="center", fontsize=10.2, color="#444")

# Konservasi
axk = fig.add_axes([0.04, 0.335, 0.92, 0.05]); axk.axis("off")
axk.add_patch(FancyBboxPatch((0.0, 0.05), 1.0, 0.9, boxstyle="round,pad=0.01,rounding_size=0.02",
                             fc="#f1f8e9", ec="#666", lw=1, ls="--", transform=axk.transAxes))
axk.text(0.5, 0.5, "KONSERVASI: pisang lewat matang yang tadinya dibuang menjadi bahan baku; kulit pisang jadi pakan "
         "atau kompos; penjemuran memakai matahari tanpa bahan bakar; minyak jelantah dan kemasan plastik "
         "tertinggal sebagai sampah", ha="center", va="center", fontsize=9.4, color="#333", transform=axk.transAxes)

# Grafik batang nilai tambah
ax2 = fig.add_axes([0.10, 0.05, 0.56, 0.25])
nama = [n for n, _ in nilai_tambah][::-1]
nt = [v for _, v in nilai_tambah][::-1]
bars = ax2.barh(nama, nt, color=warna[:4][::-1], edgecolor="#444")
for b, v in zip(bars, nt):
    ax2.text(v + 120, b.get_y() + b.get_height() / 2, rp(v), va="center", fontsize=10.5, fontweight="bold")
# Nilai tambah Thary pada jalur langsung, sebagai batang arsir di baris yang sama
idx = nama.index("Dompis Berlin Thary")
ax2.barh([idx], [8500], color="none", edgecolor="#1565c0", hatch="///", lw=1.2)
ax2.text(8500 + 120, idx + 0.28, "Rp8.500 jika jual langsung", va="center", fontsize=9, color="#1565c0")
ax2.set_xlim(0, 11000)
ax2.set_xlabel("Nilai tambah per bungkus (Rp)", fontsize=10)
ax2.set_title("Di titik mana uang paling banyak berhenti? (jalur lewat toko oleh-oleh)",
              fontsize=11.5, fontweight="bold", loc="left")
ax2.spines[["top", "right"]].set_visible(False)
ax2.tick_params(axis="y", labelsize=10)
ticks = range(0, 11001, 2000)
ax2.set_xticks(list(ticks)); ax2.set_xticklabels([rp(t) for t in ticks], fontsize=9)

lain = BAHAN_LAIN + KEMASAN
fig.text(0.70, 0.285, "Jalur lewat toko (harga akhir Rp20.000):", fontsize=10.5, va="top", fontweight="bold")
fig.text(0.70, 0.255, "petani Rp4.000 (20%), pedagang pasar Rp4.000 (20%),\nThary Rp3.500 (17,5%), toko Rp5.000 (25%),\n"
         f"pemasok tepung, minyak, gas, kemasan {rp(lain)} (17,5%)", fontsize=10, va="top")
fig.text(0.70, 0.165, "Yang paling banyak bekerja (Thary) mendapat bagian\npaling kecil kalau menitip di toko. "
         "Begitu menjual\nsendiri, bagiannya jadi yang terbesar.", fontsize=10, va="top", style="italic", color="#444")

fig.savefig("tugas-3/rantai_nilai_dompo_pisang_berlin.png", bbox_inches="tight", facecolor="white")
total = sum(v for _, v in nilai_tambah) + lain
assert total == HARGA_AKHIR, total
print("ok", nilai_tambah, total)
