import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
import textwrap
import os

# --- DOSYA YOLU / EXCEL ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
df = pd.read_excel("urunler.xlsx")

# --- PDF ---
pdf = canvas.Canvas("katalog.pdf", pagesize=A4)
width, height = A4

# --- KAPAK ---
kapak = ImageReader(os.path.join(BASE_DIR, "kapak.png"))
pdf.drawImage(kapak, 0, 0, width, height)
pdf.showPage()

# =========================
# GÜVENLİ SAYFA BOŞLUKLARI
# =========================
UST_SERIT = 25
ALT_SERIT = 25

def ciz_serit():
    pdf.setFillColorRGB(0.9, 0, 0)
    pdf.rect(0, height - UST_SERIT, width, UST_SERIT, fill=1)
    pdf.rect(0, 0, width, ALT_SERIT, fill=1)

ciz_serit()

# =========================
# GRID
# =========================
cols, rows = 4, 4
kullanilabilir_yukseklik = height - UST_SERIT - ALT_SERIT
box_w = width / cols
box_h = kullanilabilir_yukseklik / rows

x_start = 0
y_start = height - UST_SERIT - box_h
x = x_start
y = y_start
sayac = 0

# --- SABİT KONUM DEĞERLERİ ---
ISIM_Y_OFFSET = box_h - 35
KOD_Y_OFFSET  = box_h - 60

RESIM_ALT_OFFSET = box_h - 150
RESIM_UST_OFFSET = box_h - 55

KIRMIZI_Y_OFFSET = 22
ALTIN_Y_OFFSET   = 10

FIYAT_GENISLIK = box_w - 48

for _, row in df.iterrows():

    # --- KART ÇERÇEVESİ ---
    pdf.setStrokeColorRGB(0.9, 0, 0)
    pdf.roundRect(x+6, y+6, box_w-12, box_h-12, 14, stroke=1, fill=0)

    # --- ÜRÜN ADI ---
    pdf.setFillColorRGB(0, 0, 0)
    pdf.setFont("Helvetica-Bold", 8)
    isimler = textwrap.wrap(str(row["ad"]), 28)
    y_isim = y + ISIM_Y_OFFSET
    for s in isimler[:2]:
        pdf.drawString(x + 12, y_isim, s)
        y_isim -= 10

    # --- KOD + KATEGORİ ---
    pdf.setFont("Helvetica", 7)
    pdf.setFillColorRGB(0, 0, 0)
    pdf.drawString(x + 12, y + KOD_Y_OFFSET, str(row["kod"]))

    pdf.setFillColorRGB(0.4, 0.4, 0.4)
    pdf.drawRightString(
        x + box_w - 16,
        y + KOD_Y_OFFSET,
        str(row["kategori"])
    )

    # =========================
    # RESİM
    # =========================
    raw = str(row["resim"]).strip()
    dosya_adi = os.path.splitext(os.path.basename(raw))[0]

    aday_yollar = [
        os.path.join(BASE_DIR, "img", dosya_adi + ".png"),
        os.path.join(BASE_DIR, "img", dosya_adi + ".jpg"),
        os.path.join(BASE_DIR, "img", dosya_adi + ".jpeg"),
    ]

    resim_yolu = next((yol for yol in aday_yollar if os.path.exists(yol)), None)

    resim_alt = y + RESIM_ALT_OFFSET
    resim_ust = y + RESIM_UST_OFFSET
    alan_w = box_w - 28
    alan_h = resim_ust - resim_alt

    if resim_yolu:
        img = ImageReader(resim_yolu)
        iw, ih = img.getSize()
        oran = min(alan_w / iw, alan_h / ih)
        pdf.drawImage(
            img,
            x + 14 + (alan_w - iw * oran) / 2,
            resim_alt + (alan_h - ih * oran) / 2,
            iw * oran,
            ih * oran,
            mask='auto'
        )

    # =========================
    # KIRMIZI FİYAT
    # =========================
    fiyat_x = x + 24
    pdf.setFillColorRGB(0.9, 0, 0)
    pdf.roundRect(
        fiyat_x, y + KIRMIZI_Y_OFFSET,
        FIYAT_GENISLIK, 22, 10,
        fill=1, stroke=0
    )

    pdf.setFillColor(colors.white)
    pdf.setFont("Helvetica-Bold", 9)
    pdf.drawCentredString(
        fiyat_x + FIYAT_GENISLIK / 2,
        y + KIRMIZI_Y_OFFSET + 7,
        str(row["fiyat"])
    )

    # =========================
    # ALTIN (ESKİ) FİYAT
    # =========================
#    pdf.setFillColorRGB(0.85, 0.65, 0.13)
#    pdf.roundRect(
#        fiyat_x, y + ALTIN_Y_OFFSET,
#        FIYAT_GENISLIK, 22, 10,
#        fill=1, stroke=0
#    )

#    pdf.setFillColor(colors.white)
#    pdf.drawCentredString(
#        fiyat_x + FIYAT_GENISLIK / 2,
#        y + ALTIN_Y_OFFSET + 7,
#        str(row["eski_fiyat"])
#    )

    # --- GRID İLERLEME ---
    sayac += 1
    x += box_w

    if sayac % cols == 0:
        x = x_start
        y -= box_h

    if sayac % 16 == 0:
        pdf.showPage()
        ciz_serit()
        x = x_start
        y = y_start

pdf.save()
print("Katalog olusturuldu.")
