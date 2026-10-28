# -*- coding: utf-8 -*-
"""Genera logo.png placeholder (testo bianco su trasparente, visibile
sulla barra colorata dell'intestazione email). Il centro lo sostituirà
col proprio logo mantenendo lo stesso nome file."""
from PIL import Image, ImageDraw, ImageFont

W, H = 260, 64
img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# quadratino arrotondato come marchietto + scritta
box = 44
d.rounded_rectangle([0, 10, box, 10 + box], radius=9, fill=(255, 255, 255, 235))
try:
    f_big = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 22)
    f_box = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 24)
except Exception:
    f_big = f_box = ImageFont.load_default()

# lettera R dentro il marchietto (colore blu tema)
d.text((box / 2 + 1, 10 + box / 2), "R", font=f_box, anchor="mm", fill=(31, 59, 102, 255))

d.text((box + 12, 32), "RevMemo", font=f_big, anchor="lm", fill=(255, 255, 255, 255))

img.save("/home/z/my-project/download/RevMemo/logo.png")
print("logo.png creato")
