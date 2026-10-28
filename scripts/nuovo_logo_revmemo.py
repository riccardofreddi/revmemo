#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Rigenera logo.png per RevMemo: badge bianco arrotondato con scritta
blu tema. Visibile su QUALSIASI sfondo (testa colorata delle email,
fondo bianco, modalita' scura). Il centro lo sostituira' col proprio
logo PNG mantenendo lo stesso nome file."""
from PIL import Image, ImageDraw, ImageFont

PERCORSO = "/home/z/my-project/download/RevMemo/logo.png"
BLU = (31, 59, 102, 255)        # #1F3B66 (colore tema di default)
BIANCO = (255, 255, 255, 255)

W, H = 184, 52
img = Image.new("RGBA", (W, H), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# badge bianco con bordo sottile blu (cosi' e' riconoscibile anche sul bianco)
d.rounded_rectangle([0, 0, W - 1, H - 1], radius=11, fill=BIANCO,
                    outline=(31, 59, 102, 255), width=2)

font = ImageFont.truetype(
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 26)
testo = "RevMemo"
bb = d.textbbox((0, 0), testo, font=font)
tw, th = bb[2] - bb[0], bb[3] - bb[1]
x = (W - tw) / 2 - bb[0]
y = (H - th) / 2 - bb[1] - 2
d.text((x, y), testo, font=font, fill=BLU)

# pallino colorato come accento (richiama il colore tema del centro)
d.ellipse([10, H / 2 - 4, 18, H / 2 + 4], fill=BLU)

img.save(PERCORSO)
print(f"Logo salvato: {PERCORSO}  ({img.size[0]}x{img.size[1]}, "
      f"{__import__('os').path.getsize(PERCORSO)} byte)")

# verifica visiva: salva anche un'anteprima su fondo blu e su fondo bianco
for sfondo, nome in (((31, 59, 102, 255), "blu"), ((255, 255, 255, 255), "bianco")):
    ante = Image.new("RGBA", (W + 40, H + 24), sfondo)
    ante.alpha_composite(img, (20, 12))
    ante.convert("RGB").save(f"/home/z/my-project/scripts/anteprima_logo_{nome}.png")
print("Anteprime salvate in scripts/ (anteprima_logo_blu.png / anteprima_logo_bianco.png)")
