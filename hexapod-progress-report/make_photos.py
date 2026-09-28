"""Resize the team's photographs and build the paired close-up panels used in Chapter 7.
Source photos go in photos/src/ as 1.jpg ... 5.jpg."""
import os
from PIL import Image, ImageDraw, ImageFont, ImageOps
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "photos", "src"); OUT = os.path.join(HERE, "figures")
FONT = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 44)

def load(n):
    return ImageOps.exif_transpose(Image.open(os.path.join(SRC, f"{n}.jpg"))).convert("RGB")

def pair(files, labels, name, h=1400, gap=40, pad=90):
    ims = [load(f) for f in files]
    ims = [im.resize((round(im.width * h / im.height), h), Image.LANCZOS) for im in ims]
    W = sum(im.width for im in ims) + gap * (len(ims) - 1)
    canvas = Image.new("RGB", (W, h + pad), "white"); d = ImageDraw.Draw(canvas); x = 0
    for im, lab in zip(ims, labels):
        canvas.paste(im, (x, 0))
        tw = d.textlength(lab, font=FONT); d.text((x + (im.width - tw) / 2, h + 22), lab, fill="black", font=FONT)
        x += im.width + gap
    canvas.save(os.path.join(OUT, name), quality=85)

big = load(1); big.thumbnail((1800, 1800)); big.save(os.path.join(OUT, "photo_components.jpg"), quality=85)
pair([2, 3], ["(a) Arduino Uno R3", "(b) ESP32-WROOM-32 dev board"], "photo_uno_esp32.jpg")
pair([4, 5], ["(c) MG996R servo (TowerPro)", "(d) PCA9685"], "photo_servo_pca9685.jpg")
print("ok")
