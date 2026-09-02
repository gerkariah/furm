#!/usr/bin/env python3
"""Draws a FURM newsletter/LinkedIn-style banner PNG (forest/earthy theme)."""
from PIL import Image, ImageDraw, ImageFont
import math, os

W, H = 1584, 396
BARK  = (61,43,31)
CREAM = (246,239,227)
CREAM2= (251,246,236)
PINE  = (47,93,67)
PINE2 = (63,122,88)
MOSS  = (111,156,106)
SUN   = (230,178,62)
SUN2  = (242,199,102)

img = Image.new("RGB",(W,H),CREAM)
d = ImageDraw.Draw(img)

# vertical warm gradient (cream -> soft)
for y in range(H):
    t = y/H
    r = int(246 - 14*t); g = int(239 - 16*t); b = int(227 - 20*t)
    d.line([(0,y),(W,y)], fill=(r,g,b))

# left pine panel
panel_w = 470
d.rectangle([0,0,panel_w,H], fill=PINE)
d.rectangle([panel_w,0,panel_w+10,H], fill=SUN)

# heartbeat line across the cream area (sits low, below the text block)
base_y = H - 46
pts=[]
x=panel_w+40
while x < W-40:
    pts.append((x,base_y)); x+=6
# inject a beat around 2 spots
def beat(cx):
    return [(cx,base_y),(cx+10,base_y-38),(cx+22,base_y+30),(cx+34,base_y-12),(cx+46,base_y)]
line=[]
x=panel_w+40
while x < W-60:
    if abs(x-(panel_w+320))<3 or abs(x-(W-360))<3:
        line += beat(x); x+=60
    else:
        line.append((x,base_y)); x+=6
d.line(line, fill=MOSS, width=4)

# logo on the pine panel (circular)
src = "../assets/logo_circle.png" if os.path.exists("../assets/logo_circle.png") else "../assets/logo_padded.png"
if os.path.exists(src):
    lg = Image.open(src).convert("RGBA")
    d_=300
    lg = lg.resize((d_,d_))
    # cream disc behind
    disc = Image.new("RGBA",(d_+28,d_+28),(0,0,0,0))
    dd = ImageDraw.Draw(disc)
    dd.ellipse([0,0,d_+28,d_+28], fill=CREAM2, outline=SUN, width=6)
    img.paste(disc,(panel_w//2-(d_+28)//2, H//2-(d_+28)//2), disc)
    img.paste(lg,(panel_w//2-d_//2, H//2-d_//2), lg)

def font(sz, bold=False):
    paths = [
        "/System/Library/Fonts/Supplemental/Georgia Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Georgia.ttf",
        "/Library/Fonts/Georgia.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
    ]
    for p in paths:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, sz)
            except: pass
    return ImageFont.load_default()

tx = panel_w + 60
d.text((tx, 92), "FURM", font=font(96,True), fill=PINE)
d.text((tx, 200), "First-Generation Underrepresented", font=font(34,True), fill=BARK)
d.text((tx, 244), "in Research & Medicine  \u00b7  Brown PLME", font=font(30,False), fill=(90,66,48))
d.text((tx, 300), "Community \u00b7 Care \u00b7 Belonging", font=font(26,False), fill=PINE2)

img.save("../assets/newsletter_banner.png")
# also drop a copy into the FURM source folder
try:
    img.save("../../FURM/newsletter_banner.png")
except Exception as e:
    print("copy warn", e)
print("wrote ../assets/newsletter_banner.png", img.size)
