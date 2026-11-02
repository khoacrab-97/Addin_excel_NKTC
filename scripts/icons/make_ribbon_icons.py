"""Vẽ icon cho các nút trên tab Ribbon "Nhật ký thi công" (16/32/80 px, PNG nền trong suốt).

Chạy: python3 -I scripts/icons/make_ribbon_icons.py  → assets/ribbon/<tên>-<cỡ>.png
Vẽ ở 320 px rồi thu nhỏ để nét mịn.
"""
import os
import sys

from PIL import Image, ImageDraw

BLUE = (15, 108, 189, 255)
DARK = (32, 31, 30, 255)
GREEN = (16, 124, 65, 255)
ORANGE = (202, 80, 16, 255)
WHITE = (255, 255, 255, 255)
S = 320
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "assets", "ribbon")


def canvas():
    image = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    return image, ImageDraw.Draw(image)


def grid(d, box, color=GREEN, rows=4, cols=3, width=14):
    x1, y1, x2, y2 = box
    d.rectangle(box, outline=color, width=width)
    for i in range(1, rows):
        y = y1 + (y2 - y1) * i // rows
        d.line([(x1, y), (x2, y)], fill=color, width=width // 2)
    for j in range(1, cols):
        x = x1 + (x2 - x1) * j // cols
        d.line([(x, y1), (x, y2)], fill=color, width=width // 2)


def arrow_right(d, x, y, length, color=BLUE, width=30):
    d.line([(x, y), (x + length - 50, y)], fill=color, width=width)
    d.polygon([(x + length, y), (x + length - 70, y - 55), (x + length - 70, y + 55)], fill=color)


def triangle(d, x, y, size, direction, color=BLUE):
    if direction == "right":
        d.polygon([(x, y - size), (x, y + size), (x + size * 1.6, y)], fill=color)
    else:
        d.polygon([(x, y - size), (x, y + size), (x - size * 1.6, y)], fill=color)


def icon_export():
    image, d = canvas()
    grid(d, (30, 40, 200, 280))
    arrow_right(d, 150, 200, 150)
    return image


def icon_source():
    image, d = canvas()
    grid(d, (30, 40, 290, 280), rows=5, cols=4)
    d.rectangle((30, 40, 290, 88), fill=GREEN)
    d.rectangle((100, 140, 290, 280), outline=BLUE, width=22)
    return image


def icon_assign():
    image, d = canvas()
    d.rectangle((40, 160, 200, 280), outline=GREEN, width=16)
    d.rectangle((60, 40, 280, 110), fill=BLUE)
    d.line([(170, 110), (120, 200)], fill=ORANGE, width=26)
    d.polygon([(95, 235), (90, 170), (150, 195)], fill=ORANGE)
    return image


def icon_list():
    image, d = canvas()
    for i, y in enumerate((70, 160, 250)):
        d.rectangle((30, y - 22, 74, y + 22), fill=BLUE)
        d.line([(100, y), (290, y)], fill=DARK, width=24)
    return image


def icon_first():
    image, d = canvas()
    d.rectangle((60, 80, 90, 240), fill=BLUE)
    triangle(d, 250, 160, 80, "left")
    return image


def icon_prev():
    image, d = canvas()
    triangle(d, 230, 160, 90, "left")
    return image


def icon_next():
    image, d = canvas()
    triangle(d, 90, 160, 90, "right")
    return image


def icon_last():
    image, d = canvas()
    triangle(d, 70, 160, 80, "right")
    d.rectangle((230, 80, 260, 240), fill=BLUE)
    return image


def icon_goto():
    image, d = canvas()
    d.rounded_rectangle((30, 50, 290, 290), radius=24, outline=BLUE, width=18)
    d.rectangle((30, 50, 290, 110), fill=BLUE)
    for x in (90, 230):
        d.rectangle((x - 10, 25, x + 10, 75), fill=DARK)
    for r in range(2):
        for c in range(4):
            x = 60 + c * 58
            y = 140 + r * 65
            d.rectangle((x, y, x + 36, y + 36), fill=ORANGE if (r, c) == (1, 2) else (150, 150, 150, 255))
    return image


def icon_circle_letter(letter):
    image, d = canvas()
    d.ellipse((20, 20, 300, 300), fill=BLUE)
    if letter == "?":
        d.arc((105, 70, 215, 180), start=180, end=90, fill=WHITE, width=30)
        d.line([(160, 180), (160, 215)], fill=WHITE, width=30)
        d.ellipse((143, 235, 177, 269), fill=WHITE)
    else:
        d.ellipse((143, 65, 177, 99), fill=WHITE)
        d.rectangle((145, 125, 175, 255), fill=WHITE)
    return image


ICONS = {
    "export": icon_export,
    "source": icon_source,
    "assign": icon_assign,
    "list": icon_list,
    "first": icon_first,
    "prev": icon_prev,
    "next": icon_next,
    "last": icon_last,
    "goto": icon_goto,
    "help": lambda: icon_circle_letter("?"),
    "about": lambda: icon_circle_letter("i"),
}


def main():
    os.makedirs(OUT, exist_ok=True)
    for name, draw in ICONS.items():
        image = draw()
        for size in (16, 32, 80):
            image.resize((size, size), Image.LANCZOS).save(os.path.join(OUT, f"{name}-{size}.png"))
    print(f"{len(ICONS)} icon x 3 cỡ → {os.path.normpath(OUT)}")


if __name__ == "__main__":
    sys.exit(main())
