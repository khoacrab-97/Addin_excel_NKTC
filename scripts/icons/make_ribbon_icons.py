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


def icon_export():
    image, d = canvas()
    grid(d, (30, 40, 200, 280))
    arrow_right(d, 150, 200, 150)
    return image


def icon_table():
    image, d = canvas()
    grid(d, (20, 50, 250, 290), color=GREEN, rows=4, cols=3)
    d.rectangle((20, 50, 250, 110), fill=GREEN)
    d.ellipse((190, 10, 310, 130), fill=ORANGE)
    d.rectangle((242, 30, 258, 110), fill=WHITE)
    d.rectangle((210, 62, 290, 78), fill=WHITE)
    return image


def icon_template():
    image, d = canvas()
    d.rectangle((50, 20, 250, 300), fill=WHITE, outline=BLUE, width=16)
    d.rectangle((50, 20, 250, 80), fill=BLUE)
    for y in (120, 170, 220):
        d.line([(85, y), (215, y)], fill=DARK, width=14)
    d.ellipse((190, 190, 310, 310), fill=ORANGE)
    d.rectangle((242, 210, 258, 290), fill=WHITE)
    d.rectangle((210, 242, 290, 258), fill=WHITE)
    return image


def icon_assign():
    image, d = canvas()
    d.rectangle((40, 160, 200, 280), outline=GREEN, width=16)
    d.rectangle((60, 40, 280, 110), fill=BLUE)
    d.line([(170, 110), (120, 200)], fill=ORANGE, width=26)
    d.polygon([(95, 235), (90, 170), (150, 195)], fill=ORANGE)
    return image


# Bốn nút chuyển ngày dùng chung một tam giác (cùng cỡ, cùng tâm dọc); Ngày đầu / Ngày cuối thêm vạch đứng.
# Mỗi icon căn giữa theo chiều ngang để bốn nút trên Ribbon cân đối, ngang hàng.
NAV_HALF = 85  # nửa chiều cao tam giác
NAV_TRI = int(NAV_HALF * 1.6)  # chiều rộng tam giác
NAV_BAR = 32  # chiều rộng vạch đứng
NAV_GAP = 14
NAV_Y = S // 2


def nav_icon(direction, bar):
    image, d = canvas()
    width = NAV_TRI + (NAV_BAR + NAV_GAP if bar else 0)
    left = (S - width) // 2
    top, bottom = NAV_Y - NAV_HALF, NAV_Y + NAV_HALF
    if direction == "left":
        if bar:
            d.rectangle((left, top, left + NAV_BAR - 1, bottom), fill=BLUE)
            left += NAV_BAR + NAV_GAP
        d.polygon([(left + NAV_TRI, top), (left + NAV_TRI, bottom), (left, NAV_Y)], fill=BLUE)
    else:
        d.polygon([(left, top), (left, bottom), (left + NAV_TRI, NAV_Y)], fill=BLUE)
        if bar:
            x = left + NAV_TRI + NAV_GAP
            d.rectangle((x, top, x + NAV_BAR - 1, bottom), fill=BLUE)
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
    "table": icon_table,
    "export": icon_export,
    "template": icon_template,
    "assign": icon_assign,
    "first": lambda: nav_icon("left", bar=True),
    "prev": lambda: nav_icon("left", bar=False),
    "next": lambda: nav_icon("right", bar=False),
    "last": lambda: nav_icon("right", bar=True),
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
