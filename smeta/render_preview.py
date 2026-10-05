#!/usr/bin/env python3
"""Превью сметы Симагино для проверки. Не заменяет Excel."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from build_smeta import SECTIONS, TARGET, money, rub_words, section_sum

OUT_DIR = Path("/opt/cursor/artifacts")
FONT_REG = "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf"
FONT_ITAL = "/usr/share/fonts/truetype/noto/NotoSans-Italic.ttf"

NAVY = (27, 54, 93)
NAVY_DEEP = (18, 36, 63)
GOLD = (196, 163, 90)
GOLD_SOFT = (230, 211, 163)
CREAM = (251, 248, 242)
WHITE = (255, 255, 255)
INK = (31, 41, 51)
MUTED = (92, 102, 112)
STEEL = (61, 90, 128)
SAND = (243, 235, 221)
ROW = (247, 243, 234)
LINE = (215, 205, 189)
SUB = (231, 238, 246)


def font(path: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(path, size)


def text_width(draw: ImageDraw.ImageDraw, text: str, face: ImageFont.FreeTypeFont) -> int:
    return int(draw.textlength(text, font=face))


def wrap(draw: ImageDraw.ImageDraw, text: str, face: ImageFont.FreeTypeFont, width: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        trial = word if not current else f"{current} {word}"
        if text_width(draw, trial, face) <= width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]


def draw_cover() -> Path:
    width, height = 1400, 980
    image = Image.new("RGB", (width, height), CREAM)
    draw = ImageDraw.Draw(image)
    draw.rectangle([0, 0, width, 128], fill=NAVY_DEEP)
    draw.rectangle([0, 128, width, 138], fill=GOLD)
    draw.text((56, 28), "ЛОКАЛЬНАЯ СМЕТА   ·   ТОЛЬКО РАБОТЫ", font=font(FONT_BOLD, 22), fill=GOLD)
    draw.text((56, 64), "Объект Симагино", font=font(FONT_BOLD, 40), fill=WHITE)

    draw.text(
        (56, 162),
        "Электрика  ·  стены 1 этажа  ·  откосы и подоконники  ·  короб над печкой  ·  котельная",
        font=font(FONT_REG, 18),
        fill=STEEL,
    )
    draw.text(
        (56, 196),
        "ЛС-СИМ-01 от 05.10.2026   ·   материалы и оборудование в смету не входят",
        font=font(FONT_ITAL, 16),
        fill=MUTED,
    )

    draw.rounded_rectangle([56, 250, 1344, 430], radius=16, fill=NAVY)
    draw.text((80, 268), "ИТОГО РАБОТЫ", font=font(FONT_BOLD, 16), fill=GOLD)
    draw.text((80, 300), f"{money(TARGET)} ₽", font=font(FONT_BOLD, 64), fill=WHITE)
    draw.text((80, 382), rub_words(TARGET), font=font(FONT_ITAL, 18), fill=GOLD_SOFT)

    cards = []
    for section in SECTIONS:
        cards.append((f"{section['num']}. {section['title']}", section_sum(section)))
    gap = 16
    card_w = (1344 - 56 - gap * 3) // 4
    y1, y2 = 460, 620
    for i, (title, amount) in enumerate(cards):
        x = 56 + i * (card_w + gap)
        draw.rounded_rectangle([x, y1, x + card_w, y2], radius=12, fill=WHITE, outline=LINE, width=2)
        draw.rectangle([x, y1, x + 8, y2], fill=GOLD)
        lines = wrap(draw, title, font(FONT_BOLD, 16), card_w - 36)
        yy = y1 + 24
        for line in lines:
            draw.text((x + 22, yy), line, font=font(FONT_BOLD, 16), fill=NAVY)
            yy += 22
        draw.text((x + 22, y2 - 58), f"{money(amount)} ₽", font=font(FONT_BOLD, 26), fill=INK)

    draw.text((56, 660), "В смете только эти работы", font=font(FONT_BOLD, 20), fill=NAVY)
    names = [item["name"] for section in SECTIONS for item in section["items"]]
    left = names[:7]
    right = names[7:]
    y = 700
    for i, name in enumerate(left):
        draw.text((56, y), f"{i + 1}.  {name}", font=font(FONT_REG, 18), fill=INK)
        y += 30
    y = 700
    for i, name in enumerate(right):
        draw.text((720, y), f"{i + 8}.  {name}", font=font(FONT_REG, 18), fill=INK)
        y += 30

    path = OUT_DIR / "smeta_simagino_titul.png"
    image.save(path, "PNG")
    return path


def draw_table() -> Path:
    width = 1760
    row_h = 52
    header_h = 160
    footer_h = 100
    item_count = sum(len(section["items"]) for section in SECTIONS)
    section_extra = len(SECTIONS) * 2  # title + subtotal
    height = header_h + (item_count + section_extra + 1) * row_h + footer_h
    image = Image.new("RGB", (width, height), CREAM)
    draw = ImageDraw.Draw(image)

    draw.rectangle([0, 0, width, 86], fill=NAVY_DEEP)
    draw.rectangle([0, 86, width, 94], fill=GOLD)
    draw.text((32, 16), "СМЕТА НА РАБОТЫ  ·  СИМАГИНО  ·  ЛС-СИМ-01", font=font(FONT_BOLD, 22), fill=GOLD)
    draw.text((32, 48), "Только работы. Количество договорное, под итог 230 000 ₽.", font=font(FONT_REG, 16), fill=GOLD_SOFT)

    columns = [32, 110, 700, 790, 910, 1070, 1240, 1730]
    headers = ["№", "Наименование", "Ед.", "Кол-во", "Расценка", "Сумма", "Примечание"]
    y = 110
    draw.rectangle([24, y, width - 24, y + 36], fill=NAVY)
    for x, label, x2 in zip(columns, headers, columns[1:]):
        draw.text((x + 6, y + 8), label, font=font(FONT_BOLD, 15), fill=WHITE)
    y += 36

    small = font(FONT_REG, 15)
    small_bold = font(FONT_BOLD, 15)
    note_font = font(FONT_REG, 13)
    for section in SECTIONS:
        draw.rectangle([24, y, width - 24, y + row_h], fill=STEEL)
        draw.text((40, y + 16), f"{section['num']}.  {section['title']}", font=font(FONT_BOLD, 16), fill=WHITE)
        y += row_h
        for i, item in enumerate(section["items"]):
            bg = WHITE if i % 2 == 0 else ROW
            draw.rectangle([24, y, width - 24, y + row_h], fill=bg)
            draw.line([24, y + row_h, width - 24, y + row_h], fill=LINE, width=1)
            amount = item["qty"] * item["price"]
            values = [
                item["code"],
                item["name"],
                item["unit"],
                str(item["qty"]),
                money(item["price"]),
                money(amount),
            ]
            faces = [small, small_bold, small, small, small, small_bold]
            for x, value, face in zip(columns, values, faces):
                draw.text((x + 6, y + 16), value, font=face, fill=INK)
            note_lines = wrap(draw, item["note"], note_font, columns[7] - columns[6] - 16)[:2]
            note_y = y + 8 if len(note_lines) > 1 else y + 16
            for line in note_lines:
                draw.text((columns[6] + 6, note_y), line, font=note_font, fill=MUTED)
                note_y += 18
            y += row_h
        sub = section_sum(section)
        draw.rectangle([24, y, width - 24, y + row_h], fill=SUB)
        draw.text((120, y + 16), f"Итого по разделу {section['num']}", font=font(FONT_BOLD, 15), fill=NAVY)
        draw.text((columns[5] + 6, y + 16), money(sub), font=font(FONT_BOLD, 16), fill=NAVY)
        y += row_h

    draw.rectangle([24, y, width - 24, y + 56], fill=NAVY)
    draw.text((120, y + 16), "ИТОГО РАБОТЫ", font=font(FONT_BOLD, 20), fill=WHITE)
    draw.text((columns[5] + 6, y + 14), f"{money(TARGET)} ₽", font=font(FONT_BOLD, 22), fill=GOLD)
    y += 64
    draw.text((32, y), rub_words(TARGET), font=font(FONT_ITAL, 16), fill=STEEL)

    path = OUT_DIR / "smeta_simagino_smeta.png"
    image.save(path, "PNG")
    return path


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    cover = draw_cover()
    table = draw_table()
    if sum(section_sum(section) for section in SECTIONS) != TARGET:
        raise SystemExit("Превью разошлось с итогом")
    print("COVER", cover)
    print("TABLE", table)


if __name__ == "__main__":
    main()
