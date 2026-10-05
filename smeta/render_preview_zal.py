#!/usr/bin/env python3
"""Превью сметы Симагино · Зал."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from build_smeta_zal import SECTIONS, TARGET, money, rub_words, round_rub

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
ROW = (247, 243, 234)
LINE = (215, 205, 189)
SUB = (231, 238, 246)


def font(path: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(path, size)


def wrap(draw: ImageDraw.ImageDraw, text: str, face: ImageFont.FreeTypeFont, width: int) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        trial = word if not current else f"{current} {word}"
        if draw.textlength(trial, font=face) <= width:
            current = trial
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines or [""]


def draw_cover() -> Path:
    width, height = 1400, 820
    image = Image.new("RGB", (width, height), CREAM)
    draw = ImageDraw.Draw(image)
    draw.rectangle([0, 0, width, 128], fill=NAVY_DEEP)
    draw.rectangle([0, 128, width, 138], fill=GOLD)
    draw.text((56, 28), "ЛОКАЛЬНАЯ СМЕТА   ·   ТОЛЬКО РАБОТЫ", font=font(FONT_BOLD, 22), fill=GOLD)
    draw.text((56, 64), "Объект Симагино  ·  Зал", font=font(FONT_BOLD, 38), fill=WHITE)

    draw.text((56, 162), "Грунтовка стен  ·  поклейка обоев  ·  покраска", font=font(FONT_REG, 20), fill=STEEL)
    draw.text((56, 198), "ЛС-СИМ-02 от 05.10.2026   ·   материалы в смету не входят", font=font(FONT_ITAL, 16), fill=MUTED)

    draw.rounded_rectangle([56, 250, 1344, 420], radius=16, fill=NAVY)
    draw.text((80, 268), "ИТОГО РАБОТЫ", font=font(FONT_BOLD, 16), fill=GOLD)
    draw.text((80, 298), f"{money(TARGET)} ₽", font=font(FONT_BOLD, 64), fill=WHITE)
    draw.text((80, 372), rub_words(TARGET), font=font(FONT_ITAL, 18), fill=GOLD_SOFT)

    items = SECTIONS[0]["items"]
    gap = 18
    card_w = (1344 - 56 - gap * 2) // 3
    y1, y2 = 460, 640
    for i, item in enumerate(items):
        x = 56 + i * (card_w + gap)
        amount = round_rub(item["qty"] * item["price"])
        draw.rounded_rectangle([x, y1, x + card_w, y2], radius=12, fill=WHITE, outline=LINE, width=2)
        draw.rectangle([x, y1, x + 8, y2], fill=GOLD)
        draw.text((x + 24, y1 + 28), item["code"], font=font(FONT_BOLD, 16), fill=GOLD)
        lines = wrap(draw, item["name"], font(FONT_BOLD, 20), card_w - 40)
        yy = y1 + 60
        for line in lines:
            draw.text((x + 24, yy), line, font=font(FONT_BOLD, 20), fill=NAVY)
            yy += 28
        draw.text((x + 24, y2 - 70), f"{item['qty']} {item['unit']} × {money(item['price'])}", font=font(FONT_REG, 15), fill=MUTED)
        draw.text((x + 24, y2 - 42), f"{money(amount)} ₽", font=font(FONT_BOLD, 26), fill=INK)

    draw.text((56, 680), "В смету входят только эти три работы. Шпаклёвка, демонтаж старых обоев и материалы — отдельно.", font=font(FONT_REG, 16), fill=MUTED)

    path = OUT_DIR / "smeta_simagino_zal_titul.png"
    image.save(path, "PNG")
    return path


def draw_table() -> Path:
    width = 1600
    row_h = 56
    items = SECTIONS[0]["items"]
    height = 160 + (len(items) + 3) * row_h + 90
    image = Image.new("RGB", (width, height), CREAM)
    draw = ImageDraw.Draw(image)

    draw.rectangle([0, 0, width, 86], fill=NAVY_DEEP)
    draw.rectangle([0, 86, width, 94], fill=GOLD)
    draw.text((32, 16), "СМЕТА НА РАБОТЫ  ·  СИМАГИНО  ·  ЗАЛ  ·  ЛС-СИМ-02", font=font(FONT_BOLD, 22), fill=GOLD)
    draw.text((32, 48), "Только работы. Количество договорное, под итог 50 000 ₽.", font=font(FONT_REG, 16), fill=GOLD_SOFT)

    columns = [32, 110, 520, 620, 760, 940, 1120, 1570]
    headers = ["№", "Наименование", "Ед.", "Кол-во", "Расценка", "Сумма", "Примечание"]
    y = 110
    draw.rectangle([24, y, width - 24, y + 36], fill=NAVY)
    for x, label in zip(columns, headers):
        draw.text((x + 6, y + 8), label, font=font(FONT_BOLD, 15), fill=WHITE)
    y += 36

    draw.rectangle([24, y, width - 24, y + row_h], fill=STEEL)
    draw.text((40, y + 18), "1.  Отделка зала", font=font(FONT_BOLD, 18), fill=WHITE)
    y += row_h

    small = font(FONT_REG, 16)
    small_bold = font(FONT_BOLD, 16)
    note_font = font(FONT_REG, 14)
    for i, item in enumerate(items):
        bg = WHITE if i % 2 == 0 else ROW
        draw.rectangle([24, y, width - 24, y + row_h], fill=bg)
        draw.line([24, y + row_h, width - 24, y + row_h], fill=LINE, width=1)
        amount = round_rub(item["qty"] * item["price"])
        values = [item["code"], item["name"], item["unit"], str(item["qty"]), money(item["price"]), money(amount)]
        faces = [small, small_bold, small, small, small, small_bold]
        for x, value, face in zip(columns, values, faces):
            draw.text((x + 6, y + 18), value, font=face, fill=INK)
        note_lines = wrap(draw, item["note"], note_font, columns[7] - columns[6] - 16)[:2]
        note_y = y + 10 if len(note_lines) > 1 else y + 18
        for line in note_lines:
            draw.text((columns[6] + 6, note_y), line, font=note_font, fill=MUTED)
            note_y += 18
        y += row_h

    draw.rectangle([24, y, width - 24, y + row_h], fill=SUB)
    draw.text((120, y + 18), "Итого по разделу 1", font=font(FONT_BOLD, 16), fill=NAVY)
    draw.text((columns[5] + 6, y + 18), money(TARGET), font=font(FONT_BOLD, 18), fill=NAVY)
    y += row_h

    draw.rectangle([24, y, width - 24, y + 56], fill=NAVY)
    draw.text((120, y + 16), "ИТОГО РАБОТЫ", font=font(FONT_BOLD, 20), fill=WHITE)
    draw.text((columns[5] + 6, y + 14), f"{money(TARGET)} ₽", font=font(FONT_BOLD, 22), fill=GOLD)
    y += 68
    draw.text((32, y), rub_words(TARGET), font=font(FONT_ITAL, 16), fill=STEEL)

    path = OUT_DIR / "smeta_simagino_zal_smeta.png"
    image.save(path, "PNG")
    return path


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    if sum(round_rub(i["qty"] * i["price"]) for i in SECTIONS[0]["items"]) != TARGET:
        raise SystemExit("Превью разошлось с итогом")
    print("COVER", draw_cover())
    print("TABLE", draw_table())


if __name__ == "__main__":
    main()
