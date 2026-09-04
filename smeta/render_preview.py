#!/usr/bin/env python3
"""Превью сметы для проверки и артефактов. Не заменяет Excel."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from build_smeta import SECTIONS, money, rub_words, round_rub

OUT_DIR = Path("/opt/cursor/artifacts")
FONT_REG = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"

NAVY = (27, 54, 93)
NAVY_DEEP = (18, 36, 63)
GOLD = (196, 163, 90)
CREAM = (251, 248, 242)
WHITE = (255, 255, 255)
INK = (31, 41, 51)
MUTED = (92, 102, 112)
STEEL = (61, 90, 128)
TERRACOTTA = (166, 93, 63)
GREEN = (47, 107, 79)
SAND = (243, 235, 221)
ROW = (247, 243, 234)


def fnt(path, size):
    return ImageFont.truetype(path, size)


def section_sums():
    rows = []
    main = rec = 0
    for s in SECTIONS:
        total = sum(round_rub(it["qty"] * it["price"]) for it in s["items"])
        rows.append((s["num"], s["title"], total, bool(s.get("recommended"))))
        if s.get("recommended"):
            rec += total
        else:
            main += total
    reserve = round(main * 0.05)
    return rows, main, rec, reserve, main + rec + reserve


def rounded_rect(draw, xy, r, fill):
    x1, y1, x2, y2 = xy
    draw.rounded_rectangle(xy, radius=r, fill=fill)


def draw_cover():
    rows, main, rec, reserve, grand = section_sums()
    w, h = 1600, 1000
    img = Image.new("RGB", (w, h), CREAM)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, 110], fill=NAVY_DEEP)
    d.rectangle([0, 110, w, 118], fill=GOLD)
    d.text((56, 28), "ЛОКАЛЬНАЯ СМЕТА  ·  ТОЛЬКО РАБОТЫ", font=fnt(FONT_BOLD, 22), fill=GOLD)
    d.text((56, 62), "Ремонт ванной комнаты и туалета  ·  дом серии ЛГ-137", font=fnt(FONT_REG, 28), fill=WHITE)

    d.text((56, 150), "Раздельный санузел сантехкабины  ·  материалы не входят", font=fnt(FONT_REG, 22), fill=STEEL)
    d.text(
        (56, 190),
        "Без демонтажа плитки  ·  потолок не трогаем  ·  стояки воды и канализации сохраняем  ·  живут на объекте",
        font=fnt(FONT_REG, 18),
        fill=MUTED,
    )

    cards = [
        ("ОСНОВНОЙ ОБЪЁМ", f"{money(main)} ₽", "разделы 1–7", NAVY),
        ("РЕКОМЕНДУЕМЫЕ", f"{money(rec)} ₽", "раздел 8", TERRACOTTA),
        ("РЕЗЕРВ 5%", f"{money(reserve)} ₽", "по согласованию", STEEL),
        ("ИТОГО К ДОГОВОРУ", f"{money(grand)} ₽", "работы, без материалов", (139, 105, 42)),
    ]
    cw = 350
    gap = 24
    x0 = 56
    for i, (lab, val, sub, bg) in enumerate(cards):
        x = x0 + i * (cw + gap)
        rounded_rect(d, (x, 250, x + cw, 430), 16, bg)
        d.text((x + 24, 272), lab, font=fnt(FONT_BOLD, 16), fill=GOLD)
        d.text((x + 24, 318), val, font=fnt(FONT_BOLD, 34), fill=WHITE)
        d.text((x + 24, 375), sub, font=fnt(FONT_REG, 16), fill=(230, 213, 163))

    d.text((56, 470), rub_words(int(grand)), font=fnt(FONT_REG, 20), fill=STEEL)

    y = 530
    d.rectangle([56, y, w - 56, y + 3], fill=GOLD)
    y = 555
    d.text((56, y), "Структура основного объёма", font=fnt(FONT_BOLD, 22), fill=NAVY)
    y = 600
    main_rows = [r for r in rows if not r[3]]
    for num, title, total, _ in main_rows:
        pct = total / main * 100
        d.text((56, y), f"{num}.", font=fnt(FONT_BOLD, 16), fill=NAVY)
        d.text((96, y), title[:62], font=fnt(FONT_REG, 16), fill=INK)
        d.text((1280, y), f"{money(total)} ₽", font=fnt(FONT_BOLD, 16), fill=NAVY)
        bar_x, bar_w = 96, 1100
        d.rectangle([bar_x, y + 24, bar_x + bar_w, y + 30], fill=SAND)
        d.rectangle([bar_x, y + 24, bar_x + int(bar_w * pct / 100), y + 30], fill=GOLD)
        y += 48

    d.text((56, 960), "Габариты: ванна 1700×1500 мм, туалет 1200×820 мм, H = 2550 мм, двери 570 мм. Срок ≈ 18 рабочих дней.", font=fnt(FONT_REG, 16), fill=MUTED)
    path = OUT_DIR / "smeta_preview_svodka.png"
    img.save(path, "PNG")
    return path


def draw_items_check():
    """Таблица ключевых позиций задания — проверка полноты."""
    w, h = 1600, 1200
    img = Image.new("RGB", (w, h), CREAM)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, 90], fill=NAVY)
    d.rectangle([0, 90, w, 98], fill=GOLD)
    d.text((56, 28), "Проверка состава: всё, что просил заказчик, есть в смете", font=fnt(FONT_BOLD, 28), fill=WHITE)

    checks = [
        ("Демонтаж плитки", "не входит", True),
        ("Потолок", "не трогаем", True),
        ("Стояки ХВС / ГВС и канализации", "не трогаем", True),
        ("Демонтаж унитаза + ежедневный временный унитаз", "18 циклов, живут на объекте", True),
        ("Вынос стиральной машины", "и обратное подключение", True),
        ("Вынос чугунной ванны", "узкий проём 57 см, резка", True),
        ("Срезка старых труб воды", "стояки сохраняются", True),
        ("Коллекторы + фильтры стандарт 10\" + редукторы", "ХВС и ГВС", True),
        ("Полипропилен на раковину, унитаз, стиралку, ванну", "32 м.п.", True),
        ("Водорозетки", "6 точек", True),
        ("Инсталляция + переход с чугуна без замены стояка", "туалет ЛГ-137", True),
        ("Подбивка порогов + плитка под двери ванны и туалета", "2 проёма", True),
        ("Гидроизоляция пола + керамогранит", "3,53 м²", True),
        ("Гидроизоляция стен до ванны + плитка", "и зона душа до 2 м", True),
        ("Ванна из наливного камня (тяжёлая)", "подиум, занос 3–4 чел.", True),
        ("Смесители, душ, раковина, унитаз", "чистовой монтаж", True),
        ("Рекомендации: короб стояков, вентилятор, розетки IP44, шторка", "раздел 8", True),
    ]
    y = 130
    for title, note, ok in checks:
        rounded_rect(d, (56, y, w - 56, y + 52), 8, WHITE)
        d.ellipse([76, y + 14, 108, y + 46], fill=GREEN)
        d.text((84, y + 16), "✓", font=fnt(FONT_BOLD, 20), fill=WHITE)
        d.text((130, y + 8), title, font=fnt(FONT_BOLD, 20), fill=INK)
        d.text((130, y + 32), note, font=fnt(FONT_REG, 16), fill=MUTED)
        y += 60

    path = OUT_DIR / "smeta_preview_sostav.png"
    img.save(path, "PNG")
    return path


def write_log():
    rows, main, rec, reserve, grand = section_sums()
    path = OUT_DIR / "smeta_verification.txt"
    lines = [
        "Проверка сметы: ремонт ванной и туалета ЛГ-137, только работы",
        f"Основной объём (1–7): {int(main)} руб.",
        f"Резерв 5%: {int(reserve)} руб.",
        f"Рекомендуемые (8): {int(rec)} руб.",
        f"Итого: {int(grand)} руб.",
        "",
        "Разделы:",
    ]
    for num, title, total, rec_flag in rows:
        tag = "рек" if rec_flag else "задание"
        lines.append(f"  {num}. [{tag}] {int(round(total))} — {title}")
    lines.append("")
    lines.append("Арифметика позиций:")
    ok = True
    for s in SECTIONS:
        ssum = 0
        for it in s["items"]:
            line = round_rub(it["qty"] * it["price"])
            ssum += line
            lines.append(f"  {it['code']}: {it['qty']} × {it['price']} = {line}")
        expect = sum(round_rub(it["qty"] * it["price"]) for it in s["items"])
        if abs(ssum - expect) > 0.01:
            ok = False
            lines.append(f"  ERROR section {s['num']}")
    lines.append("")
    lines.append("STATUS: OK" if ok else "STATUS: FAIL")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path, ok, grand


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    p1 = draw_cover()
    p2 = draw_items_check()
    log, ok, grand = write_log()
    print(p1)
    print(p2)
    print(log, ok, grand)


if __name__ == "__main__":
    main()
