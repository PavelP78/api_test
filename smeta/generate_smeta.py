#!/usr/bin/env python3
"""Смета ВНУТРЕННЕЙ отделки и обстановки каркасного дома под посуточную сдачу.

Исключено (уже есть): окна, терраса (настил), наружные стены/кровля/фасад.
Отделка стен/потолков: вагонка (штиль) белая — как на фото, НЕ ГКЛ.
Отопление: электрические конвекторы.
Уличное освещение: включено.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
import shutil
from pathlib import Path

# ---------------------------------------------------------------------------
# Исходные данные по чертежу
# ---------------------------------------------------------------------------
AREAS = {
    "прихожая": 2.80,
    "санузел": 3.32,
    "кухня_гостиная": 18.82,
    "спальня1": 9.21,
    "спальня2": 9.21,
}
FLOOR_LIVING = round(sum(AREAS.values()), 2)  # 43.36
CEILING_H = 2.50

PARTITIONS_L = 17.7
PARTITIONS_AREA_BOTH = round(PARTITIONS_L * CEILING_H * 2, 1)
EXT_INNER_PERIM = 2 * (5.62 + 8.72)
EXT_INNER_AREA = round(EXT_INNER_PERIM * CEILING_H, 1)
GLAZING_DED = 18.0  # окна/витражи УЖЕ ЕСТЬ — вычитаем из отделки стен
WALL_FINISH_AREA = round(PARTITIONS_AREA_BOTH + EXT_INNER_AREA - GLAZING_DED, 1)

wb = Workbook()

thin = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
header_fill = PatternFill("solid", fgColor="1F4E79")
header_font = Font(bold=True, color="FFFFFF", name="Calibri", size=11)
section_fill = PatternFill("solid", fgColor="D6EAF8")
section_font = Font(bold=True, name="Calibri", size=11, color="1F4E79")
total_fill = PatternFill("solid", fgColor="FFF2CC")
total_font = Font(bold=True, name="Calibri", size=11)
grand_fill = PatternFill("solid", fgColor="C6EFCE")
grand_font = Font(bold=True, name="Calibri", size=12)
money_font = Font(name="Calibri", size=10)
title_font = Font(bold=True, name="Calibri", size=14, color="1F4E79")
note_font = Font(name="Calibri", size=9, italic=True, color="666666")
wrap = Alignment(wrap_text=True, vertical="center")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)

HEADERS = [
    "№",
    "Название",
    "Объём",
    "Ед.",
    "Цена за ед., ₽",
    "Цена общая, ₽",
    "Время монтажа, дн",
    "Примечание",
]


def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin


def add_section(ws, row, title, cols=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row=row, column=1, value=title)
    cell.fill = section_fill
    cell.font = section_font
    cell.alignment = Alignment(vertical="center")
    for c in range(1, cols + 1):
        ws.cell(row=row, column=c).fill = section_fill
        ws.cell(row=row, column=c).border = thin
    ws.row_dimensions[row].height = 22
    return row + 1


def add_item(ws, row, num, name, qty, unit, price, days, note=""):
    total = round(qty * price, 0)
    values = [num, name, qty, unit, price, total, days, note]
    for c, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.border = thin
        cell.font = money_font
        cell.alignment = wrap if c in (2, 8) else center
        if c in (5, 6):
            cell.number_format = "#,##0"
        if c == 3:
            cell.number_format = (
                "0.00" if isinstance(qty, float) and qty != int(qty) else "0"
            )
        if c == 7:
            cell.number_format = "0.0"
    ws.row_dimensions[row].height = 32 if len(name) > 45 else 20
    return row + 1, total, days


def add_subtotal(ws, row, label, total, days, cols=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    cell = ws.cell(row=row, column=1, value=label)
    cell.fill = total_fill
    cell.font = total_font
    cell.alignment = Alignment(horizontal="right", vertical="center")
    for c in range(1, 6):
        ws.cell(row=row, column=c).fill = total_fill
        ws.cell(row=row, column=c).border = thin
        ws.cell(row=row, column=c).font = total_font
    t = ws.cell(row=row, column=6, value=total)
    t.fill = total_fill
    t.font = total_font
    t.number_format = "#,##0"
    t.border = thin
    d = ws.cell(row=row, column=7, value=round(days, 1))
    d.fill = total_fill
    d.font = total_font
    d.number_format = "0.0"
    d.border = thin
    n = ws.cell(row=row, column=8, value="")
    n.fill = total_fill
    n.border = thin
    return row + 1


def set_widths(ws):
    widths = {1: 5, 2: 52, 3: 10, 4: 8, 5: 14, 6: 14, 7: 14, 8: 38}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w


# ===================================================================
# ЛИСТ 1 — Смета
# ===================================================================
ws = wb.active
ws.title = "Смета"

ws.merge_cells("A1:H1")
ws["A1"] = (
    "СМЕТА ВНУТРЯНКИ — каркасный дом 6,12×10 м "
    f"(жилая {FLOOR_LIVING} м²) под посуточную сдачу"
)
ws["A1"].font = title_font
ws["A1"].alignment = wrap
ws.row_dimensions[1].height = 32

ws.merge_cells("A2:H2")
ws["A2"] = (
    "Отделка стен и потолков — ВАГОНКА (штиль) белая, как на фото; ГКЛ НЕ используется. "
    "Уже есть и НЕ включено: окна, витражи, настил террасы, наружный контур, кровля, фасад. "
    "Внутри — каркас перегородок. Отопление: электрические конвекторы. "
    "Уличная подсветка — включена. "
    "В смете — ТОЛЬКО материалы и комплектация (без стоимости работ). "
    "«Время монтажа» — ориентир, бригада 2 чел. Цены — средние розничные РФ, 2026 г."
)
ws["A2"].font = note_font
ws["A2"].alignment = wrap
ws.row_dimensions[2].height = 52

ws.merge_cells("A3:H3")
ws["A3"] = (
    f"Площади: прихожая {AREAS['прихожая']} · с/у {AREAS['санузел']} · "
    f"кухня-гостиная {AREAS['кухня_гостиная']} · спальня1 {AREAS['спальня1']} · "
    f"спальня2 {AREAS['спальня2']} м² | H потолка {CEILING_H} м | "
    f"перегородки L≈{PARTITIONS_L} м · отделка стен ≈{WALL_FINISH_AREA} м² | "
    f"полы/потолок {FLOOR_LIVING} м²"
)
ws["A3"].font = note_font
ws["A3"].alignment = wrap
ws.row_dimensions[3].height = 32

row = 5
for i, h in enumerate(HEADERS, 1):
    ws.cell(row=row, column=i, value=h)
style_header(ws, row, 8)
ws.row_dimensions[row].height = 28
row += 1
ws.freeze_panes = "A6"

SECTIONS = []

SECTIONS.append((
    "1. ОТДЕЛКА СТЕН ВАГОНКОЙ (как на фото, без ГКЛ)",
    [
        ("Минеральная вата 50 мм в перегородки (звукоизоляция)", 22, "м²", 280, 0.5, "перегородки ~17,7×2,5 м"),
        ("Пароизоляция / мембрана", 50, "м²", 45, 0.3, "с нахлёстом"),
        ("Брусок обрешётки 40×40 / 20×40 под вагонку (стены)", 140, "м.п.", 55, 1.0, "шаг ~400–600 мм"),
        ("Вагонка штиль хвоя 12–14 мм (стены сухие)", 130, "м²", 650, 3.0, f"≈{WALL_FINISH_AREA} м² + запас; горизонтально как на фото"),
        ("Вагонка ПВХ / панели влагостойкие (стены с/у, вне мокрой зоны плитки)", 8, "м²", 550, 0.4, "если часть стен с/у не под плиткой"),
        ("Саморезы, кляймеры, крепёж для вагонки", 1, "компл.", 8000, 0.0, ""),
        ("Грунт / антисептик по дереву", 130, "м²", 60, 0.5, ""),
        ("Краска / масло-воск укрывистое белое по вагонке (2 слоя)", 130, "м²", 280, 1.5, "цвет как на фото — белый"),
        ("Уголки / раскладки деревянные на углы и стыки", 40, "м.п.", 120, 0.4, ""),
        ("Наличники / доборы деревянные белые к проёмам", 30, "м.п.", 280, 0.5, "в тон вагонке"),
        ("Откосы внутренние на окна (вагонка / сэндвич белый)", 18, "м.п.", 450, 1.0, "окна уже стоят"),
        ("Подоконники ПВХ / дерево белые", 6, "м.п.", 1200, 0.5, "если ещё не установлены"),
    ],
))

SECTIONS.append((
    "2. ПОТОЛКИ ВАГОНКОЙ (как на фото, без ГКЛ)",
    [
        ("Брусок / рейка обрешётки потолка", 80, "м.п.", 55, 0.8, ""),
        ("Вагонка штиль хвоя 12–14 мм на потолок", 48, "м²", 650, 1.5, f"{FLOOR_LIVING} м² + запас"),
        ("Кляймеры / саморезы потолок", 1, "компл.", 3500, 0.0, ""),
        ("Грунт + краска/масло белое по потолку (2 слоя)", 48, "м²", 280, 1.0, "в тон стенам"),
        ("Галтель / плинтус потолочный деревянный", 40, "м.п.", 180, 0.5, "периметр помещений"),
    ],
))

SECTIONS.append((
    "3. ПОЛЫ",
    [
        ("Фанера / ОСБ-3 18–21 мм выравнивающая (при необходимости)", 48, "м²", 650, 1.0, "уточнить по существующему черновому полу"),
        ("Подложка под ламинат 3 мм", 40, "м²", 80, 0.2, "без с/у"),
        ("Ламинат 33 кл. влагостойкий (жилые + прихожая)", 44, "м²", 1200, 1.5, f"~{FLOOR_LIVING - AREAS['санузел']} м² + запас 10%"),
        ("Плинтус МДФ/ПВХ с кабель-каналом + уголки", 45, "м.п.", 220, 0.5, ""),
        ("Порог / переход алюминиевый", 4, "шт.", 450, 0.2, ""),
        ("Наливной пол тонкий в с/у", 4, "м²", 450, 0.5, "под плитку"),
        ("Гидроизоляция обмазочная пола с/у", 8, "м²", 350, 0.5, "пол + заход на стены 150 мм"),
    ],
))

SECTIONS.append((
    "4. САНУЗЕЛ — плитка и гидроизоляция стен",
    [
        ("Гидроизоляция стен с/у", 18, "м²", 320, 0.5, ""),
        ("Плитка настенная 30×60", 20, "м²", 1100, 1.5, "стены с/у + запас"),
        ("Керамогранит на пол 60×60", 4, "м²", 1400, 0.5, f"с/у {AREAS['санузел']} м² + запас"),
        ("Клей + затирка + крестики + уголки", 1, "компл.", 6500, 0.0, ""),
        ("Ревизионный люк", 1, "шт.", 2500, 0.2, ""),
    ],
))

SECTIONS.append((
    "5. ДВЕРИ МЕЖКОМНАТНЫЕ",
    [
        ("Дверь экошпон + коробка + наличники (ДВ-1 с/у)", 1, "компл.", 12000, 0.3, "влагостойкая, ~700 мм"),
        ("Дверь экошпон (ДВ-2 спальня 1)", 1, "компл.", 11000, 0.3, "проём ~800"),
        ("Дверь экошпон (ДВ-3 спальня 2)", 1, "компл.", 11000, 0.3, "проём ~800"),
        ("Ручки + защёлки/замки на 3 двери", 3, "компл.", 1800, 0.2, "замок на с/у"),
    ],
))

SECTIONS.append((
    "6. ЭЛЕКТРИКА ВНУТРЕННЯЯ (материалы)",
    [
        ("Кабель ВВГнг-LS 3×2,5 (розетки + конвекторы)", 140, "м", 95, 1.2, "отдельные линии на конвекторы"),
        ("Кабель ВВГнг-LS 3×1,5 (освещение)", 100, "м", 70, 0.8, ""),
        ("Кабель ВВГнг-LS 3×6 (ввод / плита)", 20, "м", 220, 0.3, ""),
        ("Гофра / клипсы / кабель-канал", 1, "компл.", 4500, 0.0, ""),
        ("Щит + автоматы + УЗО/диф. (с линиями на конвекторы)", 1, "компл.", 22000, 0.6, "отдельный автомат на каждый конвектор"),
        ("Розетки + выключатели", 38, "шт.", 280, 1.0, "в т.ч. силовые под конвекторы"),
        ("Подрозетники / распредкоробки", 50, "шт.", 45, 0.3, ""),
        ("Светильники потолочные LED", 8, "шт.", 2200, 0.5, "гостиная, спальни, прихожая"),
        ("Подсветка LED кухня + БП", 1, "компл.", 4500, 0.3, ""),
        ("Светильник IP44 (с/у)", 2, "шт.", 1800, 0.2, ""),
        ("Тёплый пол электрический мат в с/у", 3.5, "м²", 1800, 0.5, ""),
        ("Терморегулятор тёплого пола", 1, "шт.", 3500, 0.1, ""),
        ("Умный замок / электронный на вход (внутр. установка)", 1, "шт.", 15000, 0.3, "коды для гостей"),
        ("Звонок / видеодомофон простой", 1, "компл.", 8000, 0.3, ""),
        ("Кабель ВВГнг-LS 3×1,5 наружный / в гофре на уличное освещение", 40, "м", 90, 0.4, "вывод на террасу/фасад"),
        ("Светильник уличный LED настенный IP65 (терраса / вход)", 3, "шт.", 3500, 0.4, "подсветка террасы и входа"),
        ("Прожектор / грунтовый светильник LED уличный", 2, "шт.", 2800, 0.3, "подсветка участка / фасада"),
        ("Датчик движения / сумеречное реле для улицы", 1, "шт.", 1800, 0.2, "автоматическое включение"),
        ("Выключатель уличного света (внутри у двери)", 1, "шт.", 350, 0.1, ""),
    ],
))

SECTIONS.append((
    "7. ОТОПЛЕНИЕ — ЭЛЕКТРИЧЕСКИЕ КОНВЕКТОРЫ",
    [
        ("Конвектор электрический 2,0 кВт с термостатом (кухня-гостиная)", 2, "шт.", 9500, 0.4, f"комната {AREAS['кухня_гостиная']} м² — 2 шт. по периметру"),
        ("Конвектор электрический 1,5 кВт с термостатом (спальня 1)", 1, "шт.", 7500, 0.2, f"{AREAS['спальня1']} м²"),
        ("Конвектор электрический 1,5 кВт с термостатом (спальня 2)", 1, "шт.", 7500, 0.2, f"{AREAS['спальня2']} м²"),
        ("Конвектор электрический 1,0 кВт (прихожая / запас)", 1, "шт.", 6000, 0.2, f"{AREAS['прихожая']} м²"),
        ("Кронштейны / ножки настенного монтажа (в комплекте или добор)", 5, "компл.", 800, 0.2, ""),
        ("Полотенцесушитель электрический в с/у", 1, "шт.", 6500, 0.2, ""),
        ("Программатор / Wi-Fi модуль для конвекторов (опция)", 3, "шт.", 3500, 0.3, "удобно для посуточной — прогрев перед заездом"),
    ],
))

SECTIONS.append((
    "8. ВЕНТИЛЯЦИЯ И ВОДОСНАБЖЕНИЕ (внутренняя разводка)",
    [
        ("Вентилятор вытяжной с/у + обратный клапан", 1, "шт.", 2800, 0.2, ""),
        ("Приточный клапан стеновой (если нет проветривания окнами)", 2, "шт.", 2500, 0.3, "опция; окна уже есть"),
        ("Труба ППР/металлопласт + фитинги (ХВС/ГВС)", 1, "компл.", 15000, 1.5, "внутренняя разводка до точек"),
        ("Канализация ПВХ 50/110 + фитинги", 1, "компл.", 8000, 1.0, ""),
        ("Водонагреватель накопительный 50–80 л", 1, "шт.", 18000, 0.4, "если нет ГВС"),
        ("Фильтр грубой очистки + редуктор", 1, "компл.", 4500, 0.2, ""),
        ("Запорная арматура, подводки, сифоны", 1, "компл.", 6000, 0.3, ""),
    ],
))

SECTIONS.append((
    "9. САНТЕХНИКА С/У",
    [
        ("Унитаз-компакт с микролифтом", 1, "шт.", 14000, 0.4, ""),
        ("Раковина + тумба подвесная 50–60 см", 1, "компл.", 16000, 0.4, "габарит с/у 3,32 м²"),
        ("Смеситель раковины", 1, "шт.", 4500, 0.1, ""),
        ("Душевая кабина / уголок 90×90 + поддон", 1, "компл.", 28000, 0.8, ""),
        ("Смеситель душа / стойка (если не в комплекте)", 1, "компл.", 5500, 0.2, ""),
        ("Зеркало / шкафчик зеркальный", 1, "шт.", 7000, 0.2, ""),
        ("Аксессуары с/у (держатель, крючки, полка)", 1, "компл.", 3500, 0.2, ""),
        ("Корзина для белья", 1, "шт.", 2000, 0.0, ""),
    ],
))

SECTIONS.append((
    "10. КУХНЯ (мебель + техника)",
    [
        ("Кухонный гарнитур 2,4–2,7 м (корпуса + фасады)", 1, "компл.", 85000, 1.5, ""),
        ("Столешница + плинтус", 3, "м.п.", 6500, 0.3, ""),
        ("Мойка нерж. + смеситель кухня", 1, "компл.", 12000, 0.3, ""),
        ("Варочная панель электрическая 4 конф.", 1, "шт.", 18000, 0.2, ""),
        ("Духовой шкаф электрический", 1, "шт.", 28000, 0.3, ""),
        ("Вытяжка", 1, "шт.", 12000, 0.3, ""),
        ("Холодильник двухкамерный", 1, "шт.", 45000, 0.2, ""),
        ("СВЧ", 1, "шт.", 9000, 0.1, ""),
        ("Посудомоечная машина 45 см", 1, "шт.", 35000, 0.4, "для посуточной очень желательно"),
        ("Стиральная машина 6–7 кг", 1, "шт.", 32000, 0.4, ""),
    ],
))

SECTIONS.append((
    "11. МЕБЕЛЬ — КУХНЯ-ГОСТИНАЯ + ПРИХОЖАЯ",
    [
        ("Диван-кровать раскладной 140–160", 1, "шт.", 45000, 0.3, "доп. спальное место"),
        ("Кресло / пуф", 1, "шт.", 8000, 0.1, ""),
        ("Журнальный столик", 1, "шт.", 6000, 0.1, ""),
        ("Тумба ТВ", 1, "шт.", 7000, 0.1, ""),
        ("Обеденный стол на 4–6 чел.", 1, "шт.", 12000, 0.2, ""),
        ("Стул обеденный", 4, "шт.", 3500, 0.2, ""),
        ("Стеллаж / шкаф открытый", 1, "шт.", 10000, 0.2, ""),
        ("Зеркало + вешалка + обувница (прихожая)", 1, "компл.", 9000, 0.3, ""),
    ],
))

SECTIONS.append((
    "12. МЕБЕЛЬ — СПАЛЬНИ (2 × 9,21 м²)",
    [
        ("Кровать 160×200 с основанием (спальня 1)", 1, "шт.", 22000, 0.3, ""),
        ("Матрас ортопедический 160×200", 1, "шт.", 18000, 0.1, ""),
        ("Кровать 140×200 с основанием (спальня 2)", 1, "шт.", 18000, 0.3, ""),
        ("Матрас 140×200", 1, "шт.", 14000, 0.1, ""),
        ("Тумба прикроватная", 4, "шт.", 3500, 0.3, ""),
        ("Шкаф 2-дверный", 2, "шт.", 22000, 0.8, ""),
        ("Комод", 1, "шт.", 10000, 0.2, ""),
        ("Зеркало настенное", 2, "шт.", 2500, 0.2, ""),
    ],
))

SECTIONS.append((
    "13. ТЕХНИКА И ЭЛЕКТРОНИКА ДЛЯ ГОСТЕЙ",
    [
        ("Телевизор Smart TV 43–50\"", 1, "шт.", 32000, 0.2, "гостиная"),
        ("Роутер Wi-Fi", 1, "компл.", 5500, 0.2, "обязательно"),
        ("Кондиционер сплит 9–12k BTU (гостиная)", 1, "компл.", 45000, 1.0, "лето — критично для сдачи"),
        ("Фен", 1, "шт.", 2500, 0.0, ""),
        ("Утюг + гладильная доска", 1, "компл.", 4500, 0.1, ""),
        ("Чайник электрический", 1, "шт.", 2500, 0.0, ""),
        ("Кофемашина капсульная / тостер", 1, "шт.", 8000, 0.0, ""),
        ("Пылесос", 1, "шт.", 7000, 0.0, ""),
    ],
))

SECTIONS.append((
    "14. ТЕКСТИЛЬ, ПОСТЕЛЬ, ШТОРЫ",
    [
        ("Комплект постельного белья (основной + сменный)", 6, "компл.", 3500, 0.0, "3 спальных места × 2"),
        ("Одеяло всесезонное", 3, "шт.", 4000, 0.0, ""),
        ("Подушка", 6, "шт.", 1500, 0.0, ""),
        ("Наматрасник непромокаемый", 3, "шт.", 2500, 0.0, "для аренды обязательно"),
        ("Покрывало / плед", 3, "шт.", 3000, 0.0, ""),
        ("Шторы блэкаут + тюль (спальни)", 2, "компл.", 8000, 0.5, ""),
        ("Шторы / римская (гостиная, на существующие окна)", 1, "компл.", 9000, 0.3, ""),
        ("Карнизы", 4, "шт.", 2000, 0.3, ""),
        ("Полотенца банные", 6, "шт.", 800, 0.0, ""),
        ("Полотенца лицевые", 6, "шт.", 400, 0.0, ""),
        ("Коврик с/у + прикроватные", 4, "шт.", 1200, 0.0, ""),
        ("Скатерть / дорожка", 1, "шт.", 1500, 0.0, ""),
    ],
))

SECTIONS.append((
    "15. ПОСУДА, БЫТ, СТАРТ ДЛЯ ПОСУТОЧНОЙ",
    [
        ("Сервиз тарелок/мисок на 6 персон", 1, "компл.", 5000, 0.0, ""),
        ("Кружки, бокалы, стаканы", 1, "компл.", 3500, 0.0, ""),
        ("Столовые приборы на 6 + запас", 1, "компл.", 3000, 0.0, ""),
        ("Кастрюли, сковороды, ножи, доски, лопатки", 1, "компл.", 8000, 0.0, ""),
        ("Контейнеры, открывалки и пр.", 1, "компл.", 2500, 0.0, ""),
        ("Сушилка для посуды, ведро для мусора", 1, "компл.", 2500, 0.0, ""),
        ("Вешалки плечики", 30, "шт.", 50, 0.0, ""),
        ("Расходники старт (пакеты, губки, мешки)", 1, "компл.", 3000, 0.0, ""),
        ("Старт гигиена: бумага, мыло, шампунь, гель", 1, "компл.", 2500, 0.0, ""),
        ("Аптечка + огнетушитель + датчик дыма", 1, "компл.", 4500, 0.2, ""),
        ("Сушилка для белья", 1, "шт.", 2000, 0.0, ""),
        ("Коврик грязезащитный у входа (внутри)", 1, "шт.", 2500, 0.0, ""),
    ],
))

SECTIONS.append((
    "16. ДЕКОР И ПРОЧЕЕ ДЛЯ СДАЧИ",
    [
        ("Картины / постеры", 1, "компл.", 5000, 0.2, ""),
        ("Декор / растения искусственные", 1, "компл.", 3000, 0.1, ""),
        ("Лампа прикроватная / торшер", 3, "шт.", 2500, 0.2, ""),
        ("Органайзеры для белья", 1, "компл.", 3000, 0.0, ""),
        ("Сейф небольшой (для хозяина)", 1, "шт.", 5000, 0.1, ""),
        ("Ключница + кейбокс / запасные ключи", 1, "компл.", 3500, 0.1, ""),
        ("Табличка Wi-Fi / правила дома", 1, "компл.", 1500, 0.0, ""),
        ("Запас: лампочки, батарейки, фильтры", 1, "компл.", 2000, 0.0, ""),
        ("Клининг-набор старт (химия, швабра, тряпки)", 1, "компл.", 6000, 0.0, ""),
    ],
))

SECTIONS.append((
    "17. РЕЗЕРВ (доставка и неучтённое)",
    [
        ("Доставка материалов", 1, "усл.", 20000, 0.0, ""),
        ("Резерв на крепёж / доборы / отходы 5%", 1, "усл.", 35000, 0.0, ""),
    ],
))

section_totals = []
n = 1
for sec_title, items in SECTIONS:
    row = add_section(ws, row, sec_title)
    sec_sum = 0.0
    sec_days = 0.0
    for name, qty, unit, price, days, note in items:
        row, total, d = add_item(ws, row, n, name, qty, unit, price, days, note)
        sec_sum += total
        sec_days += d
        n += 1
    short = sec_title.split(". ", 1)[-1] if ". " in sec_title else sec_title
    row = add_subtotal(ws, row, f"Итого: {short}", sec_sum, sec_days)
    section_totals.append((sec_title, sec_sum, sec_days))
    row += 1

grand_money = sum(s for _, s, _ in section_totals)
grand_days_sum = sum(d for _, _, d in section_totals)
calendar_days = 32

ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
cell = ws.cell(row=row, column=1, value="ИТОГО МАТЕРИАЛЫ И КОМПЛЕКТАЦИЯ (только внутрянка)")
cell.fill = grand_fill
cell.font = grand_font
for c in range(1, 6):
    ws.cell(row=row, column=c).fill = grand_fill
    ws.cell(row=row, column=c).border = thin
    ws.cell(row=row, column=c).font = grand_font
t = ws.cell(row=row, column=6, value=grand_money)
t.fill = grand_fill
t.font = grand_font
t.number_format = "#,##0"
t.border = thin
d = ws.cell(row=row, column=7, value=round(grand_days_sum, 1))
d.fill = grand_fill
d.font = grand_font
d.number_format = "0.0"
d.border = thin
ws.cell(row=row, column=8, value="сумма дней по позициям (часть работ параллельно)").fill = grand_fill
ws.cell(row=row, column=8).border = thin
ws.cell(row=row, column=8).font = note_font
row += 1

ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
ws.cell(row=row, column=1, value=(
    f"Ориентир календаря внутрянки бригадой 2 чел.: ≈ {calendar_days} раб. дней. "
    f"Работы в смету НЕ входят. Итого материалы+обстановка: ≈ {grand_money:,.0f} ₽. "
    f"С запасом 10%: ≈ {grand_money * 1.1:,.0f} ₽. "
    f"Не включено: окна, терраса, фасад, кровля, наружные работы."
).replace(",", " "))
ws.cell(row=row, column=1).font = note_font
ws.cell(row=row, column=1).alignment = wrap
ws.row_dimensions[row].height = 40

set_widths(ws)
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0

# ===================================================================
# ЛИСТ 2 — Сводка
# ===================================================================
ws2 = wb.create_sheet("Сводка по разделам")
ws2["A1"] = "Сводка затрат по разделам (только внутрянка)"
ws2["A1"].font = title_font
ws2.merge_cells("A1:D1")

for i, h in enumerate(["Раздел", "Сумма, ₽", "Доля, %", "Время монтажа, дн"], 1):
    ws2.cell(row=3, column=i, value=h)
style_header(ws2, 3, 4)

for i, (title, money, days) in enumerate(section_totals):
    r = 4 + i
    short = title.split(". ", 1)[-1] if ". " in title else title
    ws2.cell(row=r, column=1, value=short).border = thin
    c = ws2.cell(row=r, column=2, value=money)
    c.number_format = "#,##0"
    c.border = thin
    p = ws2.cell(row=r, column=3, value=round(100 * money / grand_money, 1) if grand_money else 0)
    p.number_format = "0.0"
    p.border = thin
    dcell = ws2.cell(row=r, column=4, value=round(days, 1))
    dcell.number_format = "0.0"
    dcell.border = thin

r = 4 + len(section_totals)
ws2.cell(row=r, column=1, value="ИТОГО").font = grand_font
ws2.cell(row=r, column=1).fill = grand_fill
ws2.cell(row=r, column=1).border = thin
c = ws2.cell(row=r, column=2, value=grand_money)
c.number_format = "#,##0"
c.font = grand_font
c.fill = grand_fill
c.border = thin
ws2.cell(row=r, column=3, value=100).fill = grand_fill
ws2.cell(row=r, column=3).border = thin
ws2.cell(row=r, column=3).font = grand_font
dcell = ws2.cell(row=r, column=4, value=round(grand_days_sum, 1))
dcell.fill = grand_fill
dcell.font = grand_font
dcell.border = thin

ws2.column_dimensions["A"].width = 58
ws2.column_dimensions["B"].width = 14
ws2.column_dimensions["C"].width = 10
ws2.column_dimensions["D"].width = 18

pie = PieChart()
pie.title = "Структура бюджета (внутрянка)"
labels = Reference(ws2, min_col=1, min_row=4, max_row=3 + len(section_totals))
data = Reference(ws2, min_col=2, min_row=3, max_row=3 + len(section_totals))
pie.add_data(data, titles_from_data=True)
pie.set_categories(labels)
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.dataLabels.showVal = False
pie.width = 22
pie.height = 14
ws2.add_chart(pie, "F3")

# ===================================================================
# ЛИСТ 3 — По комнатам
# ===================================================================
ws3 = wb.create_sheet("Обстановка по комнатам")
ws3["A1"] = "Чек-лист внутрянки и обстановки по комнатам (посуточная)"
ws3["A1"].font = title_font
ws3.merge_cells("A1:C1")

rooms = [
    ("Прихожая 2,80 м²", [
        "Вешалка / обувница / зеркало, коврик",
        "Конвектор 1 кВт",
        "Свет + розетки",
    ]),
    ("Кухня-гостиная 18,82 м²", [
        "Кухня + техника (холод, плита, духовка, вытяжка, СВЧ, ПММ)",
        "Обеденная зона 4–6 чел., диван-кровать, ТВ",
        "2 конвектора по 2 кВт + кондиционер",
        "Шторы на существующие окна/витражи, Wi-Fi",
    ]),
    ("Спальня 1 — 9,21 м²", [
        "Кровать 160×200 + матрас + наматрасник",
        "Шкаф, 2 тумбы, свет, шторы блэкаут",
        "Конвектор 1,5 кВт",
    ]),
    ("Спальня 2 — 9,21 м²", [
        "Кровать 140×200 + матрас + наматрасник",
        "Шкаф, 2 тумбы, свет, шторы блэкаут",
        "Конвектор 1,5 кВт",
    ]),
    ("Санузел 3,32 м²", [
        "Унитаз, раковина, душевая, стиралка",
        "Плитка, тёплый пол, полотенцесушитель эл., вентилятор",
        "Полотенца и расходники",
    ]),
]

r = 3
for room, items in rooms:
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    cell = ws3.cell(row=r, column=1, value=room)
    cell.fill = section_fill
    cell.font = section_font
    for c in range(1, 4):
        ws3.cell(row=r, column=c).fill = section_fill
        ws3.cell(row=r, column=c).border = thin
    r += 1
    for it in items:
        ws3.cell(row=r, column=1, value="•").border = thin
        ws3.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws3.cell(row=r, column=2, value=it).border = thin
        ws3.cell(row=r, column=3).border = thin
        r += 1
    r += 1

ws3.column_dimensions["A"].width = 4
ws3.column_dimensions["B"].width = 72
ws3.column_dimensions["C"].width = 16

# ===================================================================
# ЛИСТ 4 — Исходные данные
# ===================================================================
ws4 = wb.create_sheet("Исходные данные")
ws4["A1"] = "Параметры объекта"
ws4["A1"].font = title_font
data_rows = [
    ("Габарит", "6 120 × 10 000 мм"),
    ("Жилая площадь", f"{FLOOR_LIVING} м²"),
    ("Прихожая / с/у / кухня-гостиная", f"{AREAS['прихожая']} / {AREAS['санузел']} / {AREAS['кухня_гостиная']} м²"),
    ("Спальни", f"{AREAS['спальня1']} + {AREAS['спальня2']} м²"),
    ("Высота потолка (принято)", f"{CEILING_H} м"),
    ("Состояние", "Каркас перегородок; окна и терраса УЖЕ ЕСТЬ"),
    ("Отделка стен/потолков", "Вагонка штиль белая (как на фото) — без ГКЛ"),
    ("Отопление", "Электрические конвекторы (5 шт. + полотенцесушитель)"),
    ("Уличное освещение", "Включено: 3 настенных + 2 прожектора + датчик"),
    ("Смета включает", "Внутренняя отделка вагонкой + инженерия + уличная подсветка + обстановка"),
    ("Смета НЕ включает", "Окна, настил террасы, фасад, кровля, фундамент, стоимость монтажа"),
    ("Вместимость", "4–6 гостей"),
]
ws4["A3"] = "Параметр"
ws4["B3"] = "Значение"
style_header(ws4, 3, 2)
for i, (a, b) in enumerate(data_rows):
    ws4.cell(row=4 + i, column=1, value=a).border = thin
    ws4.cell(row=4 + i, column=2, value=b).border = thin
    ws4.cell(row=4 + i, column=2).alignment = wrap
ws4.column_dimensions["A"].width = 34
ws4.column_dimensions["B"].width = 78

ws4.cell(row=16, column=1, value="Календарь внутрянки (бригада 2 чел.)").font = section_font
stages = [
    "1–3 дн: электрика черновая + вода/канализация + линии под конвекторы и улицу",
    "4–10 дн: обрешётка + вагонка стены и потолки + покраска белая",
    "11–12 дн: откосы окон, наличники, галтели",
    "13–16 дн: с/у гидроизоляция + плитка + тёплый пол",
    "17–19 дн: ламинат, плинтуса",
    "20–22 дн: двери, чистовая электрика, конвекторы, уличная подсветка, сантехника",
    "23–27 дн: кухня, стиралка, кондиционер",
    "28–32 дн: мебель, текстиль, посуда, фото для объявления",
]
for i, s in enumerate(stages):
    ws4.cell(row=17 + i, column=1, value=s)
    ws4.merge_cells(start_row=17 + i, start_column=1, end_row=17 + i, end_column=2)

# Save to several obvious locations
paths = [
    Path("/workspace/Smeta_vnutryanka_posutochno.xlsx"),
    Path("/workspace/smeta/Smeta_vnutryanka_posutochno.xlsx"),
    Path("/workspace/smeta/Smeta_karkasnyi_dom_posutochno.xlsx"),
    Path("/opt/cursor/artifacts/Smeta_vnutryanka_posutochno.xlsx"),
]
primary = paths[0]
wb.save(primary)
for p in paths[1:]:
    p.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(primary, p)

print(f"Saved primary: {primary}")
for p in paths:
    print(f"  copy: {p} ({p.stat().st_size} bytes)")
print(f"Living area: {FLOOR_LIVING} m2")
print(f"Grand total: {grand_money:,.0f} RUB")
print(f"Sum of install days: {grand_days_sum:.1f}")
print(f"Line items: {n - 1}")
for title, money, days in section_totals:
    short = title.split(". ", 1)[-1]
    print(f"  {short}: {money:,.0f} ₽ / {days:.1f} дн")
