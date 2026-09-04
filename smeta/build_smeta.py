#!/usr/bin/env python3
"""Смета на работы: ремонт ванной и туалета, серия ЛГ-137. Только работы."""

from __future__ import annotations

from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins


# ---------------------------------------------------------------------------
# Palette
# ---------------------------------------------------------------------------
NAVY = "1B365D"
NAVY_DEEP = "12243F"
STEEL = "3D5A80"
GOLD = "C4A35A"
GOLD_SOFT = "E6D3A3"
CREAM = "FBF8F2"
SAND = "F3EBDD"
ROW_ALT = "F7F3EA"
WHITE = "FFFFFF"
INK = "1F2933"
MUTED = "5C6670"
LINE = "D7CDBD"
LINE_SOFT = "E8E0D4"
SUBTOTAL_BG = "E7EEF6"
REC_BG = "F6EEDC"
TOTAL_BG = "1B365D"
GREEN = "2F6B4F"
TERRACOTTA = "A65D3F"
PALE_BLUE = "EEF3F8"

OUT = Path(__file__).resolve().parent / "Smeta_raboty_vanna_tualet_LG-137.xlsx"

# Типовые габариты сантехкабины ЛГ-137
BATH_L, BATH_W, H = 1.70, 1.50, 2.55
WC_L, WC_W = 1.20, 0.82
DOOR_W, DOOR_H = 0.57, 1.98

BATH_FLOOR = round(BATH_L * BATH_W, 2)  # 2.55
WC_FLOOR = round(WC_L * WC_W, 2)  # 0.98
FLOOR = round(BATH_FLOOR + WC_FLOOR, 2)  # 3.53
DOOR_AREA = round(DOOR_W * DOOR_H, 2)  # 1.13
BATH_WALLS = round(2 * (BATH_L + BATH_W) * H - DOOR_AREA, 2)  # 15.19
WC_WALLS = round(2 * (WC_L + WC_W) * H - DOOR_AREA, 2)  # 9.17
WALLS = round(BATH_WALLS + WC_WALLS, 2)  # 24.36
BOX_TILE = 1.50
WC_TILE_WALLS = round(WC_WALLS + BOX_TILE, 2)  # 10.67
HYDRO_WALL_H = 0.70
HYDRO_BATH_WALLS = round((2 * (BATH_L + BATH_W) - DOOR_W) * HYDRO_WALL_H, 2)
HYDRO_WC_WALLS = round((2 * (WC_L + WC_W) - DOOR_W) * HYDRO_WALL_H, 2)
HYDRO_WALLS = round(HYDRO_BATH_WALLS + HYDRO_WC_WALLS, 2)
SHOWER_ZONE = round((BATH_L + 0.75) * (2.00 - HYDRO_WALL_H), 2)
GROUT = round(BATH_FLOOR + WC_FLOOR + BATH_WALLS + WC_TILE_WALLS, 2)

WORK_DAYS = 18


def money(n: float) -> str:
    return f"{int(round_rub(n)):,.0f}".replace(",", " ")


def round_rub(n: float) -> int:
    """Как Excel ROUND(..., 0) — 0,5 вверх."""
    return int(Decimal(str(n)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def rub_words(n: int) -> str:
    """Сумма прописью, рубли."""
    units = (
        ("", "", ""),
        ("тысяча", "тысячи", "тысяч"),
        ("миллион", "миллиона", "миллионов"),
    )
    ones_m = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
    ones_f = ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
    teens = [
        "десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать",
        "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать",
    ]
    tens = [
        "", "", "двадцать", "тридцать", "сорок", "пятьдесят",
        "шестьдесят", "семьдесят", "восемьдесят", "девяносто",
    ]
    hundreds = [
        "", "сто", "двести", "триста", "четыреста", "пятьсот",
        "шестьсот", "семьсот", "восемьсот", "девятьсот",
    ]

    def tri(num: int, feminine: bool) -> str:
        h, rest = divmod(num, 100)
        t, o = divmod(rest, 10)
        ones = ones_f if feminine else ones_m
        parts = []
        if h:
            parts.append(hundreds[h])
        if t == 1:
            parts.append(teens[o])
        else:
            if t:
                parts.append(tens[t])
            if o:
                parts.append(ones[o])
        return " ".join(parts)

    def form(num: int, forms: tuple[str, str, str]) -> str:
        n100 = num % 100
        n10 = num % 10
        if 11 <= n100 <= 19:
            return forms[2]
        if n10 == 1:
            return forms[0]
        if 2 <= n10 <= 4:
            return forms[1]
        return forms[2]

    if n == 0:
        return "Ноль рублей 00 копеек"

    n = int(n)
    chunks = []
    tmp = n
    while tmp:
        chunks.append(tmp % 1000)
        tmp //= 1000

    words = []
    for i, chunk in enumerate(chunks):
        if chunk == 0:
            continue
        feminine = i == 1
        w = tri(chunk, feminine)
        if i:
            w = f"{w} {form(chunk, units[i])}".strip()
        words.append(w)
    words.reverse()
    body = " ".join(words).strip()
    body = body[0].upper() + body[1:]
    rub_form = form(n, ("рубль", "рубля", "рублей"))
    return f"{body} {rub_form} 00 копеек"


# ---------------------------------------------------------------------------
# Estimate data
# ---------------------------------------------------------------------------
# kind: section | item | note
# recommended items live in section 8

SECTIONS: list[dict] = [
    {
        "num": "1",
        "title": "Организация работ в жилой квартире",
        "intro": "Квартира остаётся жилой. Унитаз каждое утро снимается для производства работ и каждый вечер ставится обратно.",
        "items": [
            {
                "code": "1.1",
                "name": "Защита коридора, дверей и чистых помещений (плёнка, картон, малярная лента)",
                "unit": "компл.",
                "qty": 1,
                "price": 4_500,
                "note": "На весь период ремонта",
            },
            {
                "code": "1.2",
                "name": "Организация рабочего места и ежедневная уборка зоны работ",
                "unit": "смена",
                "qty": WORK_DAYS,
                "price": 500,
                "note": f"{WORK_DAYS} рабочих дней",
            },
            {
                "code": "1.3",
                "name": "Временный унитаз: монтаж вечером и демонтаж утром (проживание на объекте)",
                "unit": "цикл",
                "qty": WORK_DAYS,
                "price": 1_300,
                "note": "Цикл = вечер + утро. Унитаз заказчика или технический",
            },
            {
                "code": "1.4",
                "name": "Перекрытие внутриквартирных отводов, организация временного водоснабжения",
                "unit": "компл.",
                "qty": 1,
                "price": 2_000,
                "note": "Стояки ХВС / ГВС не отключаем без необходимости",
            },
        ],
    },
    {
        "num": "2",
        "title": "Демонтаж и вынос (без сбивки плитки)",
        "intro": "Демонтаж плитки не выполняется. Потолок не затрагивается. Стояки воды и канализации сохраняются.",
        "items": [
            {
                "code": "2.1",
                "name": "Демонтаж унитаза с сохранением выпуска канализации",
                "unit": "шт.",
                "qty": 1,
                "price": 2_800,
                "note": "Аккуратно: чугунный раструб стояка не ломаем",
            },
            {
                "code": "2.2",
                "name": "Демонтаж раковины, смесителей, полок и навесных аксессуаров",
                "unit": "компл.",
                "qty": 1,
                "price": 2_200,
                "note": "",
            },
            {
                "code": "2.3",
                "name": "Отключение и вынос стиральной машины из помещения",
                "unit": "шт.",
                "qty": 1,
                "price": 2_500,
                "note": "Обратная установка — раздел 7",
            },
            {
                "code": "2.4",
                "name": "Демонтаж чугунной ванны (узкий проём 57 см, резка чаши при необходимости)",
                "unit": "шт.",
                "qty": 1,
                "price": 9_000,
                "note": "Типично для ЛГ-137: ванна шире дверного проёма",
            },
            {
                "code": "2.5",
                "name": "Спуск и вынос чугунной ванны / частей из квартиры",
                "unit": "шт.",
                "qty": 1,
                "price": 7_000,
                "note": "Лифт + вынос на контейнер / к машине заказчика",
            },
            {
                "code": "2.6",
                "name": "Срезка старых труб ХВС и ГВС внутриквартирной разводки",
                "unit": "компл.",
                "qty": 1,
                "price": 5_000,
                "note": "Стояки не трогаем, отводы сохраняем для врезки",
            },
            {
                "code": "2.7",
                "name": "Демонтаж старых отводов канализации от ванны, раковины и стиральной машины",
                "unit": "компл.",
                "qty": 1,
                "price": 3_500,
                "note": "Канализационный стояк не меняем",
            },
            {
                "code": "2.8",
                "name": "Извлечение деревянной закладной под унитаз и заделка полости",
                "unit": "шт.",
                "qty": 1,
                "price": 2_800,
                "note": "Типовой узел ЛГ-137, доска обычно прогнившая",
            },
            {
                "code": "2.9",
                "name": "Подготовка плинтусов и примыканий у порогов (без работ по стоякам)",
                "unit": "компл.",
                "qty": 1,
                "price": 2_500,
                "note": "",
            },
            {
                "code": "2.10",
                "name": "Сбор, мешкование, спуск и вынос строительного мусора",
                "unit": "компл.",
                "qty": 1,
                "price": 9_000,
                "note": "Контейнер / утилизация — по факту, если требуется отдельно",
            },
        ],
    },
    {
        "num": "3",
        "title": "Водоснабжение: коллекторы, фильтры 10\", редукторы, полипропилен, водорозетки",
        "intro": "Коллекторная схема. На ХВС и ГВС: фильтр стандарт 10\" и редуктор давления. Стояки не перевариваются.",
        "items": [
            {
                "code": "3.1",
                "name": "Сборка коллекторного узла ХВС (коллектор, фильтр стандарт 10\", редуктор, краны)",
                "unit": "компл.",
                "qty": 1,
                "price": 9_500,
                "note": "Материалы заказчика. Схема: кран → фильтр 10\" → редуктор → коллектор",
            },
            {
                "code": "3.2",
                "name": "Сборка коллекторного узла ГВС (коллектор, фильтр стандарт 10\", редуктор, краны)",
                "unit": "компл.",
                "qty": 1,
                "price": 9_500,
                "note": "Аналогично ХВС",
            },
            {
                "code": "3.3",
                "name": "Обвязка существующих отводов со стояков: краны, счётчики, обратные клапаны",
                "unit": "компл.",
                "qty": 1,
                "price": 6_000,
                "note": "Стояки и полотенцесушитель на стояке не меняем",
            },
            {
                "code": "3.4",
                "name": "Монтаж ревизионного шкафа / люка доступа к коллекторам и фильтрам",
                "unit": "шт.",
                "qty": 1,
                "price": 4_000,
                "note": "Смена картриджей 10\" без разбора короба",
            },
            {
                "code": "3.5",
                "name": "Прокладка труб полипропилен ХВС/ГВС до раковины, унитаза, стиральной машины и ванны",
                "unit": "м.п.",
                "qty": 32,
                "price": 850,
                "note": "Скрыто / в коробе. Пайка, крепление, теплоизоляция ГВС",
            },
            {
                "code": "3.6",
                "name": "Водорозетки под смеситель ванны (холодная + горячая), вывод в плоскость плитки",
                "unit": "компл.",
                "qty": 1,
                "price": 4_200,
                "note": "2 точки, жёсткая фиксация",
            },
            {
                "code": "3.7",
                "name": "Водорозетки под смеситель раковины (холодная + горячая)",
                "unit": "компл.",
                "qty": 1,
                "price": 4_200,
                "note": "2 точки",
            },
            {
                "code": "3.8",
                "name": "Водорозетка холодной воды под стиральную машину с отсечным краном",
                "unit": "шт.",
                "qty": 1,
                "price": 2_500,
                "note": "1 точка",
            },
            {
                "code": "3.9",
                "name": "Водорозетка холодной воды к инсталляции",
                "unit": "шт.",
                "qty": 1,
                "price": 2_200,
                "note": "1 точка",
            },
            {
                "code": "3.10",
                "name": "Штробление и проходы в стенах сантехкабины, заделка штраб",
                "unit": "м.п.",
                "qty": 8,
                "price": 900,
                "note": "Стены тонкие, армированные — работа аккуратная",
            },
            {
                "code": "3.11",
                "name": "Опрессовка системы водоснабжения и проверка соединений",
                "unit": "шт.",
                "qty": 1,
                "price": 4_000,
                "note": "До закрытия коробов и до чистовой отделки",
            },
        ],
    },
    {
        "num": "4",
        "title": "Канализация (без стояка) и инсталляция",
        "intro": "Стояк канализации не меняем. Для подвесного унитаза — переход с чугунного отвода на пластик и монтаж инсталляции.",
        "items": [
            {
                "code": "4.1",
                "name": "Прокладка канализации Ø50 к ванне, раковине и стиральной машине",
                "unit": "м.п.",
                "qty": 8,
                "price": 750,
                "note": "Уклоны, ревизии, хомуты",
            },
            {
                "code": "4.2",
                "name": "Прокладка канализации Ø110 от выпуска стояка до инсталляции",
                "unit": "м.п.",
                "qty": 1.5,
                "price": 1_400,
                "note": "Стояк не заменяем",
            },
            {
                "code": "4.3",
                "name": "Расчеканка чугунного отвода и переход на пластик без замены стояка",
                "unit": "шт.",
                "qty": 1,
                "price": 14_000,
                "note": "Ключевой узел ЛГ-137. Стояк чугунный сохраняется",
            },
            {
                "code": "4.4",
                "name": "Монтаж инсталляции: рама, крепление к полу и стенам, выставление",
                "unit": "шт.",
                "qty": 1,
                "price": 10_000,
                "note": "Туалет узкий: рама подбирается по месту",
            },
            {
                "code": "4.5",
                "name": "Подключение воды и канализации к инсталляции, звукоизоляция рамы",
                "unit": "шт.",
                "qty": 1,
                "price": 5_000,
                "note": "Демпферная лента, проверка бака",
            },
            {
                "code": "4.6",
                "name": "Короб инсталляции из влагостойкого ГВЛ, усиление под плитку",
                "unit": "м²",
                "qty": 2.5,
                "price": 3_200,
                "note": "Плоскость под облицовку",
            },
            {
                "code": "4.7",
                "name": "Ревизионный люк под кнопку смыва",
                "unit": "шт.",
                "qty": 1,
                "price": 2_800,
                "note": "",
            },
            {
                "code": "4.8",
                "name": "Установка кнопки смыва после облицовки",
                "unit": "шт.",
                "qty": 1,
                "price": 2_200,
                "note": "Чистовой этап",
            },
        ],
    },
    {
        "num": "5",
        "title": "Подготовка оснований и гидроизоляция",
        "intro": "Пол — гидроизоляция + керамогранит. Стены — гидроизоляция до уровня ванны + плитка. Потолок не трогаем.",
        "items": [
            {
                "code": "5.1",
                "name": "Грунтование пола ванной и туалета",
                "unit": "м²",
                "qty": FLOOR,
                "price": 280,
                "note": f"{FLOOR} м²",
            },
            {
                "code": "5.2",
                "name": "Устройство выравнивающей стяжки пола",
                "unit": "м²",
                "qty": FLOOR,
                "price": 1_500,
                "note": "Под керамогранит, с учётом порогов",
            },
            {
                "code": "5.3",
                "name": "Гидроизоляция пола обмазочная в 2 слоя с заходом на стены 150–200 мм",
                "unit": "м²",
                "qty": FLOOR,
                "price": 1_200,
                "note": "Ванна + туалет, угловые ленты",
            },
            {
                "code": "5.4",
                "name": "Грунтование стен (бетоноконтакт — если старое покрытие сохраняется)",
                "unit": "м²",
                "qty": WALLS,
                "price": 220,
                "note": "Без демонтажа плитки",
            },
            {
                "code": "5.5",
                "name": "Выравнивание стен под плитку без демонтажа старой плитки",
                "unit": "м²",
                "qty": WALLS,
                "price": 980,
                "note": "Штукатурка / ремонтный состав по существующему основанию",
            },
            {
                "code": "5.6",
                "name": "Обмазочная гидроизоляция стен до уровня ванны (h ≈ 0,70 м)",
                "unit": "м²",
                "qty": HYDRO_WALLS,
                "price": 950,
                "note": "Ванна и туалет",
            },
            {
                "code": "5.7",
                "name": "Гидроизоляция душевой зоны над ванной до высоты 2,0 м",
                "unit": "м²",
                "qty": SHOWER_ZONE,
                "price": 980,
                "note": "Необходимо: душ над ванной",
            },
        ],
    },
    {
        "num": "6",
        "title": "Пороги и облицовка (керамогранит / плитка)",
        "intro": "Подбивка порогов двух проёмов и заведение плитки под двери ванны и туалета. Потолок не облицовывается.",
        "items": [
            {
                "code": "6.1",
                "name": "Подбивка, усиление и выравнивание порога проёма ванной",
                "unit": "шт.",
                "qty": 1,
                "price": 4_500,
                "note": "Проём ≈ 570×1980 мм",
            },
            {
                "code": "6.2",
                "name": "Подбивка, усиление и выравнивание порога проёма туалета",
                "unit": "шт.",
                "qty": 1,
                "price": 4_500,
                "note": "Проём ≈ 570×1980 мм",
            },
            {
                "code": "6.3",
                "name": "Гидроизоляция порогов",
                "unit": "шт.",
                "qty": 2,
                "price": 1_300,
                "note": "Стык «мокрых» помещений с коридором",
            },
            {
                "code": "6.4",
                "name": "Укладка керамогранита на пол ванной",
                "unit": "м²",
                "qty": BATH_FLOOR,
                "price": 2_800,
                "note": "Прямая раскладка, подрезка у ванны и стен",
            },
            {
                "code": "6.5",
                "name": "Укладка керамогранита на пол туалета (стеснённые условия)",
                "unit": "м²",
                "qty": WC_FLOOR,
                "price": 3_200,
                "note": "Площадь ≈ 1 м², много подрезки",
            },
            {
                "code": "6.6",
                "name": "Облицовка стен ванной плиткой",
                "unit": "м²",
                "qty": BATH_WALLS,
                "price": 2_400,
                "note": "До существующего потолка, потолок не трогаем",
            },
            {
                "code": "6.7",
                "name": "Облицовка стен туалета и короба инсталляции плиткой",
                "unit": "м²",
                "qty": WC_TILE_WALLS,
                "price": 2_500,
                "note": "Включая плоскость инсталляции",
            },
            {
                "code": "6.8",
                "name": "Укладка плитки на пороги / заведение под дверные полотна ванны и туалета",
                "unit": "шт.",
                "qty": 2,
                "price": 3_200,
                "note": "Кусок плитки под каждую дверь",
            },
            {
                "code": "6.9",
                "name": "Затирка швов пола и стен влагостойкая",
                "unit": "м²",
                "qty": GROUT,
                "price": 380,
                "note": "",
            },
            {
                "code": "6.10",
                "name": "Раскладки / кромки, подрезка, сверление под водорозетки и выпуски",
                "unit": "м.п.",
                "qty": 16,
                "price": 380,
                "note": "Наружные углы, откосы проёмов",
            },
            {
                "code": "6.11",
                "name": "Экран ванны (блоки или каркас) с облицовкой и ревизионным люком",
                "unit": "м²",
                "qty": 1.6,
                "price": 3_400,
                "note": "Доступ к сифону",
            },
            {
                "code": "6.12",
                "name": "Герметизация примыканий ванны, раковины и углов санитарным силиконом",
                "unit": "м.п.",
                "qty": 12,
                "price": 300,
                "note": "Цвет по затирке",
            },
        ],
    },
    {
        "num": "7",
        "title": "Установка оборудования: ванна из наливного камня, смесители, душ, раковина, унитаз",
        "intro": "Ванна из наливного камня тяжёлая: обязателен опорный подиум, занос через узкий проём 57 см (3–4 человека, при необходимости через окно).",
        "items": [
            {
                "code": "7.1",
                "name": "Устройство опорного подиума / ложемента под ванну из наливного камня",
                "unit": "шт.",
                "qty": 1,
                "price": 14_000,
                "note": "Ножки производителя не держат вес — опора на пол по контуру",
            },
            {
                "code": "7.2",
                "name": "Занос тяжёлой ванны (узкий проём 57 см, 3–4 человека)",
                "unit": "шт.",
                "qty": 1,
                "price": 8_000,
                "note": "Если не проходит — занос через окно, уточняется на замере",
            },
            {
                "code": "7.3",
                "name": "Установка и выверка ванны из наливного камня, примыкание к стенам",
                "unit": "шт.",
                "qty": 1,
                "price": 12_000,
                "note": "Уровень, зазоры под силикон, защита чаши",
            },
            {
                "code": "7.4",
                "name": "Монтаж слива-перелива и сифона ванны",
                "unit": "шт.",
                "qty": 1,
                "price": 3_500,
                "note": "Выпуск в канализацию Ø50",
            },
            {
                "code": "7.5",
                "name": "Установка смесителя на ванну",
                "unit": "шт.",
                "qty": 1,
                "price": 3_200,
                "note": "На водорозетки, эксцентрики, декоративные чашки",
            },
            {
                "code": "7.6",
                "name": "Монтаж душа: стойка / лейка / шланг / держатель",
                "unit": "шт.",
                "qty": 1,
                "price": 4_200,
                "note": "Душ над ванной (кабина в габарит 1700×1500 не закладывалась)",
            },
            {
                "code": "7.7",
                "name": "Установка раковины с креплением и сифоном",
                "unit": "шт.",
                "qty": 1,
                "price": 5_000,
                "note": "Подвесная / на кронштейнах / с тумбой — по модели заказчика",
            },
            {
                "code": "7.8",
                "name": "Установка смесителя раковины",
                "unit": "шт.",
                "qty": 1,
                "price": 2_500,
                "note": "",
            },
            {
                "code": "7.9",
                "name": "Навеска унитаза на инсталляцию, регулировка, проверка смыва",
                "unit": "шт.",
                "qty": 1,
                "price": 5_000,
                "note": "После плитки и набора прочности",
            },
            {
                "code": "7.10",
                "name": "Возврат и подключение стиральной машины (вода, слив, проверка)",
                "unit": "шт.",
                "qty": 1,
                "price": 2_800,
                "note": "Вынос — в разделе 2",
            },
            {
                "code": "7.11",
                "name": "Чистовая установка отражателей водорозеток и подключение гибких подводок",
                "unit": "компл.",
                "qty": 1,
                "price": 2_000,
                "note": "6 водорозеток",
            },
        ],
    },
    {
        "num": "8",
        "title": "Рекомендуемые работы (для законченного результата)",
        "intro": "Не входят в обязательный минимум по заданию, но на практике нужны. Можно исключить до договора.",
        "recommended": True,
        "items": [
            {
                "code": "8.1",
                "name": "Декоративный короб стояков без вмешательства в стояки, ревизионный люк",
                "unit": "шт.",
                "qty": 1,
                "price": 9_000,
                "note": "Эстетика туалета, доступ к счётчикам и фильтрам",
            },
            {
                "code": "8.2",
                "name": "Установка вытяжного вентилятора в существующий канал",
                "unit": "шт.",
                "qty": 1,
                "price": 2_800,
                "note": "Потолок не разбираем, по месту канала",
            },
            {
                "code": "8.3",
                "name": "Розетка IP44 для стиральной машины, кабель, УЗО / дифавтомат",
                "unit": "шт.",
                "qty": 1,
                "price": 5_500,
                "note": "Влажная зона, заземление",
            },
            {
                "code": "8.4",
                "name": "Розетка IP44 у раковины (фен / бритва)",
                "unit": "шт.",
                "qty": 1,
                "price": 3_200,
                "note": "",
            },
            {
                "code": "8.5",
                "name": "Навеска зеркала над раковиной",
                "unit": "шт.",
                "qty": 1,
                "price": 2_500,
                "note": "",
            },
            {
                "code": "8.6",
                "name": "Крепление аксессуаров: бумагодержатель, крючки, ёрш",
                "unit": "компл.",
                "qty": 1,
                "price": 2_200,
                "note": "Сверление по плитке",
            },
            {
                "code": "8.7",
                "name": "Монтаж стеклянной шторки на ванну",
                "unit": "шт.",
                "qty": 1,
                "price": 6_500,
                "note": "Душ над ванной без залива коридора",
            },
            {
                "code": "8.8",
                "name": "Финальная уборка ванной и туалета при сдаче",
                "unit": "шт.",
                "qty": 1,
                "price": 4_500,
                "note": "",
            },
        ],
    },
]


def thin_border(color=LINE):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def med_border(color=NAVY):
    s = Side(style="medium", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def font(name="Calibri", size=11, bold=False, color=INK, italic=False):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)


def fill(color):
    return PatternFill("solid", fgColor=color)


def align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def apply_print(ws, landscape=True, fit_width=True, title="Смета"):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1 if fit_width else 0
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.7, bottom=0.7, header=0.3, footer=0.3)
    ws.oddHeader.left.text = "&K1B365DСмета на работы · ЛГ-137"
    ws.oddHeader.right.text = "&K5C6670только работы"
    ws.oddFooter.left.text = f"&K5C6670{title}"
    ws.oddFooter.center.text = "&K5C6670стр. &P из &N"
    ws.oddFooter.right.text = "&K5C6670не является публичной офертой"
    ws.print_options.horizontalCentered = True
    ws.sheet_view.showGridLines = False
    ws.sheet_view.view = "pageBreakPreview"
    ws.sheet_view.view = "normal"


def set_col_widths(ws, widths: dict[str, float]):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def paint_range(ws, row, c1, c2, fl=None, ft=None, al=None, bd=None, height=None):
    if height:
        ws.row_dimensions[row].height = height
    for col in range(c1, c2 + 1):
        cell = ws.cell(row, col)
        if fl:
            cell.fill = fl
        if ft:
            cell.font = ft
        if al:
            cell.alignment = al
        if bd:
            cell.border = bd


def kpi_box(ws, row, col, label, value, sub, bg=NAVY, fg=WHITE, accent=GOLD):
    """2-column KPI: merge col:col+1 over 3 rows."""
    r1, r2, r3 = row, row + 1, row + 2
    ws.merge_cells(start_row=r1, start_column=col, end_row=r1, end_column=col + 1)
    ws.merge_cells(start_row=r2, start_column=col, end_row=r2, end_column=col + 1)
    ws.merge_cells(start_row=r3, start_column=col, end_row=r3, end_column=col + 1)
    for r in (r1, r2, r3):
        paint_range(ws, r, col, col + 1, fill(bg), al=align("center", "center"))
        for c in (col, col + 1):
            ws.cell(r, c).fill = fill(bg)
            ws.cell(r, c).alignment = align("center", "center")
            ws.cell(r, c).border = Border(
                left=Side(style="thin", color=bg),
                right=Side(style="thin", color=bg),
                top=Side(style="thin", color=bg),
                bottom=Side(style="thin", color=bg),
            )
    ws.cell(r1, col).value = label
    ws.cell(r1, col).font = font(size=9, bold=True, color=accent)
    ws.cell(r2, col).value = value
    ws.cell(r2, col).font = font(size=16, bold=True, color=fg)
    ws.cell(r3, col).value = sub
    ws.cell(r3, col).font = font(size=8, color=GOLD_SOFT)
    ws.row_dimensions[r1].height = 18
    ws.row_dimensions[r2].height = 28
    ws.row_dimensions[r3].height = 18


def build_cover(wb: Workbook, totals: dict):
    ws = wb.active
    ws.title = "Титул"
    apply_print(ws, landscape=False, title="Титульный лист")
    set_col_widths(ws, {c: 12.5 for c in "ABCDEFGH"})
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["H"].width = 4
    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToHeight = 1
    ws.page_setup.fitToWidth = 1

    max_r, max_c = 48, 8
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            ws.cell(r, c).fill = fill(CREAM)

    # top navy bar
    for r in (1, 2, 3):
        paint_range(ws, r, 1, 8, fill(NAVY_DEEP), height=10 if r != 2 else 22)
    ws.merge_cells("B2:G2")
    ws["B2"] = "ЛОКАЛЬНАЯ СМЕТА  ·  ТОЛЬКО РАБОТЫ"
    ws["B2"].font = font(size=11, bold=True, color=GOLD)
    ws["B2"].alignment = align("left", "center")
    ws["B2"].fill = fill(NAVY_DEEP)

    paint_range(ws, 4, 1, 8, fill(GOLD), height=6)

    ws.row_dimensions[5].height = 16
    ws.row_dimensions[6].height = 36
    ws.merge_cells("B6:G6")
    ws["B6"] = "Ремонт ванной комнаты и туалета"
    ws["B6"].font = font(size=26, bold=True, color=NAVY)
    ws["B6"].alignment = align("left", "center")
    ws["B6"].fill = fill(CREAM)

    ws.merge_cells("B7:G7")
    ws["B7"] = "Дом серии ЛГ-137  ·  раздельный санузел (сантехкабина)"
    ws["B7"].font = font(size=13, color=STEEL)
    ws["B7"].alignment = align("left", "center")
    ws["B7"].fill = fill(CREAM)
    ws.row_dimensions[7].height = 22

    ws.merge_cells("B8:G8")
    ws["B8"] = "Подробный расчёт работ для заказчика  ·  материалы в смету не входят"
    ws["B8"].font = font(size=11, italic=True, color=MUTED)
    ws["B8"].fill = fill(CREAM)
    ws.row_dimensions[8].height = 20

    paint_range(ws, 9, 1, 8, fill(CREAM), height=10)

    # meta table
    meta = [
        ("Объект", "Ванная 1700×1500 мм, туалет 1200×820 мм, высота 2550 мм"),
        ("Серия дома", "ЛГ-137 (сантехкабина, дверные проёмы ≈ 570×1980 мм)"),
        ("Состав сметы", "Только работы. Сантехника, плитка, трубы, смеси — закупка заказчика"),
        ("Не входит", "Демонтаж плитки · потолок · стояки ХВС/ГВС · канализационный стояк"),
        ("Условия", "В квартире живут: ежедневный временный унитаз, вынос стиральной машины"),
        ("Ванна (новая)", "Из наливного камня, тяжёлая — подиум, занос 3–4 человека"),
        ("Срок работ", f"{WORK_DAYS} рабочих дней (ориентир, без учёта поставки материалов)"),
        ("Дата сметы", date.today().strftime("%d.%m.%Y")),
        ("Срок действия", "30 календарных дней"),
        ("НДС", "Не облагается / по статусу исполнителя — уточняется в договоре"),
    ]
    r = 10
    ws.merge_cells(f"B{r}:C{r}")
    ws.merge_cells(f"D{r}:G{r}")
    for i, (k, v) in enumerate(meta):
        rr = r + i
        ws.merge_cells(f"B{rr}:C{rr}")
        ws.merge_cells(f"D{rr}:G{rr}")
        bg = WHITE if i % 2 == 0 else SAND
        paint_range(ws, rr, 2, 7, fill(bg), height=20)
        ws.cell(rr, 2).value = k
        ws.cell(rr, 2).font = font(size=10, bold=True, color=NAVY)
        ws.cell(rr, 2).alignment = align("left", "center")
        ws.cell(rr, 4).value = v
        ws.cell(rr, 4).font = font(size=10, color=INK)
        ws.cell(rr, 4).alignment = align("left", "center")
        for c in range(2, 8):
            ws.cell(rr, c).border = thin_border(LINE_SOFT)
            ws.cell(rr, c).fill = fill(bg)

    r = 21
    ws.row_dimensions[r].height = 12
    # KPI row 22-24
    kpi_box(ws, 22, 2, "ОСНОВНОЙ ОБЪЁМ", f"{money(totals['main'])} ₽", "разделы 1–7")
    kpi_box(ws, 22, 4, "РЕКОМЕНДУЕМЫЕ", f"{money(totals['rec'])} ₽", "раздел 8, по согласованию")
    kpi_box(ws, 22, 6, "ИТОГО С РЕКОМЕНДАЦИЯМИ", f"{money(totals['grand'])} ₽", "работы, без материалов")

    r = 26
    ws.merge_cells(f"B{r}:G{r}")
    ws[f"B{r}"] = rub_words(int(totals["grand"]))
    ws[f"B{r}"].font = font(size=10, italic=True, color=STEEL)
    ws[f"B{r}"].alignment = align("left", "center")
    ws[f"B{r}"].fill = fill(CREAM)
    ws.row_dimensions[r].height = 22

    r = 28
    ws.merge_cells(f"B{r}:G{r}")
    ws[f"B{r}"] = "Что делаем  ·  и чего не делаем"
    ws[f"B{r}"].font = font(size=14, bold=True, color=NAVY)
    ws[f"B{r}"].fill = fill(CREAM)
    ws.row_dimensions[r].height = 24

    include = [
        "Демонтаж унитаза и ежедневный временный унитаз на период проживания",
        "Вынос стиральной машины и чугунной ванны (узкий проём — резка при необходимости)",
        "Срезка старой внутриквартирной разводки воды; стояки не трогаем",
        "Коллекторы ХВС/ГВС, фильтры стандарт 10\", редукторы давления",
        "Полипропилен на раковину, унитаз, стиральную машину и ванну + водорозетки",
        "Инсталляция, переход с чугунного отвода на пластик без замены стояка",
        "Гидроизоляция пола и стен до уровня ванны, усиленная зона душа",
        "Керамогранит на пол, плитка на стены, подбивка порогов и плитка под двери",
        "Подиум и установка тяжёлой ванны из наливного камня, смесители, душ, раковина, унитаз",
    ]
    exclude = [
        "Демонтаж старой плитки — не закладывался",
        "Потолок: натяжной, реечный, покраска — не входит",
        "Стояки холодной и горячей воды — без замены и сварки",
        "Канализационный стояк — без замены; только отвод к инсталляции",
        "Расширение дверных проёмов и новые двери — не входят",
        "Материалы, сантехника, плитка, доставка и подъём на этаж — отдельно",
        "Замена полотенцесушителя на стояке ГВС — не входит",
        "Тёплый пол электрический — не входит (можно добавить)",
        "Дизайн-проект и согласование перепланировки — не входят",
    ]

    ws.merge_cells("B30:D30")
    ws["B30"] = "ВХОДИТ В РАБОТЫ"
    ws["B30"].font = font(size=10, bold=True, color=WHITE)
    ws["B30"].fill = fill(GREEN)
    ws["B30"].alignment = align("center")
    paint_range(ws, 30, 2, 4, fill(GREEN), height=18)
    ws["C30"].fill = fill(GREEN)
    ws["D30"].fill = fill(GREEN)

    ws.merge_cells("E30:G30")
    ws["E30"] = "НЕ ВХОДИТ"
    ws["E30"].font = font(size=10, bold=True, color=WHITE)
    ws["E30"].alignment = align("center")
    paint_range(ws, 30, 5, 7, fill(TERRACOTTA), height=18)

    for i in range(9):
        rr = 31 + i
        ws.merge_cells(f"B{rr}:D{rr}")
        ws.merge_cells(f"E{rr}:G{rr}")
        bg = WHITE if i % 2 == 0 else PALE_BLUE
        paint_range(ws, rr, 2, 4, fill(bg), height=32)
        paint_range(ws, rr, 5, 7, fill(SAND if i % 2 == 0 else WHITE), height=32)
        ws.cell(rr, 2).value = f"  {i + 1}.  {include[i]}"
        ws.cell(rr, 2).font = font(size=8, color=INK)
        ws.cell(rr, 2).alignment = align("left", "center")
        ws.cell(rr, 5).value = f"  {i + 1}.  {exclude[i]}"
        ws.cell(rr, 5).font = font(size=8, color=INK)
        ws.cell(rr, 5).alignment = align("left", "center")

    rr = 41
    paint_range(ws, rr, 1, 8, fill(NAVY), height=8)
    ws.merge_cells("B42:G42")
    ws["B42"] = (
        "Расценки — рынок Санкт-Петербурга, 2026, работы в стеснённых санузлах сантехкабины. "
        "Итог уточняется после замера. Смета не оферта."
    )
    ws["B42"].font = font(size=8, italic=True, color=MUTED)
    ws["B42"].alignment = align("left", "center", wrap=True)
    ws["B42"].fill = fill(CREAM)
    ws.row_dimensions[42].height = 28

    ws.merge_cells("B44:C44")
    ws["B44"] = "Заказчик ______________________"
    ws["B44"].font = font(size=10, color=NAVY)
    ws["B44"].fill = fill(CREAM)
    ws.merge_cells("E44:G44")
    ws["E44"] = "Исполнитель ______________________"
    ws["E44"].font = font(size=10, color=NAVY)
    ws["E44"].fill = fill(CREAM)

    ws.merge_cells("B45:C45")
    ws["B45"] = "подпись / расшифровка"
    ws["B45"].font = font(size=8, italic=True, color=MUTED)
    ws["B45"].fill = fill(CREAM)
    ws.merge_cells("E45:G45")
    ws["E45"] = "подпись / расшифровка"
    ws["E45"].font = font(size=8, italic=True, color=MUTED)
    ws["E45"].fill = fill(CREAM)

    ws.print_area = "A1:H46"
    ws.page_setup.fitToHeight = 1
    return ws


def build_estimate(wb: Workbook) -> dict:
    ws = wb.create_sheet("Смета")
    apply_print(ws, landscape=True, title="Локальная смета")
    set_col_widths(
        ws,
        {
            "A": 8,
            "B": 78,
            "C": 10,
            "D": 12,
            "E": 14,
            "F": 16,
            "G": 42,
        },
    )
    ws.freeze_panes = "A8"
    ws.page_setup.fitToHeight = 0
    ws.oddHeader.center.text = "&K1B365DРемонт ванной и туалета · ЛГ-137 · только работы"
    ws.oddHeader.left.text = ""
    ws.oddHeader.right.text = "&K5C6670&D"

    # banner
    for r in (1, 2, 3):
        paint_range(ws, r, 1, 7, fill(NAVY), height=16)
    ws.merge_cells("A1:G1")
    ws["A1"] = "ЛОКАЛЬНАЯ СМЕТА НА РАБОТЫ"
    ws["A1"].font = font(size=16, bold=True, color=WHITE)
    ws["A1"].alignment = align("left", "center")
    ws["A1"].fill = fill(NAVY)
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:G2")
    ws["A2"] = "Ванная и туалет · дом серии ЛГ-137 · проживание на объекте · потолок и стояки не входят · плитку не сбиваем"
    ws["A2"].font = font(size=10, color=GOLD_SOFT)
    ws["A2"].fill = fill(NAVY)

    ws.merge_cells("A3:G3")
    ws["A3"] = "Колонки D и E можно править — сумма в F пересчитается. Раздел 8 не входит в основной итог, пока его не включат в договор."
    ws["A3"].font = font(size=9, italic=True, color=GOLD)
    ws["A3"].fill = fill(NAVY)

    paint_range(ws, 4, 1, 7, fill(GOLD), height=5)

    ws.merge_cells("A5:G5")
    ws["A5"] = (
        f"Исходные объёмы: ванна {BATH_L:.2f}×{BATH_W:.2f} м ({BATH_FLOOR} м²), "
        f"туалет {WC_L:.2f}×{WC_W:.2f} м ({WC_FLOOR} м²), высота {H:.2f} м. "
        f"Стены под плитку: ванна {BATH_WALLS} м², туалет+короб {WC_TILE_WALLS} м². "
        "Расчёт — лист «Объёмы»."
    )
    ws["A5"].font = font(size=9, color=MUTED)
    ws["A5"].fill = fill(SAND)
    ws["A5"].alignment = align("left", "center")
    paint_range(ws, 5, 1, 7, fill(SAND), height=32)

    headers = ["№", "Наименование работ", "Ед.", "Кол-во", "Цена, ₽", "Сумма, ₽", "Примечание"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(7, c, h)
        cell.font = font(size=10, bold=True, color=WHITE)
        cell.fill = fill(STEEL)
        cell.alignment = align("center", "center")
        cell.border = thin_border(STEEL)
    ws.row_dimensions[7].height = 22

    row = 8
    section_total_cells = []  # (title, cell, recommended)
    item_rows_main = []
    item_rows_rec = []

    for sec in SECTIONS:
        rec = bool(sec.get("recommended"))
        # section header
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        bg = TERRACOTTA if rec else NAVY
        paint_range(ws, row, 1, 7, fill(bg), height=22)
        ws.cell(row, 1).value = f"  РАЗДЕЛ {sec['num']}.  {sec['title'].upper()}"
        ws.cell(row, 1).font = font(size=11, bold=True, color=WHITE)
        ws.cell(row, 1).alignment = align("left", "center")
        for c in range(1, 8):
            ws.cell(row, c).fill = fill(bg)
        row += 1

        # intro
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        paint_range(ws, row, 1, 7, fill(REC_BG if rec else PALE_BLUE), height=28)
        ws.cell(row, 1).value = sec["intro"]
        ws.cell(row, 1).font = font(size=9, italic=True, color=STEEL)
        ws.cell(row, 1).alignment = align("left", "center")
        for c in range(1, 8):
            ws.cell(row, c).fill = fill(REC_BG if rec else PALE_BLUE)
        row += 1

        first_item = row
        for i, it in enumerate(sec["items"]):
            bg = (REC_BG if rec else ROW_ALT) if i % 2 else WHITE
            ws.cell(row, 1, it["code"]).alignment = align("center")
            ws.cell(row, 2, it["name"]).alignment = align("left", "center")
            ws.cell(row, 3, it["unit"]).alignment = align("center")
            ws.cell(row, 4, it["qty"]).alignment = align("center")
            ws.cell(row, 5, it["price"]).alignment = align("right")
            ws.cell(row, 6, f"=ROUND(D{row}*E{row},0)").alignment = align("right")
            ws.cell(row, 7, it["note"]).alignment = align("left", "center")
            ws.cell(row, 4).number_format = "0.00" if isinstance(it["qty"], float) else "0"
            ws.cell(row, 5).number_format = '#,##0'
            ws.cell(row, 6).number_format = '#,##0'
            for c in range(1, 8):
                cell = ws.cell(row, c)
                cell.font = font(size=10, color=INK)
                cell.fill = fill(bg)
                cell.border = thin_border(LINE_SOFT)
            ws.cell(row, 1).font = font(size=10, bold=True, color=STEEL)
            ws.row_dimensions[row].height = 32
            if rec:
                item_rows_rec.append(row)
            else:
                item_rows_main.append(row)
            row += 1
        last_item = row - 1

        # subtotal
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
        paint_range(ws, row, 1, 7, fill(SUBTOTAL_BG), height=22)
        ws.cell(row, 1).value = f"Итого по разделу {sec['num']}"
        ws.cell(row, 1).font = font(size=10, bold=True, color=NAVY)
        ws.cell(row, 1).alignment = align("right", "center")
        ws.cell(row, 6).value = f"=SUM(F{first_item}:F{last_item})"
        ws.cell(row, 6).font = font(size=11, bold=True, color=NAVY)
        ws.cell(row, 6).number_format = '#,##0'
        ws.cell(row, 6).alignment = align("right", "center")
        for c in range(1, 8):
            ws.cell(row, c).fill = fill(SUBTOTAL_BG)
            ws.cell(row, c).border = thin_border(LINE)
        section_total_cells.append(
            {
                "num": sec["num"],
                "title": sec["title"],
                "cell": f"F{row}",
                "row": row,
                "recommended": rec,
            }
        )
        row += 1
        # spacer
        paint_range(ws, row, 1, 7, fill(CREAM), height=10)
        row += 1

    # totals block
    totals_start = row
    paint_range(ws, row, 1, 7, fill(GOLD), height=6)
    row += 1

    main_cells = [s["cell"] for s in section_total_cells if not s["recommended"]]
    rec_cells = [s["cell"] for s in section_total_cells if s["recommended"]]

    def total_row(r, label, formula, bg, fg, size=12, height=24):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        paint_range(ws, r, 1, 7, fill(bg), height=height)
        ws.cell(r, 1).value = label
        ws.cell(r, 1).font = font(size=size, bold=True, color=fg)
        ws.cell(r, 1).alignment = align("right", "center")
        ws.cell(r, 6).value = formula
        ws.cell(r, 6).font = font(size=size, bold=True, color=fg)
        ws.cell(r, 6).number_format = '#,##0'
        ws.cell(r, 6).alignment = align("right", "center")
        for c in range(1, 8):
            ws.cell(r, c).fill = fill(bg)

    main_formula = "=" + "+".join(main_cells)
    rec_formula = "=" + "+".join(rec_cells)

    total_row(row, "ИТОГО РАБОТЫ ПО ЗАДАНИЮ  (разделы 1–7)", main_formula, NAVY, WHITE, 13, 28)
    main_total_cell = f"F{row}"
    row += 1
    total_row(
        row,
        "Резерв на непредвиденные работы 5%  (по согласованию)",
        f"=ROUND({main_total_cell}*0.05,0)",
        STEEL,
        WHITE,
        11,
        22,
    )
    reserve_cell = f"F{row}"
    row += 1
    total_row(
        row,
        "Рекомендуемые работы  (раздел 8)",
        rec_formula,
        TERRACOTTA,
        WHITE,
        11,
        22,
    )
    rec_total_cell = f"F{row}"
    row += 1
    total_row(
        row,
        "ВСЕГО К ДОГОВОРУ  (задание + резерв 5% + рекомендации)",
        f"={main_total_cell}+{reserve_cell}+{rec_total_cell}",
        GOLD,
        NAVY_DEEP,
        14,
        32,
    )
    grand_cell = f"F{row}"
    row += 1

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    ws.cell(row, 1).value = (
        "Материалы, сантехнические приборы, плитка, клей, гидроизоляция, инсталляция, ванна, смесители и доставка в итог не входят. "
        "Расценки учитывают стеснённость сантехкабины ЛГ-137 и работу в жилой квартире."
    )
    ws.cell(row, 1).font = font(size=9, italic=True, color=MUTED)
    ws.cell(row, 1).alignment = align("left", "center")
    paint_range(ws, row, 1, 7, fill(SAND), height=36)
    note_row = row
    row += 2

    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
    ws.cell(row, 1).value = "Составил ________________ / дата"
    ws.cell(row, 1).font = font(size=10, color=NAVY)
    ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=7)
    ws.cell(row, 5).value = "Заказчик с объёмом и расценками ознакомлен ________________"
    ws.cell(row, 5).font = font(size=10, color=NAVY)

    last_row = row
    ws.print_area = f"A1:G{last_row}"
    ws.page_setup.printTitleRows = "1:7"
    ws.oddFooter.left.text = "&K5C6670Локальная смета · только работы"

    wb.defined_names.add(DefinedName(name="ITOGO_ZADANIE", attr_text=f"Смета!{main_total_cell}"))
    wb.defined_names.add(DefinedName(name="ITOGO_REZERV", attr_text=f"Смета!{reserve_cell}"))
    wb.defined_names.add(DefinedName(name="ITOGO_REC", attr_text=f"Смета!{rec_total_cell}"))
    wb.defined_names.add(DefinedName(name="ITOGO_VSEGO", attr_text=f"Смета!{grand_cell}"))

    # compute python totals for cover (formulas evaluated independently)
    def sec_sum(sec):
        return sum(round_rub(it["qty"] * it["price"]) for it in sec["items"])

    main = sum(sec_sum(s) for s in SECTIONS if not s.get("recommended"))
    rec = sum(sec_sum(s) for s in SECTIONS if s.get("recommended"))
    reserve = round(main * 0.05)
    grand = main + reserve + rec

    totals = {
        "main": main,
        "rec": rec,
        "reserve": reserve,
        "grand": grand,
        "main_cell": main_total_cell,
        "rec_cell": rec_total_cell,
        "reserve_cell": reserve_cell,
        "grand_cell": grand_cell,
        "section_total_cells": section_total_cells,
        "last_row": last_row,
        "sections": [
            {
                "num": s["num"],
                "title": s["title"],
                "sum": sec_sum(s),
                "recommended": bool(s.get("recommended")),
                "cell": section_total_cells[i]["cell"],
            }
            for i, s in enumerate(SECTIONS)
        ],
    }
    return totals


def build_summary(wb: Workbook, totals: dict):
    ws = wb.create_sheet("Сводка", 1)
    apply_print(ws, landscape=True, title="Сводка")
    set_col_widths(ws, {"A": 4, "B": 8, "C": 62, "D": 16, "E": 14, "F": 18, "G": 18, "H": 4})
    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToHeight = 1

    for r in range(1, 40):
        for c in range(1, 9):
            ws.cell(r, c).fill = fill(CREAM)

    paint_range(ws, 1, 1, 8, fill(NAVY), height=26)
    ws.merge_cells("B1:G1")
    ws["B1"] = "СВОДКА ПО РАЗДЕЛАМ"
    ws["B1"].font = font(size=16, bold=True, color=WHITE)
    ws["B1"].fill = fill(NAVY)
    ws["B1"].alignment = align("left", "center")
    paint_range(ws, 2, 1, 8, fill(GOLD), height=5)

    ws.merge_cells("B3:G3")
    ws["B3"] = "Суммы связаны формулами с листом «Смета». Меняете количество или цену там — сводка обновляется."
    ws["B3"].font = font(size=9, italic=True, color=MUTED)
    ws["B3"].fill = fill(CREAM)
    ws.row_dimensions[3].height = 18

    headers = ["", "№", "Раздел", "Сумма, ₽", "Доля", "Статус", ""]
    # place starting col B
    for i, h in enumerate(["№", "Раздел", "Сумма, ₽", "Доля основного", "Статус"], 2):
        cell = ws.cell(5, i, h)
        cell.font = font(size=10, bold=True, color=WHITE)
        cell.fill = fill(STEEL)
        cell.alignment = align("center")
        cell.border = thin_border(STEEL)
    ws.row_dimensions[5].height = 20

    main_cell_ref = f"Смета!{totals['main_cell']}"
    start = 6
    for i, s in enumerate(totals["sections"]):
        r = start + i
        rec = s["recommended"]
        bg = REC_BG if rec else (WHITE if i % 2 == 0 else ROW_ALT)
        ws.cell(r, 2, s["num"]).alignment = align("center")
        ws.cell(r, 3, s["title"]).alignment = align("left", "center")
        ws.cell(r, 4, f"=Смета!{s['cell']}").number_format = '#,##0'
        ws.cell(r, 4).alignment = align("right")
        if rec:
            ws.cell(r, 5, "—")
            ws.cell(r, 6, "по согласованию")
        else:
            ws.cell(r, 5, f"=IF({main_cell_ref}=0,0,Смета!{s['cell']}/{main_cell_ref})")
            ws.cell(r, 5).number_format = "0.0%"
            ws.cell(r, 6, "входит в задание")
        ws.cell(r, 5).alignment = align("center")
        ws.cell(r, 6).alignment = align("center")
        for c in range(2, 7):
            ws.cell(r, c).font = font(size=10, color=INK)
            ws.cell(r, c).fill = fill(bg)
            ws.cell(r, c).border = thin_border(LINE_SOFT)
        ws.cell(r, 2).font = font(size=10, bold=True, color=NAVY)
        ws.row_dimensions[r].height = 22

    end_sec = start + len(totals["sections"]) - 1
    r = end_sec + 2
    rows_fin = [
        ("Итого по заданию (1–7)", f"=Смета!{totals['main_cell']}", NAVY, WHITE),
        ("Резерв 5%", f"=Смета!{totals['reserve_cell']}", STEEL, WHITE),
        ("Рекомендуемые (8)", f"=Смета!{totals['rec_cell']}", TERRACOTTA, WHITE),
        ("Всего к договору", f"=Смета!{totals['grand_cell']}", GOLD, NAVY_DEEP),
    ]
    for label, formula, bg, fg in rows_fin:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.cell(r, 2).value = label
        ws.cell(r, 2).font = font(size=11, bold=True, color=fg)
        ws.cell(r, 2).alignment = align("right", "center")
        ws.cell(r, 4).value = formula
        ws.cell(r, 4).number_format = '#,##0'
        ws.cell(r, 4).font = font(size=12, bold=True, color=fg)
        ws.cell(r, 4).alignment = align("right", "center")
        for c in range(2, 7):
            ws.cell(r, c).fill = fill(bg)
            if c != 4:
                ws.cell(r, c).font = font(size=11, bold=True, color=fg)
        ws.row_dimensions[r].height = 24
        r += 1

    # chart data on the right is from B6:D12 (main sections only, 7 rows)
    pie = PieChart()
    pie.title = "Структура основного объёма"
    labels = Reference(ws, min_col=3, min_row=6, max_row=12)
    data = Reference(ws, min_col=4, min_row=5, max_row=12)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showCatName = False
    pie.dataLabels.showSerName = False
    pie.width = 12
    pie.height = 8
    pie.style = 10
    ws.add_chart(pie, "A22")

    bar = BarChart()
    bar.type = "col"
    bar.title = "Сумма по разделам, ₽"
    bar.y_axis.title = None
    bar.x_axis.title = None
    data2 = Reference(ws, min_col=4, min_row=5, max_row=13)
    cats2 = Reference(ws, min_col=2, min_row=6, max_row=13)
    bar.add_data(data2, titles_from_data=True)
    bar.set_categories(cats2)
    bar.shape = 4
    bar.style = 10
    bar.legend = None
    bar.width = 18
    bar.height = 8
    ws.add_chart(bar, "E22")

    ws.merge_cells("B18:F18")
    ws["B18"] = (
        f"Ориентир срока: {WORK_DAYS} рабочих дней. "
        "На время работ унитаз ставится каждый вечер. "
        "Чугунную ванну из проёма 57 см обычно выносят частями. "
        "Новую ванну из наливного камня заносят 3–4 человека, при необходимости — через окно."
    )
    ws["B18"].font = font(size=9, color=STEEL)
    ws["B18"].alignment = align("left", "center", wrap=True)
    ws["B18"].fill = fill(SAND)
    paint_range(ws, 18, 2, 6, fill(SAND), height=40)

    ws.print_area = "A1:H36"
    ws.page_setup.fitToHeight = 1
    return ws


def build_volumes(wb: Workbook):
    ws = wb.create_sheet("Объёмы")
    apply_print(ws, landscape=False, title="Объёмы")
    set_col_widths(ws, {"A": 4, "B": 42, "C": 16, "D": 14, "E": 14, "F": 14, "G": 36, "H": 4})
    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToHeight = 1

    for r in range(1, 46):
        for c in range(1, 9):
            ws.cell(r, c).fill = fill(CREAM)

    paint_range(ws, 1, 1, 8, fill(NAVY), height=26)
    ws.merge_cells("B1:G1")
    ws["B1"] = "РАСЧЁТ ОБЪЁМОВ  ·  ТИПОВАЯ САНТЕХКАБИНА ЛГ-137"
    ws["B1"].font = font(size=14, bold=True, color=WHITE)
    ws["B1"].fill = fill(NAVY)
    paint_range(ws, 2, 1, 8, fill(GOLD), height=5)

    ws.merge_cells("B3:G3")
    ws["B3"] = (
        "Габариты приняты по типовой сантехкабине: ванна 1700×1500 мм, туалет 1200×820 мм, "
        "высота 2550 мм, двери 570×1980 мм. После замера объёмы на листе «Смета» правятся вручную."
    )
    ws["B3"].font = font(size=9, italic=True, color=MUTED)
    ws["B3"].alignment = align("left", "center", wrap=True)
    ws.row_dimensions[3].height = 36

    headers = ["Параметр", "Формула / основание", "Значение", "Ед.", "Куда идёт", "Примечание"]
    for i, h in enumerate(headers, 2):
        cell = ws.cell(5, i, h)
        cell.font = font(size=9, bold=True, color=WHITE)
        cell.fill = fill(STEEL)
        cell.alignment = align("center", "center")
    ws.row_dimensions[5].height = 20

    rows = [
        ("Длина ванной", "типовой ЛГ-137", BATH_L, "м", "стены, пол", "макс. 1,73 м"),
        ("Ширина ванной", "типовой ЛГ-137", BATH_W, "м", "стены, пол", "ванна поперёк"),
        ("Длина туалета", "типовой ЛГ-137", WC_L, "м", "стены, пол", "за унитазом вентшахта"),
        ("Ширина туалета", "типовой ЛГ-137", WC_W, "м", "стены, пол", "0,80–0,82 м"),
        ("Высота помещений", "сантехкабина", H, "м", "стены", "потолок не входит в работы"),
        ("Дверной проём", "2 шт.", f"{DOOR_W}×{DOOR_H}", "м", "вычет из стен, пороги", "нестандарт 57 см"),
        ("Пол ванной", "1,70 × 1,50", BATH_FLOOR, "м²", "п. 5.1–5.3, 6.4", ""),
        ("Пол туалета", "1,20 × 0,82", WC_FLOOR, "м²", "п. 5.1–5.3, 6.5", ""),
        ("Пол всего", "2,55 + 0,98", FLOOR, "м²", "гидроизоляция, стяжка", ""),
        ("Стены ванной", "периметр × H − дверь", BATH_WALLS, "м²", "п. 5.4–5.5, 6.6", "до существующего потолка"),
        ("Стены туалета", "периметр × H − дверь", WC_WALLS, "м²", "п. 5.4–5.5", ""),
        ("Стены туалет + короб", "9,17 + 1,50", WC_TILE_WALLS, "м²", "п. 6.7", "плоскость инсталляции"),
        ("Гидроизоляция стен до ванны", "(P − дверь) × 0,70", HYDRO_WALLS, "м²", "п. 5.6", "ванная и туалет"),
        ("Зона душа над ванной", "(1,70+0,75) × 1,30", SHOWER_ZONE, "м²", "п. 5.7", "от 0,70 до 2,00 м"),
        ("Затирка", "пол + стены с коробом", GROUT, "м²", "п. 6.9", ""),
        ("Полипропилен ХВС+ГВС", "трасса коллектор → 4 прибора", 32, "м.п.", "п. 3.5", "с запасом на подводы"),
        ("Канализация Ø50", "ванна, раковина, стиралка", 8, "м.п.", "п. 4.1", ""),
        ("Канализация Ø110", "выпуск стояка → инсталляция", 1.5, "м.п.", "п. 4.2", "стояк не меняем"),
        ("Водорозетки", "ванна 2 + раковина 2 + СМ 1 + унитаз 1", 6, "шт.", "п. 3.6–3.9", ""),
        ("Пороги", "два проёма", 2, "шт.", "п. 6.1–6.3, 6.8", "подбивка + плитка под дверь"),
        ("Циклы временного унитаза", "вечерний монтаж + утренний демонтаж", WORK_DAYS, "цикл", "п. 1.3", "пока нет чаши на инсталляции"),
    ]

    for i, (a, b, c, d, e, f) in enumerate(rows):
        r = 6 + i
        bg = WHITE if i % 2 == 0 else ROW_ALT
        vals = (a, b, c, d, e, f)
        for col, val in enumerate(vals, 2):
            cell = ws.cell(r, col, val)
            cell.fill = fill(bg)
            cell.font = font(size=9, color=INK)
            cell.border = thin_border(LINE_SOFT)
            cell.alignment = align("center" if col in (4, 5) else "left", "center")
            if col == 4 and isinstance(val, (int, float)):
                cell.number_format = "0.00"
        ws.row_dimensions[r].height = 20

    r = 29
    ws.merge_cells(f"B{r}:G{r}")
    ws[f"B{r}"] = "Точки водоразбора"
    ws[f"B{r}"].font = font(size=12, bold=True, color=NAVY)
    ws[f"B{r}"].fill = fill(CREAM)

    headers2 = ["Прибор", "ХВС", "ГВС", "Канализация", "Водорозетки", "Комментарий"]
    for i, h in enumerate(headers2, 2):
        cell = ws.cell(30, i, h)
        cell.font = font(size=9, bold=True, color=WHITE)
        cell.fill = fill(STEEL)
        cell.alignment = align("center")

    points = [
        ("Ванна + душ", "да", "да", "Ø50", "2", "Смеситель настенный, душ над ванной"),
        ("Раковина", "да", "да", "Ø50", "2", "Смеситель на раковину или настенный"),
        ("Стиральная машина", "да", "нет", "Ø50", "1", "Кран на водорозетке"),
        ("Унитаз (инсталляция)", "да", "нет", "Ø110", "1", "Подвод к раме, стояк не трогаем"),
    ]
    for i, rowv in enumerate(points):
        rr = 31 + i
        bg = WHITE if i % 2 == 0 else PALE_BLUE
        for col, val in enumerate(rowv, 2):
            cell = ws.cell(rr, col, val)
            cell.fill = fill(bg)
            cell.font = font(size=9, color=INK)
            cell.border = thin_border(LINE_SOFT)
            cell.alignment = align("center" if col != 2 else "left", "center")
        ws.row_dimensions[rr].height = 22

    r = 36
    ws.merge_cells(f"B{r}:G{r}")
    ws[f"B{r}"] = (
        "Схема коллектора (каждый стояк): отвод со стояка → шаровый кран → счётчик (существующий) → "
        "фильтр стандарт 10\" → редуктор давления → обратный клапан → коллектор → отводы на точки. "
        "Стояк и полотенцесушитель на стояке ГВС не переделываем."
    )
    ws[f"B{r}"].font = font(size=9, color=STEEL)
    ws[f"B{r}"].alignment = align("left", "center", wrap=True)
    paint_range(ws, r, 2, 7, fill(SAND), height=48)

    ws.print_area = "A1:H38"
    return ws


def build_conditions(wb: Workbook):
    ws = wb.create_sheet("Условия")
    apply_print(ws, landscape=False, title="Условия")
    set_col_widths(ws, {"A": 4, "B": 8, "C": 78, "D": 4})
    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToHeight = 1

    for r in range(1, 72):
        ws.cell(r, 1).fill = fill(CREAM)
        ws.cell(r, 2).fill = fill(CREAM)
        ws.cell(r, 3).fill = fill(CREAM)
        ws.cell(r, 4).fill = fill(CREAM)

    paint_range(ws, 1, 1, 4, fill(NAVY), height=26)
    ws.merge_cells("B1:C1")
    ws["B1"] = "СОСТАВ РАБОТ, ПОРЯДОК И УСЛОВИЯ ДЛЯ ЗАКАЗЧИКА"
    ws["B1"].font = font(size=14, bold=True, color=WHITE)
    ws["B1"].fill = fill(NAVY)
    paint_range(ws, 2, 1, 4, fill(GOLD), height=5)

    blocks = [
        (
            "1. Как устроен объект",
            [
                "Дом серии ЛГ-137, раздельный санузел — готовая сантехкабина: тонкие армированные стены, пол кабины ~50 мм, проёмы дверей около 570×1980 мм.",
                "Ванна стоит поперёк помещения 1700×1500 мм. Туалет 1200×820 мм: за унитазом стояки, канализация и вентшахта — места мало.",
                "Чугунный унитаз обычно зачеканен в раструб стояка и стоит на деревянной закладной в полу. Закладную при ремонте вынимаем и заделываем.",
            ],
        ),
        (
            "2. Что сознательно не делаем",
            [
                "Не сбиваем плитку. Готовим существующее основание: грунт / бетоноконтакт и выравнивание под новую облицовку.",
                "Не трогаем потолок: нет демонтажа, покраски, натяжного или реечного потолка, потолочных светильников.",
                "Не меняем стояки холодной и горячей воды и канализационный стояк. Режем только внутриквартирную разводку от отводов.",
                "Не расширяем дверные проёмы и не ставим новые двери. Только порог: подбить и положить плитку под полотно.",
            ],
        ),
        (
            "3. Проживание во время ремонта",
            [
                "Стиральную машину в первый день отключаем и выносим. После чистовых работ возвращаем и подключаем.",
                "Унитаз снимаем. Каждый вечер ставим временный (ваш или технический) на выпуск канализации, каждое утро снимаем.",
                "В смете 18 циклов — на период, пока не навешена чаша на инсталляцию (инсталляцию ставим рано, чашу — после плитки).",
                "Воду по квартире полностью не отключаем: перекрываются отводы. Короткие отключения — по согласованию, обычно утром.",
                "Коридор защищаем плёнкой. В конце каждой смены — уборка зоны работ.",
            ],
        ),
        (
            "4. Порядок работ (логика, не график поставки)",
            [
                "Защита квартиры → вынос стиральной машины и раковины → демонтаж унитаза и постановка временного.",
                "Демонтаж чугунной ванны (через проём 57 см чашу обычно режут) и срезка старых труб.",
                "Расчеканка чугунного отвода, переход на пластик, коллекторы с фильтрами 10\" и редукторами, полипропилен, водорозетки, инсталляция.",
                "Опрессовка. Стяжка и гидроизоляция пола, гидроизоляция стен до ванны и зоны душа, выравнивание стен.",
                "Подбивка порогов. Керамогранит на пол, плитка на стены и на пороги под двери, экран ванны, затирка.",
                "Подиум и занос ванны из наливного камня, смесители, душ, раковина, навеска унитаза, подключение стиральной машины.",
            ],
        ),
        (
            "5. Ванна из наливного камня — отдельно",
            [
                "Чаша тяжёлая (часто 120–200+ кг). Ставить только на заводские ножки нельзя: делаем опорный подиум / ложемент по контуру.",
                "Проём 57 см. Если ванна 70–75 см в ширину, занос через дверь может быть невозможен — тогда через окно (оговаривается на замере, такелаж может быть отдельно).",
                "Душ принимаем над ванной: стойка и лейка. Душевая кабина в эти габариты не закладывалась.",
            ],
        ),
        (
            "6. Инсталляция в туалете ЛГ-137",
            [
                "Подвесной унитаз даёт +5–10 см до двери — в этом туалете это важно.",
                "Чтобы опустить выпуск, чугунный отвод расчеканиваем и ставим пластиковый переход. Стояк при этом не меняем. Работа деликатная: чугун хрупкий.",
                "После облицовки короба — кнопка и чаша. До этого момента действует режим временного унитаза.",
            ],
        ),
        (
            "7. Что заказчик закупает сам (работы без материалов)",
            [
                "Трубы ПП и фитинги, коллекторы, корпуса фильтров 10\" и картриджи, редукторы, краны, водорозетки, канализационные трубы.",
                "Инсталляция с кнопкой и унитаз, раковина, смесители, душевой комплект, сифоны, ванна из наливного камня, слив-перелив.",
                "Грунты, штукатурка, стяжка, обмазочная гидроизоляция, ленты, клей, затирка, силикон, ГВЛ, люки.",
                "Керамогранит на пол, плитка на стены, раскладки, люки ревизии. Запас плитки 10–15% из-за подрезки в маленьких помещениях.",
            ],
        ),
        (
            "8. Рекомендуемый блок (раздел 8)",
            [
                "Короб стояков с люком — иначе коллекторы и фильтры остаются открытыми.",
                "Вентилятор в существующий канал: окна плотные, естественной тяги после ремонта часто не хватает. Потолок не разбираем.",
                "Розетки IP44 и УЗО на стиральную машину — норма для мокрой зоны, в домах до 2000-х заземления в кабине часто нет.",
                "Стеклянная шторка — чтобы душ над ванной не заливал пол у порога.",
            ],
        ),
        (
            "9. Риски, которые лучше знать заранее",
            [
                "При демонтаже унитаза можно повредить чугунный раструб стояка. В смете заложена аккуратная расчеканка; ремонт самого стояка — вне объёма.",
                "Стены кабины тонкие и армированные, отверстия под водорозетки могут быть сквозными — крепим площадки, не «на дюбель в пустоту».",
                "Если старая плитка держится плохо, выравнивание без демонтажа может не получиться. Тогда сбивка плитки — доп. соглашение (сейчас не входит).",
                "Расценки — ориентир рынка СПб 2026 для стеснённых санузлов. Итог фиксируется после замера и выбора сантехники.",
            ],
        ),
        (
            "10. Оплата и гарантия (типовые условия, вносятся в договор)",
            [
                "Смета — приложение к договору. Цены на работы, не на материалы.",
                "Типовой график: 40% аванс на выход бригады, 40% после скрытых работ и опрессовки, 20% по акту сдачи.",
                "Гарантия на работы: 24 месяца на сантехническую разводку и гидроизоляцию, 12 месяцев на облицовку и монтаж приборов — при эксплуатации по нормам.",
                "Скрытые работы фотофиксируем до закрытия. Опрессовка — до облицовки коробов.",
            ],
        ),
    ]

    r = 4
    for title, lines in blocks:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        paint_range(ws, r, 2, 3, fill(NAVY), height=20)
        ws.cell(r, 2).value = f"  {title}"
        ws.cell(r, 2).font = font(size=11, bold=True, color=WHITE)
        ws.cell(r, 2).fill = fill(NAVY)
        ws.cell(r, 3).fill = fill(NAVY)
        r += 1
        for i, line in enumerate(lines):
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            bg = WHITE if i % 2 == 0 else ROW_ALT
            paint_range(ws, r, 2, 3, fill(bg), height=36)
            ws.cell(r, 2).value = f"   {line}"
            ws.cell(r, 2).font = font(size=9, color=INK)
            ws.cell(r, 2).alignment = align("left", "center", wrap=True)
            ws.cell(r, 2).fill = fill(bg)
            ws.cell(r, 3).fill = fill(bg)
            ws.cell(r, 2).border = thin_border(LINE_SOFT)
            ws.cell(r, 3).border = thin_border(LINE_SOFT)
            r += 1
        r += 1

    ws.print_area = f"A1:D{r}"
    return ws


def build_materials(wb: Workbook):
    """Чек-лист закупки заказчика: в смету работ не входит."""
    ws = wb.create_sheet("Закупка заказчика")
    apply_print(ws, landscape=False, title="Закупка заказчика")
    set_col_widths(ws, {"A": 4, "B": 10, "C": 52, "D": 22, "E": 28, "F": 12, "G": 4})
    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToHeight = 1

    for r in range(1, 70):
        for c in range(1, 8):
            ws.cell(r, c).fill = fill(CREAM)

    paint_range(ws, 1, 1, 7, fill(NAVY), height=26)
    ws.merge_cells("B1:F1")
    ws["B1"] = "ЧЕК-ЛИСТ ЗАКУПКИ ЗАКАЗЧИКА  ·  В СМЕТУ РАБОТ НЕ ВХОДИТ"
    ws["B1"].font = font(size=14, bold=True, color=WHITE)
    ws["B1"].fill = fill(NAVY)
    paint_range(ws, 2, 1, 7, fill(GOLD), height=5)

    ws.merge_cells("B3:F3")
    ws["B3"] = (
        "Это не смета материалов и не цены. Список того, что нужно привезти на объект, "
        "чтобы бригада могла выполнить работы. Марку и модель выбираете вы (или вместе на замере)."
    )
    ws["B3"].font = font(size=9, italic=True, color=MUTED)
    ws["B3"].alignment = align("left", "center", wrap=True)
    ws.row_dimensions[3].height = 36

    headers = ["Код", "Позиция", "Количество / ориентир", "Для каких работ", "✓"]
    for i, h in enumerate(headers, 2):
        cell = ws.cell(5, i, h)
        cell.font = font(size=9, bold=True, color=WHITE)
        cell.fill = fill(STEEL)
        cell.alignment = align("center")
    ws.row_dimensions[5].height = 20

    groups = [
        (
            "Вода и канализация",
            [
                ("К-01", "Коллекторы ХВС и ГВС (по числу отводов, обычно 3+3 или 4+2)", "2 шт.", "п. 3.1–3.2"),
                ("К-02", "Фильтры стандарт 10\" (корпус + картридж) на ХВС и ГВС", "2 шт.", "п. 3.1–3.2"),
                ("К-03", "Редукторы давления ХВС и ГВС", "2 шт.", "п. 3.1–3.2"),
                ("К-04", "Шаровые краны, американки, обратные клапаны, прокладки", "комплект", "п. 3.3"),
                ("К-05", "Ревизионный шкаф / люк под коллекторы и фильтры", "1 шт.", "п. 3.4"),
                ("К-06", "Труба полипропилен PN20 + фитинги, клипсы, теплоизоляция ГВС", "≈ 32 м + фитинги", "п. 3.5"),
                ("К-07", "Водорозетки (внутренняя резьба 1/2\")", "6 шт.", "п. 3.6–3.9"),
                ("К-08", "Кран угловой на стиральную машину", "1 шт.", "п. 3.8"),
                ("К-09", "Канализация ПВХ Ø50, Ø110, отводы, манжеты на чугун", "Ø50 ≈ 8 м, Ø110 ≈ 1,5 м", "п. 4.1–4.3"),
            ],
        ),
        (
            "Инсталляция и сантехника",
            [
                ("С-01", "Инсталляция для подвесного унитаза + кнопка", "1 компл.", "п. 4.4–4.8"),
                ("С-02", "Унитаз подвесной совместимый с инсталляцией", "1 шт.", "п. 7.9"),
                ("С-03", "Ванна из наливного камня (проверить проход в проём 57 см / окно)", "1 шт.", "п. 7.1–7.3"),
                ("С-04", "Слив-перелив и сифон ванны под выбранную модель", "1 шт.", "п. 7.4"),
                ("С-05", "Смеситель на ванну", "1 шт.", "п. 7.5"),
                ("С-06", "Душевой комплект: стойка / лейка / шланг / держатель", "1 шт.", "п. 7.6"),
                ("С-07", "Раковина + крепёж + сифон", "1 шт.", "п. 7.7"),
                ("С-08", "Смеситель раковины", "1 шт.", "п. 7.8"),
                ("С-09", "Люк ревизии экрана ванны, люк кнопки инсталляции", "2 шт.", "п. 4.7, 6.11"),
                ("С-10", "Стеклянная шторка на ванну (если берёте п. 8.7)", "1 шт.", "п. 8.7"),
            ],
        ),
        (
            "Отделка",
            [
                ("О-01", "Грунт, бетоноконтакт, штукатурка / ровнитель стен", "по объёму стен 24 м²", "п. 5.4–5.5"),
                ("О-02", "Смесь для стяжки пола", "≈ 3,5 м²", "п. 5.2"),
                ("О-03", "Обмазочная гидроизоляция + угловая лента + грунт к ней", "пол 3,5 м² + стены ~10 м²", "п. 5.3, 5.6–5.7"),
                ("О-04", "Керамогранит на пол + запас 10–15%", "≈ 4,1 м² с запасом", "п. 6.4–6.5, 6.8"),
                ("О-05", "Плитка на стены + запас 10–15%", "≈ 30 м² с запасом", "п. 6.6–6.7"),
                ("О-06", "Клей для плитки, влагостойкая затирка, уголки / раскладки", "комплект", "п. 6.9–6.10"),
                ("О-07", "Санитарный силикон в цвет затирки", "2–3 тубы", "п. 6.12"),
                ("О-08", "Влагостойкий ГВЛ, брусок / профиль на короб инсталляции", "≈ 2,5 м²", "п. 4.6"),
                ("О-09", "Блоки / клей на подиум и экран ванны", "по месту", "п. 6.11, 7.1"),
            ],
        ),
        (
            "Прочее",
            [
                ("П-01", "Временный унитаз (если нет запасного) + гофра", "1 шт. на 18 дней", "п. 1.3"),
                ("П-02", "Вентилятор в существующий канал (если берёте п. 8.2)", "1 шт.", "п. 8.2"),
                ("П-03", "Кабель, розетки IP44, УЗО / дифавтомат (если берёте п. 8.3–8.4)", "комплект", "п. 8.3–8.4"),
                ("П-04", "Зеркало и аксессуары", "по выбору", "п. 8.5–8.6"),
                ("П-05", "Мешки для мусора 120 л, плёнка, картон, скотч", "на период ремонта", "п. 1.1, 2.10"),
            ],
        ),
    ]

    r = 6
    dv = DataValidation(type="list", formula1='"☐,☑"', allow_blank=True)
    dv.error = "Выберите ☐ или ☑"
    dv.errorTitle = "Отметка"
    dv.prompt = "Отметьте, когда куплено"
    dv.promptTitle = "Закуплено"
    ws.add_data_validation(dv)

    first_check = None
    last_check = None
    for gtitle, items in groups:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        paint_range(ws, r, 2, 6, fill(NAVY), height=20)
        ws.cell(r, 2).value = f"  {gtitle.upper()}"
        ws.cell(r, 2).font = font(size=10, bold=True, color=WHITE)
        for c in range(2, 7):
            ws.cell(r, c).fill = fill(NAVY)
        r += 1
        for i, (code, name, qty, where) in enumerate(items):
            bg = WHITE if i % 2 == 0 else ROW_ALT
            vals = (code, name, qty, where, "☐")
            for col, val in enumerate(vals, 2):
                cell = ws.cell(r, col, val)
                cell.fill = fill(bg)
                cell.border = thin_border(LINE_SOFT)
                cell.font = font(size=9, bold=(col == 2), color=NAVY if col == 2 else INK)
                cell.alignment = align("center" if col in (2, 4, 6) else "left", "center")
            if first_check is None:
                first_check = r
            last_check = r
            ws.row_dimensions[r].height = 22
            r += 1
        r += 1

    if first_check and last_check:
        dv.add(f"F{first_check}:F{last_check}")

    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(r, 2).value = (
        "Перед закупкой плитки и ванны — замер. Для ванны из наливного камня проверьте, "
        "проходит ли упаковка в проём 57 см или нужен занос через окно."
    )
    ws.cell(r, 2).font = font(size=9, italic=True, color=STEEL)
    ws.cell(r, 2).alignment = align("left", "center", wrap=True)
    paint_range(ws, r, 2, 6, fill(SAND), height=40)

    ws.print_area = f"A1:G{r}"
    ws.page_setup.fitToHeight = 1
    return ws


def add_workbook_props(wb: Workbook, totals: dict):
    wb.properties.title = "Смета на работы — ванная и туалет ЛГ-137"
    wb.properties.creator = "Локальная смета"
    wb.properties.subject = "Только работы, без материалов"
    wb.properties.description = (
        f"Основной объём {money(totals['main'])} ₽; "
        f"с резервом и рекомендациями {money(totals['grand'])} ₽"
    )
    wb.properties.keywords = "смета, ЛГ-137, ванная, туалет, работы"
    wb.properties.category = "Смета"


def main():
    wb = Workbook()
    totals = build_estimate(wb)
    build_cover(wb, totals)
    # cover is created after estimate; move to front
    # build_cover used active which is already not first if estimate created first.
    # We created estimate on new sheet, cover on active (first). Order: Титул, then we insert Сводка.
    build_summary(wb, totals)
    build_volumes(wb)
    build_conditions(wb)
    build_materials(wb)

    # sheet order
    desired = ["Титул", "Сводка", "Смета", "Объёмы", "Условия", "Закупка заказчика"]
    for i, name in enumerate(desired):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb["Титул"].sheet_properties.tabColor = GOLD
    wb["Сводка"].sheet_properties.tabColor = STEEL
    wb["Смета"].sheet_properties.tabColor = NAVY
    wb["Объёмы"].sheet_properties.tabColor = GREEN
    wb["Условия"].sheet_properties.tabColor = TERRACOTTA
    wb["Закупка заказчика"].sheet_properties.tabColor = GOLD_SOFT

    add_workbook_props(wb, totals)
    wb.save(OUT)

    # verification dump
    print("FILE", OUT)
    print("MAIN", int(totals["main"]))
    print("RESERVE", int(totals["reserve"]))
    print("REC", int(totals["rec"]))
    print("GRAND", int(totals["grand"]))
    print("WORDS", rub_words(int(totals["grand"])))
    for s in totals["sections"]:
        tag = "REC" if s["recommended"] else "IN"
        print(f"SEC {s['num']} [{tag}] {int(s['sum'])}  {s['title']}")
    print("BATH_FLOOR", BATH_FLOOR)
    print("WC_FLOOR", WC_FLOOR)
    print("BATH_WALLS", BATH_WALLS)
    print("WC_WALLS", WC_WALLS)
    print("WC_TILE_WALLS", WC_TILE_WALLS)
    print("HYDRO_WALLS", HYDRO_WALLS)
    print("SHOWER_ZONE", SHOWER_ZONE)
    print("GROUT", GROUT)


if __name__ == "__main__":
    main()
