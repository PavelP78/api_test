"""Проверки геометрии участка по координатам выписки ЕГРН."""

import json
import sys
from pathlib import Path

from shapely.geometry import Polygon, box

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

EXTRACT = json.loads((ROOT / "extract.json").read_text(encoding="utf-8"))


def _en(pts, origin):
    x0, y0 = origin
    return [(y - y0, x - x0) for x, y in pts]


def shoelace(pts):
    s = 0.0
    for i, (x1, y1) in enumerate(pts):
        x2, y2 = pts[(i + 1) % len(pts)]
        s += x1 * y2 - x2 * y1
    return abs(s) / 2.0


def test_plot_area_matches_egrn():
    pts = [(p["x"], p["y"]) for p in EXTRACT["geometry"]["points"]]
    area = shoelace(pts)
    declared = EXTRACT["object"]["area_m2"]
    err = EXTRACT["object"]["area_error_m2"]
    assert abs(area - declared) <= err
    assert abs(area - EXTRACT["object"]["area_computed_m2"]) < 0.01


def test_gas_zone_area_matches_egrn():
    pts = [(p["x"], p["y"]) for p in EXTRACT["geometry"]["zone_part_points"]]
    area = shoelace(pts)
    declared = EXTRACT["restrictions"]["parts"][0]["area_m2"]
    assert abs(area - declared) < 1.0
    assert area / EXTRACT["object"]["area_m2"] < 0.08


def test_sides_lengths():
    pts = [(p["x"], p["y"]) for p in EXTRACT["geometry"]["points"]]
    for i, side in enumerate(EXTRACT["geometry"]["sides"]):
        a, b = pts[i], pts[(i + 1) % 4]
        length = ((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2) ** 0.5
        assert abs(length - side["length_m"]) < 0.02


def test_house_fits_outside_gas_and_setback():
    origin = (EXTRACT["geometry"]["points"][0]["x"], EXTRACT["geometry"]["points"][0]["y"])
    plot = Polygon(_en([(p["x"], p["y"]) for p in EXTRACT["geometry"]["points"]], origin))
    zone = Polygon(_en([(p["x"], p["y"]) for p in EXTRACT["geometry"]["zone_part_points"]], origin))
    usable = plot.buffer(-3.0).difference(zone.buffer(0.5))
    house = box(-9.0, -20.0, -9.0 + 6.12, -20.0 + 10.0)
    # Ищем любую посадку 6.12×10 в usable.
    minx, miny, maxx, maxy = usable.bounds
    fits = False
    y = miny + 0.3
    while y <= maxy - 10.0:
        x = minx + 0.3
        while x <= maxx - 6.12:
            cand = box(x, y, x + 6.12, y + 10.0)
            if usable.contains(cand):
                fits = True
                house = cand
                break
            x += 0.8
        if fits:
            break
        y += 0.8
    assert fits, "дом 6,12×10 м не помещается вне ЗОУИТ и 3 м отступа"
    assert not house.intersects(zone)
    assert plot.contains(house)


def test_schema_generator_writes_artifacts(tmp_path, monkeypatch):
    import generate_schema

    monkeypatch.setattr(generate_schema, "ROOT", ROOT)
    generate_schema.main()
    assert (ROOT / "schema_uchastka.png").is_file()
    metrics = json.loads((ROOT / "metrics.json").read_text(encoding="utf-8"))
    assert metrics["house_fits"] is True
    assert 600 <= metrics["plot_area_m2"] <= 606
    assert 36 <= metrics["zone_area_m2"] <= 39
    assert metrics["usable_after_gas_and_setback_m2"] > 250
