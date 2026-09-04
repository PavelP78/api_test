#!/usr/bin/env python3
"""Схема участка 47:01:1516001:2322 по координатам выписки ЕГРН.

Строит план в МСК-47 (X — север, Y — восток), накладывает охранную зону
газопровода и ориентир дома 6,12×10 м из сметы отделки.
"""

from __future__ import annotations

import json
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyBboxPatch, Polygon as MplPolygon
from shapely.geometry import Polygon, box
from shapely.affinity import translate

ROOT = Path(__file__).resolve().parent
EXTRACT = json.loads((ROOT / "extract.json").read_text(encoding="utf-8"))

# МСК-47: X = север, Y = восток. Для чертежа: east, north.
PLOT_MSK = [(p["x"], p["y"]) for p in EXTRACT["geometry"]["points"]]
ZONE_MSK = [(p["x"], p["y"]) for p in EXTRACT["geometry"]["zone_part_points"]]

HOUSE_W = 6.12  # восток–запад, м
HOUSE_L = 10.00  # север–юг, м
SETBACK = 3.0
GAS_BUFFER = 0.5


def to_en(pts):
    """(X north, Y east) → (east, north) относительно точки 1."""
    x0, y0 = PLOT_MSK[0]
    return [(y - y0, x - x0) for x, y in pts]


def shoelace(pts):
    s = 0.0
    n = len(pts)
    for i in range(n):
        x1, y1 = pts[i]
        x2, y2 = pts[(i + 1) % n]
        s += x1 * y2 - x2 * y1
    return abs(s) / 2.0


def place_house(usable: Polygon):
    """Ставит дом 6,12×10 ближе к центру свободного контура, ось С–Ю."""
    minx, miny, maxx, maxy = usable.bounds
    cx, cy = usable.centroid.x, usable.centroid.y
    best = None
    ys = np.arange(miny + 0.2, maxy - HOUSE_L - 0.2, 0.4)
    xs = np.arange(minx + 0.2, maxx - HOUSE_W - 0.2, 0.4)
    for y in ys:
        for x in xs:
            cand = box(x, y, x + HOUSE_W, y + HOUSE_L)
            if usable.contains(cand):
                dx = x + HOUSE_W / 2 - cx
                dy = y + HOUSE_L / 2 - cy
                score = dx * dx + dy * dy
                if best is None or score < best[0]:
                    best = (score, cand)
    if best:
        return best[1]
    return translate(box(0, 0, HOUSE_W, HOUSE_L), cx - HOUSE_W / 2, cy - HOUSE_L / 2)


def main():
    plot_en = to_en(PLOT_MSK)
    zone_en = to_en(ZONE_MSK)
    plot = Polygon(plot_en)
    zone = Polygon(zone_en)
    inset = plot.buffer(-SETBACK)
    if inset.is_empty:
        raise SystemExit("3 м отступ съедает весь участок — проверьте координаты")
    usable = inset.difference(zone.buffer(GAS_BUFFER))
    house = place_house(usable if not usable.is_empty else inset)

    plot_area = shoelace(PLOT_MSK)
    zone_area = shoelace(ZONE_MSK)
    buildable = max(plot_area - zone_area, 0.0)
    zone_width = zone_area / EXTRACT["geometry"]["sides"][3]["length_m"]

    metrics = {
        "crs": "МСК-47, зона 1; чертёж: восток–север, м, точка 1 = (0, 0)",
        "plot_area_m2": round(plot_area, 2),
        "zone_area_m2": round(zone_area, 2),
        "buildable_without_setbacks_m2": round(buildable, 2),
        "zone_share_pct": round(100.0 * zone_area / plot_area, 2),
        "zone_typical_width_m": round(zone_width, 2),
        "setback_m": SETBACK,
        "inset_area_m2": round(inset.area, 2),
        "usable_after_gas_and_setback_m2": round(usable.area, 2),
        "house_m": [HOUSE_W, HOUSE_L],
        "house_fits": bool(usable.contains(house)),
        "house_sw_east_north_m": [round(house.bounds[0], 2), round(house.bounds[1], 2)],
        "perimeter_m": round(plot.length, 2),
    }
    (ROOT / "metrics.json").write_text(
        json.dumps(metrics, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )

    plt.rcParams["font.family"] = "DejaVu Sans"
    fig, ax = plt.subplots(figsize=(11.5, 13.2), dpi=160)
    fig.patch.set_facecolor("#f7f4ee")
    ax.set_facecolor("#f7f4ee")

    def add_poly(coords, **kw):
        patch = MplPolygon(coords, closed=True, **kw)
        ax.add_patch(patch)
        return patch

    add_poly(
        plot_en,
        facecolor="#c5e0b4",
        edgecolor="#1f4e79",
        linewidth=2.2,
        zorder=2,
        label="Участок :2322 (603 м²)",
    )
    add_poly(
        zone_en,
        facecolor="#f4b183",
        edgecolor="#c45911",
        linewidth=1.4,
        hatch="///",
        alpha=0.92,
        zorder=3,
        label=f"ЗОУИТ газа :2322/1 ({zone_area:.0f} м²)",
    )

    if not inset.is_empty:
        ix, iy = inset.exterior.xy
        ax.plot(
            ix,
            iy,
            color="#7f6000",
            linestyle="--",
            linewidth=1.2,
            zorder=4,
            label=f"Отступ {SETBACK:.0f} м от границ",
        )

    hx, hy = house.exterior.xy
    ax.fill(
        hx,
        hy,
        facecolor="#5b9bd5",
        edgecolor="#1f4e79",
        linewidth=1.6,
        alpha=0.85,
        zorder=5,
        label=f"Ориентир дома {HOUSE_W:g}×{HOUSE_L:g} м",
    )
    ax.text(
        house.centroid.x,
        house.centroid.y,
        f"дом\n{HOUSE_W:g}×{HOUSE_L:g} м\n≈{HOUSE_W * HOUSE_L:.0f} м²",
        ha="center",
        va="center",
        fontsize=8.5,
        color="#08306b",
        zorder=6,
        fontweight="bold",
    )

    # Подписи точек и длин сторон.
    for i, (e, n) in enumerate(plot_en, 1):
        ax.plot(e, n, "o", color="#1f4e79", markersize=6, zorder=7)
        ax.annotate(
            str(i),
            (e, n),
            textcoords="offset points",
            xytext=(6, 6),
            fontsize=9,
            fontweight="bold",
            color="#1f4e79",
            zorder=8,
        )

    sides = EXTRACT["geometry"]["sides"]
    for i, side in enumerate(sides):
        a = plot_en[i]
        b = plot_en[(i + 1) % 4]
        mx, my = (a[0] + b[0]) / 2, (a[1] + b[1]) / 2
        dx, dy = b[0] - a[0], b[1] - a[1]
        L = (dx**2 + dy**2) ** 0.5
        nx, ny = -dy / L, dx / L
        # Вынести подпись наружу.
        ax.annotate(
            f"{side['length_m']} м",
            (mx + nx * 1.6, my + ny * 1.6),
            ha="center",
            va="center",
            fontsize=8,
            color="#1f4e79",
            bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="#1f4e79", alpha=0.9),
            zorder=8,
        )

    # Соседи — снаружи соответствующих сторон.
    neighbor_labels = {
        0: "смежный участок\nне указан",
        1: "смежный\n:2323",
        2: "смежный\n:2270",
        3: "смежный :201\n+ ЗОУИТ газа",
    }
    for i, text in neighbor_labels.items():
        a = np.array(plot_en[i])
        b = np.array(plot_en[(i + 1) % 4])
        mid = (a + b) / 2
        dx, dy = b - a
        L = np.hypot(dx, dy)
        nrm = np.array([-dy, dx]) / L
        pos = mid + nrm * 5.4
        ax.text(
            pos[0],
            pos[1],
            text,
            ha="center",
            va="center",
            fontsize=7.5,
            color="#595959",
            zorder=8,
        )

    xs_p, ys_p = zip(*plot_en)
    north_x = max(xs_p) + 8.5
    ax.annotate(
        "",
        xy=(north_x, 5.5),
        xytext=(north_x, 0.5),
        arrowprops=dict(arrowstyle="-|>", color="#333", lw=1.6),
        zorder=9,
    )
    ax.text(north_x, 6.1, "С", ha="center", fontsize=11, fontweight="bold")

    scale_y = min(ys_p) - 4.6
    scale_x0 = max(xs_p) - 2.0
    ax.plot([scale_x0, scale_x0 + 10], [scale_y, scale_y], color="#333", lw=2.4, solid_capstyle="butt")
    ax.plot([scale_x0, scale_x0], [scale_y - 0.3, scale_y + 0.3], color="#333", lw=2.4)
    ax.plot([scale_x0 + 10, scale_x0 + 10], [scale_y - 0.3, scale_y + 0.3], color="#333", lw=2.4)
    ax.text(scale_x0 + 5, scale_y - 1.3, "10 м", ha="center", fontsize=8)

    ax.set_aspect("equal")
    ax.autoscale()
    x0, x1 = ax.get_xlim()
    y0, y1 = ax.get_ylim()
    ax.set_xlim(x0 - 3, x1 + 5)
    ax.set_ylim(y0 - 3, y1 + 3)
    ax.set_xlabel("Восток, м (от точки 1)")
    ax.set_ylabel("Север, м (от точки 1)")
    ax.grid(True, linestyle=":", color="#d0d0d0", zorder=1)
    ax.legend(loc="lower left", framealpha=0.95, fontsize=8)

    title = (
        "Земельный участок 47:01:1516001:2322\n"
        "Красносельское СП, Выборгский район, ЛО · пос. Коробицыно"
    )
    ax.set_title(title, fontsize=13, fontweight="bold", color="#1f4e79", pad=12)

    info = (
        f"Площадь по координатам {plot_area:.0f} м² (в ЕГРН 603±9). "
        f"Охранная зона газа {zone_area:.0f} м² ≈ {metrics['zone_typical_width_m']} м "
        f"вдоль северной границы (15,49 м). "
        f"После отступа 3 м и ЗОУИТ свободно ≈ {metrics['usable_after_gas_and_setback_m2']:.0f} м². "
        f"Дом {HOUSE_W:g}×{HOUSE_L:g} м "
        + ("размещается вне зоны газа." if metrics["house_fits"] else "требует уточнения посадки.")
    )
    fig.text(0.5, 0.018, info, ha="center", va="bottom", fontsize=8, color="#333", wrap=True)

    header = FancyBboxPatch(
        (0.015, 0.955),
        0.97,
        0.038,
        transform=fig.transFigure,
        boxstyle="round,pad=0.004",
        facecolor="#1f4e79",
        edgecolor="none",
    )
    fig.patches.append(header)
    fig.text(
        0.5,
        0.974,
        "Выписка ЕГРН № КУВИ-001/2026-112602140 от 25.08.2026  ·  схема по координатам МСК-47",
        ha="center",
        va="center",
        color="white",
        fontsize=8.5,
        fontweight="bold",
    )

    fig.tight_layout(rect=(0.03, 0.045, 0.98, 0.95))
    out = ROOT / "schema_uchastka.png"
    fig.savefig(out, dpi=160)
    plt.close(fig)
    print(json.dumps(metrics, ensure_ascii=False, indent=2))
    print(f"wrote {out}")


if __name__ == "__main__":
    main()
