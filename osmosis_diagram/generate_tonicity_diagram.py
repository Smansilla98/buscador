#!/usr/bin/env python3
"""Generate a Spanish tonicity/osmosis diagram with correct arrow directions.

Convention:
  - Solid arrows  = water movement OUT of the cell
  - Dashed arrows = water movement INTO the cell
"""

from __future__ import annotations

from pathlib import Path

import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Ellipse, FancyArrowPatch, FancyBboxPatch, Polygon


OUT_DIR = Path(__file__).resolve().parent
ARTIFACT_DIR = Path("/opt/cursor/artifacts/assets")

BG = "#F7F1E8"
PANEL = "#FFFDF9"
INK = "#1F2A37"
ACCENT = "#0B3A5B"
CELL = "#F4C7C0"
CELL_EDGE = "#C97B72"
SOLUTE = "#3FA66A"
ARROW = "#1F2A37"
CAPTION_BG = "#EEF3F7"


def draw_arrow(ax, start, end, *, dashed: bool) -> None:
    """Draw a directional arrow. dashed=True means INTO the cell."""
    style = dict(
        arrowstyle="-|>",
        mutation_scale=16,
        linewidth=2.2,
        color=ARROW,
        shrinkA=0,
        shrinkB=0,
    )
    if dashed:
        style["linestyle"] = (0, (5, 3.5))
    ax.add_patch(FancyArrowPatch(start, end, **style))


def draw_solutes(ax, points, radius=0.085) -> None:
    for x, y in points:
        ax.add_patch(
            Circle(
                (x, y),
                radius,
                facecolor=SOLUTE,
                edgecolor="#2E7A4D",
                linewidth=1.2,
                zorder=3,
            )
        )


def draw_caption(ax, text: str) -> None:
    ax.add_patch(
        FancyBboxPatch(
            (0.08, 0.05),
            0.84,
            0.14,
            boxstyle="round,pad=0.02,rounding_size=0.04",
            facecolor=CAPTION_BG,
            edgecolor="#C5D2DE",
            linewidth=1.2,
            transform=ax.transAxes,
            zorder=4,
        )
    )
    ax.text(
        0.5,
        0.12,
        text,
        ha="center",
        va="center",
        fontsize=11,
        color=INK,
        transform=ax.transAxes,
        zorder=5,
        wrap=True,
    )


def draw_title(ax, title: str) -> None:
    ax.text(
        0.5,
        0.93,
        title,
        ha="center",
        va="center",
        fontsize=16,
        fontweight="bold",
        color=ACCENT,
        transform=ax.transAxes,
        zorder=5,
    )


def panel_frame(ax) -> None:
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.add_patch(
        FancyBboxPatch(
            (0.02, 0.02),
            0.96,
            0.96,
            boxstyle="round,pad=0.01,rounding_size=0.03",
            facecolor=PANEL,
            edgecolor="#D7DEE6",
            linewidth=1.4,
            zorder=0,
        )
    )


def draw_isotonic(ax) -> None:
    panel_frame(ax)
    draw_title(ax, "Isotónica")

    cx, cy, rx, ry = 0.50, 0.52, 0.18, 0.16
    ax.add_patch(
        Ellipse(
            (cx, cy),
            2 * rx,
            2 * ry,
            facecolor=CELL,
            edgecolor=CELL_EDGE,
            linewidth=2.4,
            zorder=2,
        )
    )

    # Equal solute concentration inside/outside (3 + 3)
    draw_solutes(
        ax,
        [
            (0.44, 0.55),
            (0.52, 0.48),
            (0.56, 0.56),
            (0.20, 0.64),
            (0.80, 0.60),
            (0.28, 0.36),
        ],
    )

    # Dashed INTO cell (from top-left); solid OUT of cell (to top-right)
    draw_arrow(ax, (0.28, 0.78), (0.40, 0.66), dashed=True)
    draw_arrow(ax, (0.60, 0.66), (0.72, 0.78), dashed=False)

    draw_caption(ax, "No hay flujo neto de agua")


def draw_hypotonic(ax) -> None:
    panel_frame(ax)
    draw_title(ax, "Hipotónica")

    cx, cy, r = 0.50, 0.50, 0.22
    ax.add_patch(
        Circle(
            (cx, cy),
            r,
            facecolor=CELL,
            edgecolor=CELL_EDGE,
            linewidth=2.4,
            zorder=2,
        )
    )

    # More solutes inside than outside
    draw_solutes(
        ax,
        [
            (0.43, 0.54),
            (0.52, 0.46),
            (0.57, 0.55),
            (0.18, 0.58),
        ],
    )

    # Only dashed arrows INTO the swollen cell
    draw_arrow(ax, (0.28, 0.80), (0.40, 0.68), dashed=True)
    draw_arrow(ax, (0.50, 0.84), (0.50, 0.72), dashed=True)
    draw_arrow(ax, (0.72, 0.80), (0.60, 0.68), dashed=True)

    draw_caption(ax, "Entrada neta de agua: la célula se hincha")


def shriveled_cell_polygon(cx: float, cy: float, scale: float = 0.18):
    """Irregular jagged outline for a crenated cell."""
    import math

    pts = []
    radii = [1.00, 0.72, 0.96, 0.70, 1.02, 0.68, 0.98, 0.74, 1.00, 0.70, 0.97, 0.73]
    for i, rad in enumerate(radii):
        ang = -math.pi / 2 + i * (2 * math.pi / len(radii))
        pts.append((cx + scale * rad * math.cos(ang), cy + scale * rad * math.sin(ang) * 0.92))
    return pts


def draw_hypertonic(ax) -> None:
    panel_frame(ax)
    draw_title(ax, "Hipertónica")

    pts = shriveled_cell_polygon(0.50, 0.52, 0.17)
    ax.add_patch(
        Polygon(
            pts,
            closed=True,
            facecolor=CELL,
            edgecolor=CELL_EDGE,
            linewidth=2.4,
            zorder=2,
        )
    )

    # More solutes outside than inside
    draw_solutes(
        ax,
        [
            (0.45, 0.54),
            (0.52, 0.48),
            (0.55, 0.56),
            (0.18, 0.62),
            (0.82, 0.60),
            (0.22, 0.38),
            (0.78, 0.36),
        ],
    )

    # Solid OUT; dashed IN (net exit because more solid arrows)
    draw_arrow(ax, (0.38, 0.66), (0.26, 0.80), dashed=False)
    draw_arrow(ax, (0.62, 0.66), (0.74, 0.80), dashed=False)
    draw_arrow(ax, (0.50, 0.36), (0.50, 0.22), dashed=False)
    draw_arrow(ax, (0.24, 0.34), (0.38, 0.44), dashed=True)
    draw_arrow(ax, (0.76, 0.34), (0.62, 0.44), dashed=True)

    draw_caption(ax, "Salida neta de agua: la célula se encoge")


def draw_legend_final(fig) -> None:
    legend_ax = fig.add_axes([0.06, 0.03, 0.88, 0.11])
    legend_ax.set_xlim(0, 1)
    legend_ax.set_ylim(0, 1)
    legend_ax.axis("off")
    legend_ax.add_patch(
        FancyBboxPatch(
            (0.0, 0.05),
            1.0,
            0.90,
            boxstyle="round,pad=0.015,rounding_size=0.04",
            facecolor="#FFFFFF",
            edgecolor="#D7DEE6",
            linewidth=1.2,
        )
    )

    items = [
        ("solute", 0.03, "Solutos (sustancias disueltas)"),
        ("dashed", 0.36, "Movimiento de agua hacia el interior"),
        ("solid", 0.70, "Movimiento de agua hacia el exterior"),
    ]

    for kind, x, label in items:
        if kind == "solute":
            legend_ax.add_patch(
                Circle((x + 0.025, 0.50), 0.07, facecolor=SOLUTE, edgecolor="#2E7A4D", linewidth=1.2, transform=legend_ax.transData)
            )
            # Circle uses data coords; with xlim 0-1 radius 0.07 is fine.
            legend_ax.text(x + 0.06, 0.50, label, va="center", ha="left", fontsize=10.5, color=INK)
        elif kind == "dashed":
            legend_ax.annotate(
                "",
                xy=(x + 0.055, 0.50),
                xytext=(x + 0.005, 0.50),
                arrowprops=dict(
                    arrowstyle="-|>",
                    color=ARROW,
                    lw=2.0,
                    linestyle=(0, (4, 3)),
                    mutation_scale=13,
                ),
            )
            legend_ax.text(x + 0.07, 0.50, label, va="center", ha="left", fontsize=10.5, color=INK)
        else:
            legend_ax.annotate(
                "",
                xy=(x + 0.055, 0.50),
                xytext=(x + 0.005, 0.50),
                arrowprops=dict(arrowstyle="-|>", color=ARROW, lw=2.0, mutation_scale=13),
            )
            legend_ax.text(x + 0.07, 0.50, label, va="center", ha="left", fontsize=10.5, color=INK)


def main() -> None:
    fig = plt.figure(figsize=(14.5, 6.2), facecolor=BG)
    fig.suptitle(
        "Efecto de la concentración de solutos sobre la célula",
        fontsize=18,
        fontweight="bold",
        color=ACCENT,
        y=0.97,
    )

    axes = [
        fig.add_axes([0.035, 0.18, 0.30, 0.72]),
        fig.add_axes([0.35, 0.18, 0.30, 0.72]),
        fig.add_axes([0.665, 0.18, 0.30, 0.72]),
    ]

    draw_isotonic(axes[0])
    draw_hypotonic(axes[1])
    draw_hypertonic(axes[2])
    draw_legend_final(fig)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    ARTIFACT_DIR.mkdir(parents=True, exist_ok=True)

    outputs = [
        OUT_DIR / "tonicity_osmosis_diagram.png",
        ARTIFACT_DIR / "tonicity_osmosis_diagram.png",
    ]
    for path in outputs:
        fig.savefig(path, dpi=200, facecolor=fig.get_facecolor(), bbox_inches="tight")
        print(f"Wrote {path}")

    plt.close(fig)


if __name__ == "__main__":
    main()
