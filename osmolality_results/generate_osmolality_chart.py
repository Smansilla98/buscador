#!/usr/bin/env python3
"""Generate an Excel-style osmolality results chart with specification ranges."""

from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
import numpy as np
from openpyxl import Workbook
from openpyxl.chart import BarChart, ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import SeriesLabel
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT_DIR = Path(__file__).resolve().parent

PRODUCTS = [
    {"product": "DO1", "spec_min": 280, "spec_max": 320, "sample_1": 282, "sample_2": 280},
    {"product": "DO2", "spec_min": 240, "spec_max": 340, "sample_1": 286, "sample_2": 283},
    {"product": "DO3", "spec_min": 250, "spec_max": 350, "sample_1": 280, "sample_2": 281},
    {"product": "DO4", "spec_min": 260, "spec_max": 320, "sample_1": 281, "sample_2": 276},
    {"product": "DO5", "spec_min": 260, "spec_max": 320, "sample_1": 276, "sample_2": 275},
    {"product": "DO6", "spec_min": 240, "spec_max": 370, "sample_1": 287, "sample_2": 280},
    {"product": "DO7", "spec_min": 270, "spec_max": 330, "sample_1": 310, "sample_2": 312},
    {"product": "DO8", "spec_min": None, "spec_max": None, "sample_1": 573, "sample_2": 579},
]


def build_chart_png(path: Path) -> None:
    """Excel-like chart: shaded spec ranges + circular sample markers."""
    products = [p["product"] for p in PRODUCTS]
    x = np.arange(len(products))

    fig, ax = plt.subplots(figsize=(12, 7), dpi=160)
    fig.patch.set_facecolor("#FFFFFF")
    ax.set_facecolor("#FFFFFF")

    bar_width = 0.55
    for i, p in enumerate(PRODUCTS):
        if p["spec_min"] is None or p["spec_max"] is None:
            continue
        height = p["spec_max"] - p["spec_min"]
        ax.bar(
            i,
            height,
            bottom=p["spec_min"],
            width=bar_width,
            color="#9DC3E6",
            alpha=0.45,
            edgecolor="#5B9BD5",
            linewidth=0.8,
            zorder=1,
        )

    s1 = [p["sample_1"] for p in PRODUCTS]
    s2 = [p["sample_2"] for p in PRODUCTS]
    offset = 0.08
    ax.scatter(
        x - offset,
        s1,
        s=55,
        facecolors="#ED7D31",
        edgecolors="#C55A11",
        linewidths=1.0,
        zorder=3,
        label="Sample 1",
    )
    ax.scatter(
        x + offset,
        s2,
        s=55,
        facecolors="#70AD47",
        edgecolors="#548235",
        linewidths=1.0,
        zorder=3,
        label="Sample 2",
    )

    ax.set_axisbelow(True)
    ax.yaxis.grid(True, which="major", color="#D9D9D9", linestyle="-", linewidth=0.8)
    ax.xaxis.grid(True, which="major", color="#F2F2F2", linestyle="-", linewidth=0.6)
    ax.set_xlim(-0.6, len(products) - 0.4)
    ax.set_ylim(200, 620)
    ax.set_xticks(x)
    ax.set_xticklabels(products, fontsize=11)
    ax.set_ylabel("Osmolality (mOsm/kg)", fontsize=11)
    ax.set_xlabel("Product", fontsize=11)
    ax.set_title(
        "Osmolality Test Results by Product",
        fontsize=14,
        fontweight="bold",
        pad=14,
        color="#1F4E79",
    )

    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.spines["left"].set_color("#BFBFBF")
    ax.spines["bottom"].set_color("#BFBFBF")
    ax.tick_params(colors="#595959")

    legend_handles = [
        mpatches.Patch(
            facecolor="#9DC3E6",
            edgecolor="#5B9BD5",
            alpha=0.45,
            label="Specification range",
        ),
        Line2D(
            [0],
            [0],
            marker="o",
            color="w",
            markerfacecolor="#ED7D31",
            markeredgecolor="#C55A11",
            markersize=8,
            label="Sample 1",
        ),
        Line2D(
            [0],
            [0],
            marker="o",
            color="w",
            markerfacecolor="#70AD47",
            markeredgecolor="#548235",
            markersize=8,
            label="Sample 2",
        ),
    ]
    ax.legend(handles=legend_handles, loc="upper left", frameon=True, fancybox=False)

    ax.annotate(
        "DO8: no specification range",
        xy=(7, 576),
        xytext=(5.2, 540),
        fontsize=8,
        color="#595959",
        arrowprops=dict(arrowstyle="->", color="#595959", lw=0.8),
    )

    fig.tight_layout()
    fig.savefig(path, bbox_inches="tight")
    plt.close(fig)


def within_spec(value: float, lo, hi) -> str:
    if lo is None or hi is None:
        return "N/A (no specification)"
    return "Pass" if lo <= value <= hi else "Out of range"


def build_workbook(chart_png: Path, path: Path) -> None:
    wb = Workbook()

    # ---- Summary ----
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum["A1"] = "Osmolality Test — Summary"
    ws_sum["A1"].font = Font(bold=True, size=16, color="1F4E79")
    ws_sum["A3"] = (
        "All products with a specification range are within limits for both samples."
    )
    ws_sum["A4"] = (
        "DO8 has no specification range; measured values are ~573–579 mOsm/kg."
    )
    ws_sum["A6"] = "Contents"
    ws_sum["A6"].font = Font(bold=True)
    ws_sum["A7"] = "• Sheet 'Data': raw results + native Excel combo chart"
    ws_sum["A8"] = "• Sheet 'Chart': high-resolution Excel-style chart image"
    ws_sum["A9"] = "• Companion PNG file for presentations / reports"
    ws_sum.column_dimensions["A"].width = 90

    # ---- Data ----
    ws = wb.create_sheet("Data")
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Border(
        left=Side(style="thin", color="B4B4B4"),
        right=Side(style="thin", color="B4B4B4"),
        top=Side(style="thin", color="B4B4B4"),
        bottom=Side(style="thin", color="B4B4B4"),
    )
    center = Alignment(horizontal="center", vertical="center")

    headers = [
        "Product",
        "Spec Min (mOsm/kg)",
        "Spec Max (mOsm/kg)",
        "Sample 1 (mOsm/kg)",
        "Sample 2 (mOsm/kg)",
        "Sample 1 Status",
        "Sample 2 Status",
        "Base (hidden)",
        "Range Height",
        "X Index",
        "Sample 1 Y",
        "Sample 2 Y",
    ]

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(1, col, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin

    for row_idx, p in enumerate(PRODUCTS, start=2):
        values = [
            p["product"],
            p["spec_min"] if p["spec_min"] is not None else "—",
            p["spec_max"] if p["spec_max"] is not None else "—",
            p["sample_1"],
            p["sample_2"],
            within_spec(p["sample_1"], p["spec_min"], p["spec_max"]),
            within_spec(p["sample_2"], p["spec_min"], p["spec_max"]),
            p["spec_min"] if p["spec_min"] is not None else 0,
            (p["spec_max"] - p["spec_min"]) if p["spec_min"] is not None else 0,
            row_idx - 1,
            p["sample_1"],
            p["sample_2"],
        ]
        for col, value in enumerate(values, start=1):
            cell = ws.cell(row_idx, col, value)
            cell.alignment = center
            cell.border = thin
            if col in (6, 7):
                if value == "Pass":
                    cell.fill = PatternFill("solid", fgColor="C6EFCE")
                    cell.font = Font(color="006100")
                elif str(value).startswith("Out"):
                    cell.fill = PatternFill("solid", fgColor="FFC7CE")
                    cell.font = Font(color="9C0006")
                else:
                    cell.fill = PatternFill("solid", fgColor="FFF2CC")
                    cell.font = Font(color="9C5700")

    widths = [12, 18, 18, 18, 18, 22, 22, 14, 14, 10, 12, 12]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width

    for col in range(8, 13):
        ws.column_dimensions[get_column_letter(col)].hidden = True

    ws["A11"] = "Notes"
    ws["A11"].font = Font(bold=True, color="1F4E79")
    ws["A12"] = (
        "Specification range is shown as a shaded band for each product. "
        "DO8 has no specification; only sample markers are displayed."
    )
    ws["A13"] = "Units: mOsm/kg. Markers: Sample 1 (orange), Sample 2 (green)."
    ws.merge_cells("A12:G12")
    ws.merge_cells("A13:G13")

    # Floating stacked columns for specification shade
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "stacked"
    bar.overlap = 100
    bar.title = "Osmolality Test Results by Product"
    bar.y_axis.title = "Osmolality (mOsm/kg)"
    bar.x_axis.title = "Product"
    bar.y_axis.scaling.min = 200
    bar.y_axis.scaling.max = 620
    bar.style = 10
    bar.width = 18
    bar.height = 12

    cats = Reference(ws, min_col=1, min_row=2, max_row=9)
    data_bar = Reference(ws, min_col=8, min_row=1, max_col=9, max_row=9)
    bar.add_data(data_bar, titles_from_data=True)
    bar.set_categories(cats)

    # Invisible base series
    bar.series[0].graphicalProperties.noFill = True
    bar.series[0].graphicalProperties.line.noFill = True
    bar.series[0].title = SeriesLabel(v=" ")

    # Visible range shade
    bar.series[1].graphicalProperties.solidFill = "9DC3E6"
    bar.series[1].graphicalProperties.line.solidFill = "5B9BD5"
    bar.series[1].title = SeriesLabel(v="Specification range")

    # Scatter markers for both samples
    scatter = ScatterChart()
    scatter.style = 10

    xvalues = Reference(ws, min_col=10, min_row=2, max_row=9)
    y1 = Reference(ws, min_col=11, min_row=1, max_row=9)
    y2 = Reference(ws, min_col=12, min_row=1, max_row=9)

    # Series constructor: Series(values, xvalues, title=...)
    # When title is a string and values include header row, use titles carefully.
    # Use values without header + explicit title.
    y1_vals = Reference(ws, min_col=11, min_row=2, max_row=9)
    y2_vals = Reference(ws, min_col=12, min_row=2, max_row=9)

    ser1 = Series(y1_vals, xvalues, title="Sample 1")
    ser1.marker = Marker(symbol="circle", size=7)
    ser1.marker.graphicalProperties.solidFill = "ED7D31"
    ser1.marker.graphicalProperties.line.solidFill = "C55A11"
    ser1.graphicalProperties.line.noFill = True
    scatter.series.append(ser1)

    ser2 = Series(y2_vals, xvalues, title="Sample 2")
    ser2.marker = Marker(symbol="circle", size=7)
    ser2.marker.graphicalProperties.solidFill = "70AD47"
    ser2.marker.graphicalProperties.line.solidFill = "548235"
    ser2.graphicalProperties.line.noFill = True
    scatter.series.append(ser2)

    bar.y_axis.majorGridlines.spPr = None  # keep default gridlines
    bar += scatter
    ws.add_chart(bar, "A15")

    # ---- Chart image sheet ----
    ws_chart = wb.create_sheet("Chart")
    ws_chart["A1"] = "Osmolality Test Results — Excel-style chart"
    ws_chart["A1"].font = Font(bold=True, size=14, color="1F4E79")
    ws_chart["A2"] = (
        "X axis: products (DO1–DO8) | Y axis: osmolality (mOsm/kg) | "
        "Shade: specification range | Circles: Sample 1 & Sample 2"
    )
    ws_chart.merge_cells("A2:H2")
    ws_chart.column_dimensions["A"].width = 20

    if chart_png.exists():
        img = XLImage(str(chart_png))
        img.width = 960
        img.height = 560
        ws_chart.add_image(img, "A4")

    wb.save(path)


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    png_path = OUTPUT_DIR / "osmolality_chart.png"
    xlsx_path = OUTPUT_DIR / "osmolality_results.xlsx"

    build_chart_png(png_path)
    build_workbook(png_path, xlsx_path)
    print(f"Wrote {png_path}")
    print(f"Wrote {xlsx_path}")


if __name__ == "__main__":
    main()
