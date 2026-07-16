"""
Generates the two output Excel tabs (Loads Summary, Orders Line Summary)
and one landscape schematic page per load (multi-page PDF), from the
load_builder prototype's results.

Schematic v3: ink-friendly (hatch patterns instead of solid fills, no big
color blocks), legend moved to the top so trailers get the full page width,
weight/floor/cube badges per trailer, and full order/SKU/customer text on
every bundle.
"""
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.backends.backend_pdf import PdfPages

from load_builder import (
    load_workbook_data, enrich_lines, assemble_loads, pack_load, TRUCK_TYPES, group_label
)

NAVY = "#1F3352"
STEEL = "#4A6B8A"
INK = "#1A1A1A"

# Distinct customer colors used ONLY as thin borders/text accents (cheap on
# ink); the bundle fill itself is white/very light with a black hatch pattern
# so the whole schematic prints fine on a low-ink black & white printer.
ACCENTS = ["#4C78A8", "#F58518", "#54A24B", "#B2323C", "#4FA9A5",
           "#B8860B", "#7B4F9D", "#C2568B", "#8C5A2B", "#5C5C5C"]
HATCHES = ["///", "\\\\\\", "|||", "---", "+++", "xxx", "...", "ooo", "***", "\\\\|"]


def _customer_style_map(load):
    codes = sorted({p.get("location_code") for t in load["packing"].values() for p in t["placements"]})
    return {code: (ACCENTS[i % len(ACCENTS)], HATCHES[i % len(HATCHES)]) for i, code in enumerate(codes)}

THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(bold=True)


def build_all():
    sites, customers, skus, orders, excluded = load_workbook_data()
    lines = enrich_lines(orders, customers, skus, sites)
    loads, unassigned, fleet_left = assemble_loads(lines)
    for load in loads:
        packing, leftover = pack_load(load)
        load["packing"] = packing
        load["pack_leftover"] = leftover
    return loads, unassigned, fleet_left, lines


def sheet_style_print(ws, n_cols):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.5, bottom=0.5)
    ws.print_options.horizontalCentered = True
    for i in range(1, n_cols + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16


def write_loads_summary(wb, loads, unassigned):
    ws = wb.active
    ws.title = "Loads Summary"
    headers = ["Load ID", "Site", "Truck Type", "Direction/Area Group",
               "Total m3", "Vol Util %", "Total KG", "Weight Util %",
               "# Orders", "Full/Partial", "# Unique SKUs", "# Customers/Drops",
               "Bundles Not Placed (packing)"]
    ws.append(headers)
    for c in ws[1]:
        c.font = BOLD
        c.border = BORDER
        c.alignment = Alignment(horizontal="center")

    order_loads = defaultdict(set)
    for load in loads:
        for ln in load["lines"]:
            order_loads[ln["sales_order"]].add(load["load_id"])
    order_unassigned = defaultdict(bool)
    for ln in unassigned:
        order_unassigned[ln["sales_order"]] = True

    for load in loads:
        spec = TRUCK_TYPES[load["truck_type"]]
        vol_util = load["total_m3"] / spec["cube_cap_m3"] * 100
        wt_util = load["total_kg"] / (spec["payload_cap_t"] * 1000) * 100
        so_set = {ln["sales_order"] for ln in load["lines"]}
        sku_set = {ln["sku"] for ln in load["lines"]}
        cust_set = {ln["location_code"] for ln in load["lines"]}
        full_flags = []
        for so in so_set:
            is_full = len(order_loads[so]) == 1 and not order_unassigned[so]
            full_flags.append("Full" if is_full else "Partial")
        full_partial = "Full" if all(f == "Full" for f in full_flags) else "Partial"
        leftover_n = len(load.get("pack_leftover", []))
        row = [load["load_id"], load["site"], load["truck_type"], group_label(load["group"]),
               round(load["total_m3"], 2), round(vol_util, 0), round(load["total_kg"], 0), round(wt_util, 0),
               len(so_set), full_partial, len(sku_set), len(cust_set), leftover_n]
        ws.append(row)
        for c in ws[ws.max_row]:
            c.border = BORDER
            c.alignment = Alignment(horizontal="center")

    sheet_style_print(ws, len(headers))
    ws.freeze_panes = "A2"
    return ws


def write_orders_summary(wb, loads, unassigned):
    ws = wb.create_sheet("Orders Line Summary")
    headers = ["Sales Order", "Due Date", "Customer", "Location Code", "SKU",
               "Bundles", "m3", "KG", "Assigned Load ID", "Site", "Status"]
    ws.append(headers)
    for c in ws[1]:
        c.font = BOLD
        c.border = BORDER
        c.alignment = Alignment(horizontal="center")

    def add_row(ln, load_id, status):
        kg = ln["bundle_kg"] * ln["bundles"]
        ws.append([ln["sales_order"], ln["due"].strftime("%Y-%m-%d") if ln["due"] else "",
                   ln["delivery_name"], ln["location_code"], ln["sku"], ln["bundles"],
                   round(ln["m3"], 2), round(kg, 1), load_id, ln["site"], status])
        for c in ws[ws.max_row]:
            c.border = BORDER

    for load in loads:
        for ln in load["lines"]:
            add_row(ln, load["load_id"], "Loaded")
    for ln in unassigned:
        add_row(ln, "-", "UNASSIGNED (no truck met minimum / fleet exhausted)")

    sheet_style_print(ws, len(headers))
    ws.freeze_panes = "A2"
    return ws


def _rounded_bundle(ax, x, y, w, h, accent, hatch):
    pad = min(w, h) * 0.05
    box = patches.FancyBboxPatch(
        (x + pad, y + pad), max(w - 2 * pad, 0.01), max(h - 2 * pad, 0.01),
        boxstyle="round,pad=0,rounding_size=%.4f" % (min(w, h) * 0.06),
        linewidth=1.3, edgecolor=accent, facecolor="white", hatch=hatch, zorder=3,
    )
    ax.add_patch(box)


def _badge(ax, x, y, w, h, pct, label):
    """Ink-light badge: white fill, colored outline + text only (no solid fill)."""
    pct_clamped = max(0, min(pct, 100))
    color = "#2E7D32" if pct_clamped < 70 else ("#B8860B" if pct_clamped < 92 else "#B2323C")
    ax.add_patch(patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0,rounding_size=%.3f" % (h * 0.4),
                                         linewidth=1.1, edgecolor=color, facecolor="white", zorder=3))
    ax.text(x + w / 2, y + h / 2, "%s %.0f%%" % (label, pct_clamped), fontsize=6.2, va="center",
            ha="center", color=color, fontweight="bold", zorder=4)


def draw_trailer(ax, trailer_info, title, style_map):
    spec = trailer_info["spec"]
    length_cap, height_cap = spec["length_m"], spec["height_m"]
    top_pad = height_cap * 0.55
    ax.set_xlim(-0.6, length_cap + 0.3)
    ax.set_ylim(-0.35, height_cap * 2.2 + top_pad)
    ax.axis("off")

    ax.text(length_cap / 2, height_cap * 2.2 + top_pad * 0.68, title,
            fontsize=12, fontweight="bold", ha="center", color=NAVY)

    weight_pct = trailer_info["used_weight"] / (spec["weight_cap_t"] * 1000) * 100
    length_pct = trailer_info["used_length"] / spec["length_m"] * 100
    cube_pct = trailer_info["used_volume"] / trailer_info["cube_cap_m3"] * 100 if trailer_info["cube_cap_m3"] else 0
    badge_w = length_cap * 0.30
    badge_h = top_pad * 0.20
    badge_y = height_cap * 2.2 + top_pad * 0.10
    gap = (length_cap - 3 * badge_w) / 2
    _badge(ax, 0, badge_y, badge_w, badge_h, weight_pct, "WT")
    _badge(ax, badge_w + gap, badge_y, badge_w, badge_h, cube_pct, "CUBE")
    _badge(ax, 2 * (badge_w + gap), badge_y, badge_w, badge_h, length_pct, "FLOOR")

    ax.plot([-0.2, length_cap + 0.1], [-0.18, -0.18], color="#888888", linewidth=1.6, zorder=1)
    for wx in [length_cap * 0.18, length_cap * 0.5, length_cap * 0.82]:
        ax.add_patch(patches.Circle((wx, -0.18), 0.10, facecolor="none", edgecolor="#333333", linewidth=1.2, zorder=1))

    for slot_idx, slot_label in ((0, "LOWER"), (1, "UPPER")):
        y0 = height_cap * 1.2 if slot_idx == 1 else 0
        ax.add_patch(patches.FancyBboxPatch((0, y0), length_cap, height_cap,
                                             boxstyle="round,pad=0,rounding_size=0.08", linewidth=1.3,
                                             edgecolor=STEEL, facecolor="none", zorder=0))
        ax.text(-0.45, y0 + height_cap / 2, slot_label, rotation=90, va="center", ha="center",
                fontsize=7, color=STEEL, fontweight="bold")

    for p in trailer_info["placements"]:
        y_offset = height_cap * 1.2 if p["slot"] == 1 else 0
        placements_same_bay_slot = [q for q in trailer_info["placements"]
                                     if q["bay"] == p["bay"] and q["slot"] == p["slot"] and q["level"] < p["level"]]
        y = y_offset + sum(q["bundle_height_m"] for q in placements_same_bay_slot)
        accent, hatch = style_map.get(p.get("location_code"), (ACCENTS[0], HATCHES[0]))
        _rounded_bundle(ax, p["x"], y, p["bundle_length_m"], p["bundle_height_m"], accent, hatch)
        so_short = p["sales_order"].replace("SFP-", "").replace("-SO", "")
        label = "%s | %s\n%s" % (so_short, p["sku"][-6:], p["delivery_name"][:20])
        h = p["bundle_height_m"]
        fontsize = max(3.6, min(6.4, h * 16))
        clip_box = patches.Rectangle((p["x"], y), p["bundle_length_m"], p["bundle_height_m"], transform=ax.transData)
        ax.text(p["x"] + p["bundle_length_m"] / 2, y + p["bundle_height_m"] / 2, label,
                ha="center", va="center", fontsize=fontsize, color=INK, fontweight="medium",
                zorder=4, linespacing=0.95, clip_path=clip_box, clip_on=True)

    ax.text(length_cap / 2, -0.32, "Length (m)  -->  towards rear", fontsize=7, ha="center", color="#777777")


def _draw_top_legend(fig, load, style_map, rect):
    ax = fig.add_axes(rect)
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    names = {}
    for t in load["packing"].values():
        for p in t["placements"]:
            names[p.get("location_code")] = p.get("delivery_name", "")
    items = sorted(names.items(), key=lambda kv: kv[1])
    n = max(len(items), 1)
    per_row = min(n, 5)
    col_w = 1.0 / per_row
    ax.text(0, 0.92, "DROPS ON THIS LOAD:", fontsize=7.5, fontweight="bold", color=NAVY, va="top")
    for i, (code, name) in enumerate(items):
        row = i // per_row
        col = i % per_row
        x = col * col_w
        y = 0.55 - row * 0.55
        accent, hatch = style_map.get(code, (ACCENTS[0], HATCHES[0]))
        ax.add_patch(patches.FancyBboxPatch((x, y - 0.12), 0.028, 0.20, boxstyle="round,pad=0,rounding_size=0.01",
                                             linewidth=1.0, edgecolor=accent, facecolor="white", hatch=hatch,
                                             clip_on=False))
        ax.text(x + 0.045, y - 0.02, (name or code)[:28], fontsize=6.6, va="center", color="#333333")


def write_schematics_pdf(loads, path):
    with PdfPages(path) as pdf:
        for load in loads:
            n_trailers = len(load["packing"])
            fig = plt.figure(figsize=(11.7, 8.3))
            fig.patch.set_facecolor("white")

            header_ax = fig.add_axes([0, 0.93, 1, 0.07])
            header_ax.axis("off")
            header_ax.add_patch(patches.Rectangle((0, 0), 1, 0.05, transform=header_ax.transAxes,
                                                   facecolor=NAVY, edgecolor="none"))
            header_ax.text(0.01, 0.62, "LOAD %s" % load["load_id"], fontsize=17, fontweight="bold",
                            color=NAVY, va="center", transform=header_ax.transAxes)
            subtitle = ("Site: %s   |   Direction: %s   |   Truck: %s   |   %.1f m3, %.0f kg   |   %s drops, %s lines" % (
                load["site"], group_label(load["group"]), load["truck_type"], load["total_m3"], load["total_kg"],
                load["n_customers"], len(load["lines"])))
            header_ax.text(0.01, 0.22, subtitle, fontsize=9, color="#444444", va="center",
                            transform=header_ax.transAxes)
            if load["pack_leftover"]:
                header_ax.text(0.99, 0.5, "WARNING: %d bundle(s) could not be placed" % len(load["pack_leftover"]),
                                fontsize=9, color="#B2323C", fontweight="bold", ha="right", va="center",
                                transform=header_ax.transAxes)

            style_map = _customer_style_map(load)
            _draw_top_legend(fig, load, style_map, [0.01, 0.80, 0.98, 0.12])

            plot_area_width = 0.98
            names = list(load["packing"].keys())
            spans = [load["packing"][name]["spec"]["length_m"] + 0.9 for name in names]
            total_span = sum(spans)
            x_cursor = 0.01
            trailer_top = 0.74
            for i, name in enumerate(names):
                w = plot_area_width * spans[i] / total_span
                ax = fig.add_axes([x_cursor, 0.06, w - 0.02, trailer_top])
                draw_trailer(ax, load["packing"][name], "%s TRAILER" % name.upper(), style_map)
                x_cursor += w

            pdf.savefig(fig, orientation="landscape")
            plt.close(fig)


if __name__ == "__main__":
    loads, unassigned, fleet_left, lines = build_all()

    wb = openpyxl.Workbook()
    write_loads_summary(wb, loads, unassigned)
    write_orders_summary(wb, loads, unassigned)
    wb.save("Load Building Output.xlsx")
    print("Wrote Load Building Output.xlsx")

    write_schematics_pdf(loads, "Load Schematics.pdf")
    print("Wrote Load Schematics.pdf")

    print("Summary: " + str(len(loads)) + " loads, " + str(len(unassigned)) + " unassigned lines, fleet left: " + str(fleet_left))
