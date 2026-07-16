"""
Generates the two output Excel tabs (Loads Summary, Orders Line Summary)
and one landscape schematic page per load (multi-page PDF), from the
load_builder prototype's results.
"""
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.patches as patches

from load_builder import (
    load_workbook_data, enrich_lines, assemble_loads, pack_load, TRUCK_TYPES
)

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

    # figure out full/partial per sales order across all loads + unassigned
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
        row = [load["load_id"], load["site"], load["truck_type"], load["group"],
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


def draw_trailer(ax, trailer_info, title):
    spec = trailer_info["spec"]
    length_cap, height_cap = spec["length_m"], spec["height_m"]
    ax.set_xlim(0, length_cap)
    ax.set_ylim(0, height_cap * 2.2)  # *2.2 to stack the two width-slots visually
    ax.set_title(title, fontsize=9, fontweight="bold")
    ax.set_xlabel("Length (m)", fontsize=7)
    ax.set_yticks([])
    ax.add_patch(patches.Rectangle((0, 0), length_cap, height_cap, fill=False, edgecolor="black", linewidth=1.5))
    ax.add_patch(patches.Rectangle((0, height_cap * 1.2), length_cap, height_cap, fill=False, edgecolor="black", linewidth=1.5))
    ax.text(-0.3, height_cap / 2, "Slot A", rotation=90, va="center", fontsize=6)
    ax.text(-0.3, height_cap * 1.2 + height_cap / 2, "Slot B", rotation=90, va="center", fontsize=6)

    for p in trailer_info["placements"]:
        y_offset = height_cap * 1.2 if p["slot"] == 1 else 0
        # reconstruct cumulative height at this level within its stack
        placements_same_bay_slot = [q for q in trailer_info["placements"]
                                     if q["bay"] == p["bay"] and q["slot"] == p["slot"] and q["level"] < p["level"]]
        y = y_offset + sum(q["bundle_height_m"] for q in placements_same_bay_slot)
        rect = patches.Rectangle((p["x"], y), p["bundle_length_m"], p["bundle_height_m"],
                                  fill=False, edgecolor="black", linewidth=0.8)
        ax.add_patch(rect)
        label = f"{p['sales_order'].replace('SFP-', '').replace('-SO','')}\n{p['sku'][-6:]}\n{p['delivery_name'][:14]}"
        fontsize = 4.5 if p["bundle_length_m"] < 1.5 else 5.5
        ax.text(p["x"] + p["bundle_length_m"] / 2, y + p["bundle_height_m"] / 2, label,
                ha="center", va="center", fontsize=fontsize)


def write_schematics_pdf(loads, path):
    with PdfPages(path) as pdf:
        for load in loads:
            fig, axes = plt.subplots(1, len(load["packing"]), figsize=(11.7, 8.3))
            if len(load["packing"]) == 1:
                axes = [axes]
            fig.suptitle(
                f"Load {load['load_id']}  |  Site: {load['site']}  |  Truck: {load['truck_type']}  |  "
                f"{load['total_m3']:.1f} m3, {load['total_kg']:.0f} kg  |  "
                f"{load['n_customers']} drops, {len(load['lines'])} lines"
                + (f"   ** {len(load['pack_leftover'])} bundles could NOT be physically placed **"
                   if load["pack_leftover"] else ""),
                fontsize=10, fontweight="bold"
            )
            names = list(load["packing"].keys())
            for ax, name in zip(axes, names):
                draw_trailer(ax, load["packing"][name], f"{name.upper()} trailer")
            plt.tight_layout(rect=[0, 0, 1, 0.92])
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

    print(f"\nSummary: {len(loads)} loads, {len(unassigned)} unassigned lines, fleet left: {fleet_left}")
