"""
Load-Building Prototype
========================
Implements the logic described in the Instruction document against the
real data in "Claude Input File.xlsx".

Assembly v2 adds an "adaptive" grouping mode: instead of pre-slicing
customers into fixed-width direction buckets (which can strand volume that
sits just across a bucket boundary), it seeds each new load with the
oldest-due pending order, then greedily pulls in the geographically nearest
other pending orders (by bearing, not a fixed bucket) until the truck is as
full as it can be. It also does a best-effort bin-covering scan (skips an
order that doesn't fit and keeps checking smaller ones after it, instead of
stopping at the first that doesn't fit).

Packing v2 uses a skyline (shelf) algorithm per trailer width-slot instead
of a single-item-per-bay model, so leftover shelf space beside a shorter
stacked bundle stays open for a later, shorter bundle to fill.
"""
import math
import os
from collections import defaultdict

import openpyxl

LB_VERSION = "v24"
INPUT_FILE = "Claude Input File.xlsx"

DIRECTION_DEGREES = 30
GROUP_MODE = "adaptive"   # "adaptive" (recommended), "direction", or "province"
MAX_DROPS_PER_LOAD = 10

TRUCK_TYPES = {
    "34T": {
        "trailers": [
            {"name": "front", "length_m": 6, "weight_cap_t": 12, "width_m": 2.5, "height_m": 2.6},
            {"name": "rear",  "length_m": 12, "weight_cap_t": 24, "width_m": 2.5, "height_m": 2.6},
        ],
        "payload_cap_t": 36, "cube_cap_m3": 117,
        "min_weight_t": 28, "min_vol_m3": 55,
        "fleet_count": 10,
    },
    "30T": {
        "trailers": [{"name": "single", "length_m": 18, "weight_cap_t": 30, "width_m": 2.5, "height_m": 2.6}],
        "payload_cap_t": 30, "cube_cap_m3": 117,
        "min_weight_t": 23, "min_vol_m3": 48,
        "fleet_count": 0,
    },
    "14T": {
        "trailers": [{"name": "single", "length_m": 14, "weight_cap_t": 14, "width_m": 2.5, "height_m": 2.6}],
        "payload_cap_t": 14, "cube_cap_m3": 91,
        "min_weight_t": 11, "min_vol_m3": 20,
        "fleet_count": 0,
    },
    "8T": {
        "trailers": [{"name": "single", "length_m": 8, "weight_cap_t": 8, "width_m": 2.5, "height_m": 2.5}],
        "payload_cap_t": 8, "cube_cap_m3": 50,
        "min_weight_t": 6, "min_vol_m3": 12,
        "fleet_count": 0,
    },
}
TRUCK_TRY_ORDER = ["34T", "30T", "14T", "8T"]

SITE_ALIAS = {"SFP_LSM": "LSM", "SFP_WSM": "WSM"}
SITE_SWAP = {"LSM": "WSM", "WSM": "LSM"}

COMPASS_POINTS = ["N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE",
                  "S", "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW"]


def group_label(group_value):
    """Human-readable label for a group key: a compass bearing range for
    direction mode (int bucket), the province code as-is for province mode,
    or a centroid +/- spread description for adaptive corridors."""
    if isinstance(group_value, tuple) and group_value and group_value[0] == "adaptive":
        _, centroid, spread = group_value
        compass = COMPASS_POINTS[int((centroid % 360) / 22.5) % 16]
        return "~%d° +/-%d° (%s), adaptive corridor" % (round(centroid), round(spread), compass)
    if isinstance(group_value, int):
        lo = group_value * DIRECTION_DEGREES
        hi = lo + DIRECTION_DEGREES
        compass = COMPASS_POINTS[int(((lo + hi) / 2 % 360) / 22.5) % 16]
        return "%d°-%d° (%s)" % (lo, hi, compass)
    return str(group_value)


# Ships with the app; holds the Customer Locations / SKU Bundle Dimensions /
# Site Locations / Truck Dimensions reference tabs so the daily upload only
# needs the orders tab. Resolved relative to THIS file so it's found no
# matter what folder the server happens to run from.
REFERENCE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reference_data.xlsx")
_REF_WB = None


def _find_sheet_or_none(wb, *keywords):
    for name in wb.sheetnames:
        n = " ".join(str(name).strip().lower().split())
        if all(k in n for k in keywords):
            return wb[name]
    return None


def _find_sheet(wb, *keywords, use_reference=True):
    """Find a tab whose name contains all the given keywords, ignoring
    case, extra spaces, and small renames. If the tab isn't in the uploaded
    workbook and it's a reference tab (customers/SKUs/sites), fall back to
    the reference_data.xlsx bundled with the app -- so the daily upload
    only ever needs to contain the orders tab."""
    ws = _find_sheet_or_none(wb, *keywords)
    if ws is not None:
        return ws
    if use_reference:
        import os
        global _REF_WB
        if _REF_WB is None and os.path.exists(REFERENCE_FILE):
            _REF_WB = openpyxl.load_workbook(REFERENCE_FILE, data_only=True)
        if _REF_WB is not None:
            ws = _find_sheet_or_none(_REF_WB, *keywords)
            if ws is not None:
                return ws
    raise ValueError(
        "Could not find a tab matching %r in the uploaded workbook%s. "
        "Tabs found in the upload: %s." % (
            " + ".join(keywords),
            " or in the app's reference data" if use_reference else "",
            ", ".join(wb.sheetnames)))


def _norm(v):
    return " ".join(str(v).strip().lower().replace("³", "3").split()) if v is not None else ""


def _f(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def _site_code(v):
    """SFP-LSM / SFP_LSM / LSM all mean the Langeni mill -- normalise."""
    return str(v).strip().upper().replace("SFP-", "").replace("SFP_", "")


def _tab_columns(ws, must_have):
    """Locate the header row (first row containing every name in must_have,
    ignoring case/spacing) and return (header_row_number, {name: col_index})."""
    limit = min(12, ws.max_row)
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=limit, values_only=True), 1):
        names = [_norm(v) for v in r]
        if all(m in names for m in must_have):
            colmap = {}
            for j, n in enumerate(names):
                if n and n not in colmap:
                    colmap[n] = j
            return i, colmap
    raise ValueError('Could not find a header row containing [%s] on tab "%s".'
                     % (", ".join(must_have), ws.title))


def load_workbook_data(source=None):
    wb = openpyxl.load_workbook(source or INPUT_FILE, data_only=True)

    # ---- Customer Locations: carries BOTH the customers and (in its first
    # few rows, marked Location Type = FACILITY) the loading sites/mills. ----
    ws = _find_sheet(wb, "customer", "location")
    hdr, cm = _tab_columns(ws, ["location code", "lat", "lon"])
    customers, sites = {}, {}
    for r in ws.iter_rows(min_row=hdr + 1, max_row=ws.max_row, values_only=True):
        def g(*names):
            for n in names:
                j = cm.get(n)
                if j is not None and j < len(r):
                    return r[j]
            return None
        code = g("location code")
        lat, lon = _f(g("lat")), _f(g("lon"))
        if code is None or lat is None or lon is None:
            continue
        code_s = str(code).strip()
        desc = g("customer description", "description")
        if _norm(g("location type")) == "facility" or code_s.upper().startswith("SFP"):
            sites[_site_code(code_s)] = {"name": desc, "lat": lat, "lon": lon}
        else:
            customers[code_s] = {
                "area": g("area"), "city": g("city"), "desc": desc,
                "lat": lat, "lon": lon,
                "province": g("province name", "province", "prov"),
            }

    # ---- Site Locations tab is now OPTIONAL: if present (in the upload or
    # the reference file), it fills in any site the FACILITY rows didn't. ----
    ws = _find_sheet_or_none(wb, "site", "location")
    if ws is None:
        import os
        global _REF_WB
        if _REF_WB is None and os.path.exists(REFERENCE_FILE):
            _REF_WB = openpyxl.load_workbook(REFERENCE_FILE, data_only=True)
        if _REF_WB is not None:
            ws = _find_sheet_or_none(_REF_WB, "site", "location")
    if ws is not None:
        try:
            hdr, cm = _tab_columns(ws, ["site code", "lat", "lon"])
            for r in ws.iter_rows(min_row=hdr + 1, max_row=ws.max_row, values_only=True):
                code = r[cm["site code"]] if cm["site code"] < len(r) else None
                lat, lon = _f(r[cm["lat"]]), _f(r[cm["lon"]])
                if code is None or lat is None or lon is None:
                    continue
                name_j = cm.get("site")
                sites.setdefault(_site_code(code), {
                    "name": r[name_j] if name_j is not None else "", "lat": lat, "lon": lon})
        except ValueError:
            pass
    if not sites:
        raise ValueError(
            "No loading sites found. Expected FACILITY rows at the top of the "
            "Customer Locations tab, or a Site Locations tab.")

    # ---- SKU Bundle Dimensions: header-name based; prefers the
    # "incl dunnage" columns when present. ----
    ws = _find_sheet(wb, "sku")
    hdr, cm = _tab_columns(ws, ["sku code"])

    def _col_like(*words):
        for key, j in cm.items():
            if all(w in key for w in words):
                return j
        return None

    c_code = cm["sku code"]
    c_len = _col_like("unit length")
    c_h_inc, c_h = _col_like("bundle height", "inc"), _col_like("bundle height")
    c_w = _col_like("bundle width")
    c_m3_inc, c_m3 = _col_like("bundle cube", "inc"), _col_like("bundle cube")
    c_kg_inc, c_kg = _col_like("bundle kg", "inc"), _col_like("bundle kg")
    # True payload weight = Unit Weight x Qty (NOP) + dunnage. The master's
    # own "Bundle KG" column is density x OUTER bundle volume (including the
    # air between boards), which overstates the timber's real weight ~15-20%.
    c_uw = _col_like("unit weight")
    c_qty = _col_like("qty (nop)") or _col_like("qty nop")
    c_dun = cm.get("dunnage kg")

    def _val(r, j_pref, j_base=None):
        v = r[j_pref] if j_pref is not None and j_pref < len(r) else None
        if v in (None, "", 0) and j_base is not None and j_base < len(r):
            v = r[j_base]
        return v

    skus = {}
    for r in ws.iter_rows(min_row=hdr + 1, max_row=ws.max_row, values_only=True):
        code = r[c_code] if c_code < len(r) else None
        if code is None:
            continue
        uw, qty = _f(_val(r, c_uw)), _f(_val(r, c_qty))
        dun = _f(_val(r, c_dun)) or 0.0
        if uw and qty:
            bundle_kg = uw * qty + dun     # actual timber mass + dunnage
        else:
            bundle_kg = _f(_val(r, c_kg_inc, c_kg))   # fallback: master's column
        skus[str(code).strip()] = {
            "unit_length_mm": _f(_val(r, c_len)),
            "bundle_height_mm": _f(_val(r, c_h_inc, c_h)),
            "bundle_width_mm": _f(_val(r, c_w)),
            "bundle_cubes_m3": _f(_val(r, c_m3_inc, c_m3)),
            "bundle_kg": bundle_kg,
        }

    # Read the orders tab by HEADER NAME, not fixed column positions, so
    # adding/removing/reordering columns in the tab never breaks the app.
    # Orders must come from the daily upload itself -- never fall back to
    # the reference file here, or the app would silently build yesterday's loads.
    ws = _find_sheet(wb, "order", use_reference=False)
    all_rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))

    header_idx, col = None, {}
    for i, r in enumerate(all_rows[:10]):
        names = [_norm(v) for v in r]
        if "due" in names and "sku" in names:
            header_idx = i
            col = {n: j for j, n in enumerate(names) if n}
            break
    if header_idx is None:
        raise ValueError('Could not find the header row on the "Fully Allocated Orders" tab '
                         '(expected a row containing "Due" and "SKU").')

    def _pick(row, *names):
        for n in names:
            if n in col:
                return row[col[n]]
        return None

    required = [("due",), ("sales order",), ("c/d",), ("location code",),
                ("delivery name",), ("site",), ("sku",), ("m3",), ("bundles",)]
    missing = [names[0] for names in required if not any(n in col for n in names)]
    if missing:
        raise ValueError('The "Fully Allocated Orders" tab is missing expected column(s): %s. '
                         'Found columns: %s' % (", ".join(missing), ", ".join(sorted(col))))

    orders = []
    excluded_collects = 0
    for r in all_rows[header_idx + 1:]:
        cd = _pick(r, "c/d")
        sku = _pick(r, "sku")
        if cd == "C":
            excluded_collects += 1
            continue
        if cd != "D" or sku is None:
            continue
        loc = _pick(r, "location code")
        site = _pick(r, "site")
        m3 = _pick(r, "m3")
        bundles = _pick(r, "bundles")
        m3_bundle = _pick(r, "m3 per bundle", "m3/bundle", "m3 bundle")
        if m3_bundle is None and m3 and bundles:
            m3_bundle = m3 / bundles
        orders.append({
            "due": _pick(r, "due"), "sales_order": _pick(r, "sales order"),
            "location_code": str(loc).strip() if loc is not None else "",
            "delivery_name": _pick(r, "delivery name"),
            "area": _pick(r, "area"), "prov": _pick(r, "prov", "province"),
            "site_raw": site, "site": _site_code(site),
            "sku": sku, "m3": m3, "m3_per_bundle": m3_bundle, "bundles": bundles,
        })

    return sites, customers, skus, orders, excluded_collects


def bearing_degrees(lat1, lon1, lat2, lon2):
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dlambda = math.radians(lon2 - lon1)
    x = math.sin(dlambda) * math.cos(phi2)
    y = math.cos(phi1) * math.sin(phi2) - math.sin(phi1) * math.cos(phi2) * math.cos(dlambda)
    theta = math.atan2(x, y)
    return (math.degrees(theta) + 360) % 360


def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def _circular_angle_diff(a, b):
    d = abs(a - b) % 360
    return min(d, 360 - d)


def _circular_mean(angles):
    if not angles:
        return 0.0
    sx = sum(math.sin(math.radians(a)) for a in angles)
    cx = sum(math.cos(math.radians(a)) for a in angles)
    return math.degrees(math.atan2(sx, cx)) % 360


def enrich_lines(orders, customers, skus, sites):
    enriched = []
    for o in orders:
        cust = customers.get(o["location_code"])
        sku = skus.get(o["sku"])
        site = sites.get(o["site"])
        if cust is None or sku is None or site is None:
            continue
        dist_km = haversine_km(site["lat"], site["lon"], cust["lat"], cust["lon"])
        brg = bearing_degrees(site["lat"], site["lon"], cust["lat"], cust["lon"])
        line = dict(o)
        line.update({
            "cust_lat": cust["lat"], "cust_lon": cust["lon"], "province_full": cust["province"],
            "dist_km": dist_km, "bearing": brg,
            "dir_bucket": int(brg // DIRECTION_DEGREES),
            "bundle_length_m": (sku["unit_length_mm"] or 0) / 1000,
            "bundle_height_m": (sku["bundle_height_mm"] or 0) / 1000,
            "bundle_width_m": (sku["bundle_width_mm"] or 0) / 1000,
            "bundle_cubes_m3": sku["bundle_cubes_m3"] or 0,
            "bundle_kg": sku["bundle_kg"] or 0,
        })
        enriched.append(line)
    return enriched


def group_key(line):
    if GROUP_MODE == "province":
        return (line["site"], line["prov"])
    return (line["site"], line["dir_bucket"])


# ---------------------------------------------------------------------------
# Shared batch-filling + truck-selection helpers
# ---------------------------------------------------------------------------
def _greedy_fill_batch(candidates, max_drops):
    """Best-effort bin-covering: scan candidates in priority order, add
    whichever ones fit (skip -- don't stop -- on ones that don't), so a
    smaller order further down the list can still fill space a bigger one
    couldn't. Returns (batch, batch_m3, batch_kg, custset, leftover)."""
    batch, batch_m3, batch_kg, custset = [], 0.0, 0.0, set()
    leftover = []
    cap_m3 = TRUCK_TYPES["34T"]["cube_cap_m3"]
    cap_kg = TRUCK_TYPES["34T"]["payload_cap_t"] * 1000
    for ln in candidates:
        new_custset = custset | {ln["location_code"]}
        would_exceed_drops = len(new_custset) > max_drops
        would_exceed_cap = (batch_m3 + ln["m3"] > cap_m3) or (batch_kg + ln["bundle_kg"] * ln["bundles"] > cap_kg)
        if would_exceed_drops or would_exceed_cap:
            leftover.append(ln)
            continue
        batch.append(ln)
        batch_m3 += ln["m3"]
        batch_kg += ln["bundle_kg"] * ln["bundles"]
        custset = new_custset
    if not batch and candidates:
        ln = candidates[0]
        batch = [ln]
        batch_m3 = ln["m3"]
        batch_kg = ln["bundle_kg"] * ln["bundles"]
        custset = {ln["location_code"]}
        leftover = candidates[1:]
    return batch, batch_m3, batch_kg, custset, leftover


def _choose_truck_type(batch_m3, batch_kg, fleet_left, is_last_batch):
    batch_t = batch_kg / 1000
    for tname in TRUCK_TRY_ORDER:
        spec = TRUCK_TYPES[tname]
        if fleet_left[tname] <= 0:
            continue
        meets_min = batch_t >= spec["min_weight_t"] or batch_m3 >= spec["min_vol_m3"]
        fits = batch_t <= spec["payload_cap_t"] and batch_m3 <= spec["cube_cap_m3"]
        if fits and (meets_min or is_last_batch):
            return tname
    for tname in TRUCK_TRY_ORDER:
        spec = TRUCK_TYPES[tname]
        if fleet_left[tname] > 0 and batch_t <= spec["payload_cap_t"] and batch_m3 <= spec["cube_cap_m3"]:
            return tname
    return None


# ---------------------------------------------------------------------------
# Fixed-bucket assembly (direction / province toggle, as literally specified)
# ---------------------------------------------------------------------------
def _assemble_loads_fixed(lines):
    fleet_left = {k: v["fleet_count"] for k, v in TRUCK_TYPES.items()}
    groups = defaultdict(list)
    for ln in lines:
        groups[group_key(ln)].append(ln)

    loads = []
    unassigned = []
    load_id_counter = 1

    for gkey, glines in groups.items():
        glines.sort(key=lambda x: (x["due"] is None, x["due"]))
        pending = list(glines)

        while pending:
            batch, batch_m3, batch_kg, custset, leftover = _greedy_fill_batch(pending, MAX_DROPS_PER_LOAD)
            is_last_batch = len(leftover) == 0
            chosen_truck = _choose_truck_type(batch_m3, batch_kg, fleet_left, is_last_batch)

            if chosen_truck is None:
                unassigned.extend(batch)
                pending = leftover
                continue

            fleet_left[chosen_truck] -= 1
            loads.append({
                "load_id": f"L{load_id_counter:03d}",
                "site": gkey[0],
                "group": gkey[1],
                "truck_type": chosen_truck,
                "lines": batch,
                "total_m3": batch_m3,
                "total_kg": batch_kg,
                "n_customers": len(custset),
            })
            load_id_counter += 1
            pending = leftover

    return loads, unassigned, fleet_left


# ---------------------------------------------------------------------------
# Adaptive assembly: seed each load with the oldest-due pending order, then
# pull in the nearest-bearing other pending orders (regardless of a fixed
# bucket boundary) until the truck is as full as it can be.
# ---------------------------------------------------------------------------
def _assemble_loads_adaptive(lines):
    fleet_left = {k: v["fleet_count"] for k, v in TRUCK_TYPES.items()}
    by_site = defaultdict(list)
    for ln in lines:
        by_site[ln["site"]].append(ln)

    loads = []
    unassigned = []
    load_id_counter = 1

    for site, site_lines in by_site.items():
        pending = list(site_lines)

        while pending:
            pending.sort(key=lambda x: (x["due"] is None, x["due"]))
            seed_bearing = pending[0]["bearing"]
            candidates = sorted(pending, key=lambda x: _circular_angle_diff(x["bearing"], seed_bearing))

            batch, batch_m3, batch_kg, custset, leftover = _greedy_fill_batch(candidates, MAX_DROPS_PER_LOAD)
            batch_ids = {id(x) for x in batch}
            new_pending = [x for x in pending if id(x) not in batch_ids]
            is_last_batch = len(new_pending) == 0

            chosen_truck = _choose_truck_type(batch_m3, batch_kg, fleet_left, is_last_batch)

            if chosen_truck is None:
                unassigned.extend(batch)
                pending = new_pending
                continue

            fleet_left[chosen_truck] -= 1
            bearings = [ln["bearing"] for ln in batch]
            centroid = _circular_mean(bearings)
            spread = max((_circular_angle_diff(b, centroid) for b in bearings), default=0)
            loads.append({
                "load_id": f"L{load_id_counter:03d}",
                "site": site,
                "group": ("adaptive", centroid, spread),
                "truck_type": chosen_truck,
                "lines": batch,
                "total_m3": batch_m3,
                "total_kg": batch_kg,
                "n_customers": len(custset),
            })
            load_id_counter += 1
            pending = new_pending

    return loads, unassigned, fleet_left


def assemble_loads(lines):
    if GROUP_MODE == "adaptive":
        return _assemble_loads_adaptive(lines)
    return _assemble_loads_fixed(lines)


OVERHANG_ALLOW = 0.15


def _new_trailer_state(trailer_spec):
    length_cap = trailer_spec["length_m"]
    return {
        "skyline": [
            [{"x0": 0.0, "x1": length_cap, "h": 0.0}],
            [{"x0": 0.0, "x1": length_cap, "h": 0.0}],
        ],
        "used_weight": 0.0, "used_volume": 0.0, "placements": [],
    }


def _try_place_unit(state, trailer_spec, u):
    """Best-fit skyline placement: finds the tightest open shelf position
    across both width-slots. Splits the shelf so any leftover length beside
    a shorter bundle stays open for a later (shorter) bundle to fill."""
    height_cap = trailer_spec["height_m"]
    weight_cap = trailer_spec["weight_cap_t"] * 1000
    if state["used_weight"] + u["bundle_kg"] > weight_cap:
        return False

    best = None
    for slot_idx in (0, 1):
        for seg_index, seg in enumerate(state["skyline"][slot_idx]):
            seg_len = seg["x1"] - seg["x0"]
            if seg_len <= 1e-9:
                continue
            if seg["h"] + u["bundle_height_m"] > height_cap + 1e-9:
                continue
            on_floor = seg["h"] <= 1e-9
            max_allowed = seg_len if on_floor else seg_len * (1 + OVERHANG_ALLOW)
            if u["bundle_length_m"] > max_allowed + 1e-9:
                continue
            overhang_amt = max(0.0, u["bundle_length_m"] - seg_len)
            resulting_height = seg["h"] + u["bundle_height_m"]
            waste = abs(seg_len - u["bundle_length_m"])
            score = (1 if overhang_amt > 1e-9 else 0, round(resulting_height, 4), waste)
            if best is None or score < best[0]:
                best = (score, slot_idx, seg_index)

    if best is None:
        return False

    _, slot_idx, seg_index = best
    seg = state["skyline"][slot_idx][seg_index]
    x0, x1, h = seg["x0"], seg["x1"], seg["h"]
    seg_len = x1 - x0
    claim_len = min(u["bundle_length_m"], seg_len)

    new_segments = [{"x0": x0, "x1": x0 + claim_len, "h": h + u["bundle_height_m"]}]
    if claim_len < seg_len - 1e-9:
        new_segments.append({"x0": x0 + claim_len, "x1": x1, "h": h})
    state["skyline"][slot_idx][seg_index:seg_index + 1] = new_segments

    state["used_weight"] += u["bundle_kg"]
    state["used_volume"] += u.get("bundle_cubes_m3", 0)
    state["placements"].append({"slot": slot_idx, "x": x0, "y": h, **u})
    return True


def _used_length(state):
    """Average occupied floor length across BOTH width-slots (not just the
    furthest reach of whichever slot has the most freight in it) -- a
    trailer with one lane full and the other lane empty is genuinely only
    ~50% floor-utilised, not 100%, since the floor is 2 lanes wide."""
    total = 0.0
    for slot in state["skyline"]:
        total += sum(seg["x1"] - seg["x0"] for seg in slot if seg["h"] > 1e-9)
    return total / len(state["skyline"])


def _expand_units(bundles):
    units = []
    for b in bundles:
        for _ in range(int(b["bundles"])):
            units.append(dict(b))
    units.sort(key=lambda u: (-u["bundle_length_m"], -u["bundle_kg"]))
    return units


def pack_trailer(trailer_spec, bundles):
    """Single-trailer packer (kept for compatibility / single-trailer trucks)."""
    units = _expand_units(bundles)
    state = _new_trailer_state(trailer_spec)
    leftover = []
    for u in units:
        if not _try_place_unit(state, trailer_spec, u):
            leftover.append(u)
    return state["placements"], leftover, _used_length(state), state["used_weight"], state["used_volume"]


def pack_load(load):
    spec = TRUCK_TYPES[load["truck_type"]]
    trailers = spec["trailers"]
    total_trailer_length = sum(t["length_m"] for t in trailers)

    bundle_recs = []
    for ln in load["lines"]:
        bundle_recs.append({
            "bundle_length_m": ln["bundle_length_m"], "bundle_height_m": ln["bundle_height_m"],
            "bundle_width_m": ln["bundle_width_m"], "bundle_kg": ln["bundle_kg"],
            "bundle_cubes_m3": ln["bundle_cubes_m3"],
            "bundles": ln["bundles"], "sales_order": ln["sales_order"], "sku": ln["sku"],
            "delivery_name": ln["delivery_name"], "location_code": ln["location_code"],
            "dist_km": ln["dist_km"],
        })

    if len(trailers) == 1:
        placements, leftover, used_len, used_wt, used_vol = pack_trailer(trailers[0], bundle_recs)
        return {trailers[0]["name"]: {
            "placements": placements, "spec": trailers[0],
            "used_length": used_len, "used_weight": used_wt, "used_volume": used_vol,
            "cube_cap_m3": spec["cube_cap_m3"],
        }}, leftover

    front_spec = next(t for t in trailers if t["name"] == "front")
    rear_spec = next(t for t in trailers if t["name"] == "rear")
    front_state = _new_trailer_state(front_spec)
    rear_state = _new_trailer_state(rear_spec)

    units = _expand_units(bundle_recs)

    # Balanced two-trailer allocation: give every DROP a presence in BOTH
    # trailers, in proportion to each trailer's weight capacity (1/3 front,
    # 2/3 rear on a 34T interlink). As the route is unloaded drop by drop,
    # both trailers then empty gradually together -- avoiding the risk case
    # of a nearly-empty front pulling a still-full 12m rear trailer.
    front_share = front_spec["weight_cap_t"] / (front_spec["weight_cap_t"] + rear_spec["weight_cap_t"])
    drop_front_kg = defaultdict(float)
    drop_total_kg = defaultdict(float)

    leftover = []
    for u in units:
        d = u["location_code"]
        want_front = drop_front_kg[d] < front_share * (drop_total_kg[d] + u["bundle_kg"])
        order = ((front_state, front_spec), (rear_state, rear_spec)) if want_front \
            else ((rear_state, rear_spec), (front_state, front_spec))
        placed_state = None
        for st, sp in order:
            if _try_place_unit(st, sp, u):
                placed_state = st
                break
        if placed_state is None:
            leftover.append(u)
            continue
        drop_total_kg[d] += u["bundle_kg"]
        if placed_state is front_state:
            drop_front_kg[d] += u["bundle_kg"]

    front_cube_cap = spec["cube_cap_m3"] * front_spec["length_m"] / total_trailer_length
    rear_cube_cap = spec["cube_cap_m3"] * rear_spec["length_m"] / total_trailer_length

    return {
        "front": {"placements": front_state["placements"], "spec": front_spec,
                  "used_length": _used_length(front_state), "used_weight": front_state["used_weight"],
                  "used_volume": front_state["used_volume"], "cube_cap_m3": front_cube_cap},
        "rear": {"placements": rear_state["placements"], "spec": rear_spec,
                 "used_length": _used_length(rear_state), "used_weight": rear_state["used_weight"],
                 "used_volume": rear_state["used_volume"], "cube_cap_m3": rear_cube_cap},
    }, leftover


if __name__ == "__main__":
    sites, customers, skus, orders, excluded = load_workbook_data()
    print(f"Sites: {len(sites)}  Customers: {len(customers)}  SKUs: {len(skus)}")
    print(f"Delivery lines loaded: {len(orders)}  (collects excluded: {excluded})")
    lines = enrich_lines(orders, customers, skus, sites)
    print(f"Enriched lines (joined ok): {len(lines)}")

    loads, unassigned, fleet_left = assemble_loads(lines)
    print(f"\nMode: {GROUP_MODE}")
    print(f"Loads built: {len(loads)}")
    print(f"Unassigned lines: {len(unassigned)}")
    print(f"Fleet remaining: {fleet_left}")

    total_leftover = 0
    for load in loads:
        packing, leftover = pack_load(load)
        load["packing"] = packing
        load["pack_leftover"] = leftover
        total_leftover += len(leftover)
        util_m3 = load["total_m3"] / TRUCK_TYPES[load["truck_type"]]["cube_cap_m3"] * 100
        util_kg = load["total_kg"] / (TRUCK_TYPES[load["truck_type"]]["payload_cap_t"] * 1000) * 100
        print("  %s site=%s group=%s truck=%s m3=%.1f (%.0f%%) kg=%.0f (%.0f%%) drops=%d not_placed=%d"
              % (load["load_id"], load["site"], group_label(load["group"]), load["truck_type"],
                 load["total_m3"], util_m3, load["total_kg"], util_kg,
                 load["n_customers"], len(leftover)))

    print("Total bundles not physically placed: %d" % total_leftover)
