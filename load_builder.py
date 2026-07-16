"""
Load-Building Prototype
========================
Implements the logic described in the Instruction document against the
real data in "Claude Input File.xlsx".
"""
import math
from collections import defaultdict

import openpyxl

INPUT_FILE = "Claude Input File.xlsx"

DIRECTION_DEGREES = 30
GROUP_MODE = "direction"
MAX_DROPS_PER_LOAD = 5

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
    direction mode (int bucket), or the province code as-is for province mode."""
    if isinstance(group_value, int):
        lo = group_value * DIRECTION_DEGREES
        hi = lo + DIRECTION_DEGREES
        compass = COMPASS_POINTS[int(((lo + hi) / 2 % 360) / 22.5) % 16]
        return "%d°-%d° (%s)" % (lo, hi, compass)
    return str(group_value)


def load_workbook_data(source=None):
    wb = openpyxl.load_workbook(source or INPUT_FILE, data_only=True)

    ws = wb["Site Locations"]
    sites = {}
    for r in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        if r[1] is None:
            continue
        sites[r[1]] = {"name": r[2], "lat": r[4], "lon": r[5]}

    ws = wb["Customer Locations"]
    customers = {}
    for r in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        code = r[7]
        if code is None:
            continue
        customers[str(code).strip()] = {
            "area": r[4], "city": r[6], "desc": r[10],
            "lat": r[14], "lon": r[16], "province": r[21],
        }

    ws = wb["SKU Bundle Dimensions"]
    skus = {}
    for r in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
        code = r[1]
        if code is None:
            continue
        skus[str(code).strip()] = {
            "unit_length_mm": r[4],
            "bundle_height_mm": r[19] or r[11],
            "bundle_width_mm": r[12],
            "bundle_cubes_m3": r[20] or r[13],
            "bundle_kg": r[21] or r[14],
        }

    ws = wb["Fully Allocated Orders"]
    rows = list(ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True))
    orders = []
    excluded_collects = 0
    for r in rows:
        due, so, cd, loc, name, area, prov, site, sku, m3, m3_bundle, bundles = r[1:13]
        if cd == "C":
            excluded_collects += 1
            continue
        if cd != "D" or sku is None:
            continue
        orders.append({
            "due": due, "sales_order": so, "location_code": str(loc).strip(),
            "delivery_name": name, "area": area, "prov": prov,
            "site_raw": site, "site": SITE_ALIAS.get(site, site),
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


def assemble_loads(lines):
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
            batch, batch_m3, batch_kg, custset = [], 0.0, 0.0, set()
            i = 0
            while i < len(pending):
                ln = pending[i]
                new_custset = custset | {ln["location_code"]}
                would_exceed_drops = len(new_custset) > MAX_DROPS_PER_LOAD
                would_exceed_34t = (batch_m3 + ln["m3"] > TRUCK_TYPES["34T"]["cube_cap_m3"]) or \
                                   (batch_kg + ln["bundle_kg"] * ln["bundles"] > TRUCK_TYPES["34T"]["payload_cap_t"] * 1000)
                if would_exceed_drops or would_exceed_34t:
                    break
                batch.append(ln)
                batch_m3 += ln["m3"]
                batch_kg += ln["bundle_kg"] * ln["bundles"]
                custset = new_custset
                i += 1

            if not batch:
                ln = pending[0]
                batch = [ln]
                batch_m3 = ln["m3"]
                batch_kg = ln["bundle_kg"] * ln["bundles"]
                custset = {ln["location_code"]}
                i = 1

            remaining_after = pending[i:]
            batch_t = batch_kg / 1000

            chosen_truck = None
            is_last_batch_in_group = len(remaining_after) == 0
            for tname in TRUCK_TRY_ORDER:
                spec = TRUCK_TYPES[tname]
                if fleet_left[tname] <= 0:
                    continue
                meets_min = batch_t >= spec["min_weight_t"] or batch_m3 >= spec["min_vol_m3"]
                fits = batch_t <= spec["payload_cap_t"] and batch_m3 <= spec["cube_cap_m3"]
                if fits and (meets_min or is_last_batch_in_group):
                    chosen_truck = tname
                    break
            if chosen_truck is None:
                for tname in TRUCK_TRY_ORDER:
                    spec = TRUCK_TYPES[tname]
                    if fleet_left[tname] > 0 and batch_t <= spec["payload_cap_t"] and batch_m3 <= spec["cube_cap_m3"]:
                        chosen_truck = tname
                        break

            if chosen_truck is None:
                unassigned.extend(batch)
                pending = remaining_after
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
            pending = remaining_after

    return loads, unassigned, fleet_left


OVERHANG_ALLOW = 0.15


def _new_trailer_state():
    return {"bays": [], "used_length": 0.0, "used_weight": 0.0, "used_volume": 0.0, "placements": []}


def _try_place_unit(state, trailer_spec, u):
    """Attempt to place one bundle unit into a trailer's running state.
    Mutates state in place and returns True if it fit."""
    length_cap, height_cap, width_cap, weight_cap = (
        trailer_spec["length_m"], trailer_spec["height_m"],
        trailer_spec["width_m"], trailer_spec["weight_cap_t"] * 1000,
    )
    if state["used_weight"] + u["bundle_kg"] > weight_cap:
        return False
    for bay in state["bays"]:
        for slot_idx in (0, 1):
            if slot_idx == 1 and not bay["two_wide"]:
                continue
            stack = bay["stacks"][slot_idx]
            cur_height = sum(s["bundle_height_m"] for s in stack)
            support_len = stack[-1]["bundle_length_m"] if stack else bay["base_length"]
            if cur_height + u["bundle_height_m"] <= height_cap and \
               u["bundle_length_m"] <= support_len * (1 + OVERHANG_ALLOW):
                stack.append(u)
                state["used_weight"] += u["bundle_kg"]
                state["used_volume"] += u.get("bundle_cubes_m3", 0)
                state["placements"].append({"bay": bay["idx"], "slot": slot_idx,
                                             "x": bay["start_x"], "level": len(stack) - 1, **u})
                return True
    if state["used_length"] + u["bundle_length_m"] <= length_cap:
        bay_idx = len(state["bays"])
        two = width_cap >= 2 * u["bundle_width_m"]
        bay = {"idx": bay_idx, "start_x": state["used_length"], "base_length": u["bundle_length_m"],
               "two_wide": two, "stacks": [[], []]}
        bay["stacks"][0].append(u)
        state["bays"].append(bay)
        state["used_length"] += u["bundle_length_m"]
        state["used_weight"] += u["bundle_kg"]
        state["used_volume"] += u.get("bundle_cubes_m3", 0)
        state["placements"].append({"bay": bay_idx, "slot": 0, "x": bay["start_x"], "level": 0, **u})
        return True
    return False


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
    state = _new_trailer_state()
    leftover = []
    for u in units:
        if not _try_place_unit(state, trailer_spec, u):
            leftover.append(u)
    return state["placements"], leftover, state["used_length"], state["used_weight"], state["used_volume"]


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

    # Two trailers (34T): bias closer-to-site customers into the front trailer
    # (empties first / drop 1) and farther customers into the rear trailer, but
    # let a unit fall back to the OTHER trailer if its preferred one is full --
    # otherwise a lot of freight gets stranded even when the truck as a whole
    # still has spare capacity, purely because of a rigid 50/50 split.
    front_spec = next(t for t in trailers if t["name"] == "front")
    rear_spec = next(t for t in trailers if t["name"] == "rear")
    front_state = _new_trailer_state()
    rear_state = _new_trailer_state()

    units = _expand_units(bundle_recs)  # longest/heaviest first -- stability rule
    dists = sorted(u["dist_km"] for u in units)
    median_dist = dists[len(dists) // 2] if dists else 0

    leftover = []
    for u in units:
        if u["dist_km"] <= median_dist:
            preferred, preferred_spec, other, other_spec = front_state, front_spec, rear_state, rear_spec
        else:
            preferred, preferred_spec, other, other_spec = rear_state, rear_spec, front_state, front_spec
        if _try_place_unit(preferred, preferred_spec, u):
            continue
        if _try_place_unit(other, other_spec, u):
            continue
        leftover.append(u)

    # Cube capacity isn't split by trailer in the source data (only a single
    # truck-level total), so approximate each trailer's share proportional to
    # its length -- both trailers share the same width/height on a 34T truck.
    front_cube_cap = spec["cube_cap_m3"] * front_spec["length_m"] / total_trailer_length
    rear_cube_cap = spec["cube_cap_m3"] * rear_spec["length_m"] / total_trailer_length

    return {
        "front": {"placements": front_state["placements"], "spec": front_spec,
                  "used_length": front_state["used_length"], "used_weight": front_state["used_weight"],
                  "used_volume": front_state["used_volume"], "cube_cap_m3": front_cube_cap},
        "rear": {"placements": rear_state["placements"], "spec": rear_spec,
                 "used_length": rear_state["used_length"], "used_weight": rear_state["used_weight"],
                 "used_volume": rear_state["used_volume"], "cube_cap_m3": rear_cube_cap},
    }, leftover


if __name__ == "__main__":
    sites, customers, skus, orders, excluded = load_workbook_data()
    print(f"Sites: {len(sites)}  Customers: {len(customers)}  SKUs: {len(skus)}")
    print(f"Delivery lines loaded: {len(orders)}  (collects excluded: {excluded})")
    lines = enrich_lines(orders, customers, skus, sites)
    print(f"Enriched lines (joined ok): {len(lines)}")

    loads, unassigned, fleet_left = assemble_loads(lines)
    print(f"\nLoads built: {len(loads)}")
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
        print(f"  {load['load_id']} site={load['site']} group={group_label(load['group'])} truck={load['truck_type']} "
              f"m3={load['total_m3']:.1f} ({util_m3:.0f}%) kg={load['total_kg']:.0f} ({util_kg:.0f}%) "
              f"drops={load['n_customers']} lines={len(load['lines'])} pack_leftover={len(leftover)}")
    print(f"\nTotal bundle units that didn't physically fit (pack_leftover): {total_leftover}")
