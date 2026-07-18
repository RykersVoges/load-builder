"""
Load-Building web app.
Upload the input Excel, adjust the parameters in the sidebar, click Build Loads,
then download the two output files. Deploy this on Streamlit Community Cloud
(free) to get a shareable URL your team can use -- see the deployment guide
for step-by-step instructions.
"""
import io
import tempfile

import streamlit as st
import openpyxl
import pandas as pd

import load_builder as lb
from build_outputs import (write_loads_summary, write_orders_summary,
                           write_schematics_pdf, write_transport_orders,
                           default_schedule_cfg, DAY_ABBR)

APP_VERSION = "v39 (18 Jul 2026)"

st.set_page_config(page_title="Load Builder", layout="wide")
st.title("Load Builder")
lb_ver = getattr(lb, "LB_VERSION", "v31 or older")
if lb_ver != "v39":
    st.error(f"FILE MISMATCH: load_builder.py on GitHub is {lb_ver}, but this app expects v39. "
             "Re-upload load_builder.py and reboot the app.")
st.caption(f"Upload your orders/customers/SKU/truck workbook, set your parameters, and build loads.  \n**Build {APP_VERSION}, engine {lb_ver}** -- both must say v39, otherwise a file on GitHub is outdated.")

with st.sidebar:
    st.header("Parameters")
    group_mode = st.radio(
        "Group customers by",
        ["Adaptive (recommended)", "Direction (fixed degrees)", "Province"],
        index=0,
        help=("Adaptive seeds each load with the oldest-due order, then pulls in the nearest "
              "other pending orders by direction (not a fixed degree boundary) until the truck "
              "is as full as possible -- generally fills trucks fuller than a fixed bucket. "
              "Direction/Province match the original fixed-bucket instructions exactly."),
    )
    lb.GROUP_MODE = {"Adaptive (recommended)": "adaptive", "Direction (fixed degrees)": "direction",
                      "Province": "province"}[group_mode]
    lb.DIRECTION_DEGREES = st.number_input(
        "Direction bucket size (degrees)", min_value=5, max_value=180, value=30, step=5,
        disabled=(lb.GROUP_MODE != "direction"))
    lb.MAX_CORRIDOR_SPREAD_DEG = st.number_input(
        "Max corridor spread (degrees)", min_value=10, max_value=180, value=60, step=5,
        disabled=(lb.GROUP_MODE != "adaptive"),
        help=("Adaptive mode will never combine freight whose bearing from the site differs "
              "from the load's seed order by more than this, no matter how empty the truck "
              "still is. Prevents e.g. Limpopo and Cape Town freight ending up on the same "
              "load just because they happen to share a similar-looking bearing from the site "
              "-- a load ships under-full instead of reaching that far. Lower = tighter, more "
              "geographically sensible loads but more (smaller) loads overall."))
    lb.MAX_DROPS_PER_LOAD = st.number_input(
        "Max customer drops per load", min_value=1, max_value=20, value=10,
        help=("Raising this lets adaptive mode pull more nearby orders into the same truck "
              "before calling it full -- directly raises volume utilisation. Tested against "
              "your real data: 5 drops = 9 loads (highest 64.2 m3), 10 drops = 7 loads "
              "(highest 64.2 m3, 2 loads over 50 m3), 12 drops = 6 loads (3 loads over 50 m3, "
              "one at 51.4 m3). Fewer, fuller trucks."))

    lb.OVERHANG_ALLOW = st.number_input(
        "Max overhang %", min_value=0, max_value=50, value=15, step=5,
        help=("How far a bundle may stick out past the stack supporting it, as a % of the "
              "supporting surface's length. 0 = every bundle fully supported. Around 15% is "
              "realistic; too much risks bundles bending or tipping.")) / 100.0
    lb.MIN_WT_UTIL_PCT = st.number_input(
        "Min weight utilisation % to dispatch", min_value=0, max_value=100, value=75, step=5,
        help=("Only affects WHICH TRUCK SIZE gets picked when more than one type still has "
              "fleet available -- it prefers a truck this batch fills well. It does NOT strand "
              "freight: as long as some enabled truck type still has fleet left and the batch "
              "physically fits, that truck is used regardless of this %. Freight only goes "
              "unassigned when every enabled truck type has run out, or the batch is too big "
              "for even the largest one."))
    lb.MIN_VOL_UTIL_PCT = st.number_input(
        "Min volume utilisation % to dispatch", min_value=0, max_value=100, value=75, step=5,
        help="Alternative dispatch threshold: % of the truck's cube capacity (same caveat as above).")

    st.subheader("Demand Buckets")
    use_demand_buckets = st.checkbox(
        "Include Demand Buckets / Transport Orders tab", value=True,
        help=("Ticked (default): builds a 'Transport Orders' tab in the Excel output, matching "
              "every SKU line back to its Demand Bucket ID from the uploaded Demand Buckets tab, "
              "on the loading/delivery schedule set below. Untick if you only want the Loads "
              "Summary and Orders Line Summary tabs -- no Transport Orders tab will be built at all."))

    to_start = st.number_input(
        "First Transport Order number", min_value=1, value=1870584, step=1,
        disabled=not use_demand_buckets,
        help=("The Transport Orders tab numbers TOs sequentially starting here. Set it to "
              "follow on from the last TO number already in your TMS."))

    sched = default_schedule_cfg()
    with st.expander("Loading & delivery schedule", expanded=False):
        sched["offset_days"] = st.number_input(
            "Load in how many days from today", min_value=0, max_value=14, value=2,
            help="Load building today, trucks loaded this many days later.")
        sched["duration_h"] = st.number_input(
            "Load duration (hours)", min_value=0.5, max_value=12.0, value=2.0, step=0.5)
        sched["spacing_h"] = st.number_input(
            "Time between load starts (hours)", min_value=0.5, max_value=24.0, value=2.0, step=0.5,
            help="Start of one load to the start of the next at the same site.")
        sched["deliver_start"] = st.text_input("Delivery window opens (HH:MM)", "08:00")
        sched["deliver_end"] = st.text_input("Delivery window closes (HH:MM)", "17:00")

    for label, code in (("Weza (WSM) loading shifts", "WSM"), ("Langeni (LSM) loading shifts", "LSM")):
        with st.expander(label):
            s = sched["shifts"][code]
            st.caption("Format HH:MM-HH:MM. Leave a shift BLANK if the site doesn't load then. "
                       "A night shift ending earlier than it starts runs past midnight.")
            s["weekday_day"] = st.text_input("Mon-Fri day shift", s["weekday_day"], key=code + "wd")
            s["weekday_night"] = st.text_input("Mon-Fri night shift", s["weekday_night"], key=code + "wn")
            s["sat_day"] = st.text_input("Saturday day shift", s["sat_day"], key=code + "sd")
            s["sat_night"] = st.text_input("Saturday night shift", s["sat_night"], key=code + "sn")
            s["sun_day"] = st.text_input("Sunday day shift", s["sun_day"], key=code + "ud")
            s["sun_night"] = st.text_input("Sunday night shift", s["sun_night"], key=code + "un")
            s["skip_night_days"] = st.multiselect(
                "Days with NO night shift", DAY_ABBR, default=s["skip_night_days"], key=code + "sk")

    st.subheader("Fleet available (number of trucks)")
    lb.TRUCK_TYPES["34T"]["fleet_count"] = st.number_input("34 Ton Tautliner", min_value=0, value=20)
    lb.TRUCK_TYPES["FD"]["fleet_count"] = st.number_input(
        "34 Ton Flat Deck", min_value=0, value=0,
        help=("Dimensions/capacity come from the matching column in the uploaded Truck "
              "Dimensions tab if present (matched on the word 'flat' in the header, whatever "
              "the exact tonnage/name), otherwise a placeholder spec is used. Set to 0 if you "
              "don't have any."))
    lb.TRUCK_TYPES["30T"]["fleet_count"] = st.number_input("30 Ton Tri Axle Tautliner", min_value=0, value=0)
    lb.TRUCK_TYPES["14T"]["fleet_count"] = st.number_input("14 Ton Tautliner", min_value=0, value=0)
    lb.TRUCK_TYPES["8T"]["fleet_count"] = st.number_input("8 Ton Tautliner", min_value=0, value=0)

    build_everything = st.checkbox(
        "Build every load regardless of fleet size (decide what to dispatch afterward)", value=False,
        help=("Unticked (default): a realistic plan limited to the exact truck counts entered "
              "above -- if a fleet runs out, remaining freight is reported as unassigned instead "
              "of silently getting a truck that doesn't exist. Tick to assume unlimited trucks of "
              "any type with at least 1 entered above, so every delivery line ends up in some "
              "load -- nothing goes unassigned just because the fleet count ran out. Review the "
              "utilisation numbers below and decide which loads to actually dispatch."))

uploaded = st.file_uploader("Input workbook (.xlsx)", type=["xlsx"])

# Cache the uploaded bytes into session_state as soon as a file lands, and
# use THAT (not the live `uploaded` widget value) for everything below.
# file_uploader's return value can go back to None on a rerun it wasn't
# itself the trigger for (this has been observed after other widgets --
# e.g. a download button -- cause a rerun); if the rest of the app depended
# on `uploaded` directly, that alone would make the whole page collapse
# back to "upload to get started" even though a file was clearly already
# provided. Caching decouples the two.
if uploaded is not None:
    st.session_state["uploaded_bytes"] = uploaded.getvalue()
    st.session_state["uploaded_name"] = uploaded.name
have_file = "uploaded_bytes" in st.session_state

# Build results live in session_state, NOT inside the button's if-block --
# in Streamlit, EVERY widget interaction (including clicking a button
# further down the page) reruns the whole script, and a plain
# "if st.button(...)" only reads True on the exact run it was clicked. If
# the results were nested inside that block, clicking anything below them
# (like a truck-type override) would make the button re-evaluate to False
# and the entire page would appear to reset. Storing the result means it
# keeps rendering on every later rerun until a new "Build Loads" click
# replaces it.
if have_file and st.button("Build Loads", type="primary"):
    with st.spinner("Building loads..."):
        try:
            sites, customers, skus, orders, excluded = lb.load_workbook_data(
                io.BytesIO(st.session_state["uploaded_bytes"]))
        except ValueError as e:
            st.error(f"Problem with the uploaded workbook: {e}")
            st.stop()
        lines = lb.enrich_lines(orders, customers, skus, sites)

        saved_fleet = {k: lb.TRUCK_TYPES[k]["fleet_count"] for k in lb.TRUCK_TYPES}
        if build_everything:
            for k in lb.TRUCK_TYPES:
                if saved_fleet[k] > 0:
                    lb.TRUCK_TYPES[k]["fleet_count"] = 9999
        try:
            loads, unassigned, fleet_left = lb.assemble_loads(lines)
        finally:
            for k in lb.TRUCK_TYPES:
                lb.TRUCK_TYPES[k]["fleet_count"] = saved_fleet[k]

        for load in loads:
            packing, leftover = lb.pack_load(load)
            load["packing"] = packing
            load["pack_leftover"] = leftover

    st.session_state["result"] = {
        "loads": loads, "unassigned": unassigned, "fleet_left": fleet_left,
        "lines": lines, "sites": sites, "excluded": excluded,
        "build_everything": build_everything,
    }
    st.session_state["truck_overrides"] = {}  # fresh load IDs -- clear any stale overrides

if "result" in st.session_state:
    res = st.session_state["result"]
    loads = res["loads"]
    unassigned = res["unassigned"]
    fleet_left = res["fleet_left"]
    lines = res["lines"]
    sites = res["sites"]
    excluded = res["excluded"]
    build_everything_used = res["build_everything"]

    if "rebuild_message" in st.session_state:
        kind, msg = st.session_state.pop("rebuild_message")
        getattr(st, kind)(msg)

    st.success(f"Built {len(loads)} loads from {len(lines)} delivery lines "
               f"({excluded} collects excluded). {len(unassigned)} lines could not be placed.")

    col1, col2, col3 = st.columns(3)
    col1.metric("Loads built", len(loads))
    col2.metric("Unassigned lines", len(unassigned),
                 help="Delivery lines that couldn't be grouped into any viable load at all -- "
                      "either no enabled truck type could physically fit this freight, or (if "
                      "'Build every load regardless of fleet size' is unticked) the fleet you "
                      "entered above ran out.")
    col3.metric("Fleet remaining",
                 "Ignored (unlimited)" if build_everything_used else sum(fleet_left.values()))

    st.subheader("Loads")
    st.caption(
        "**Group** = the direction/area this load's customers were clustered by (a compass range for "
        "Direction/Adaptive modes, or the province code). **Not physically placed** = bundles that were "
        "assigned to this load on paper but didn't fit once real stacking, weight, and floor-length "
        "constraints were simulated -- 0 is the goal; anything else needs attention before dispatch."
    )
    rows = []
    for load in loads:
        spec = lb.TRUCK_TYPES[load["truck_type"]]
        total_bundles = sum(ln["bundles"] for ln in load["lines"])
        rows.append({
            "Load ID": load["load_id"], "Site": load["site"],
            "Truck": lb.truck_display_name(load["truck_type"]),
            "Group": lb.group_label(load["group"]), "m3": round(load["total_m3"], 1),
            "Vol Util %": round(load["total_m3"] / spec["cube_cap_m3"] * 100),
            "KG": round(load["total_kg"]),
            "Weight Util %": round(load["total_kg"] / (spec["payload_cap_t"] * 1000) * 100),
            "Drops": load["n_customers"], "Lines": len(load["lines"]),
            "Bundles placed": total_bundles - len(load["pack_leftover"]),
            "Not physically placed": len(load["pack_leftover"]),
        })
    st.dataframe(rows, width="stretch")

    st.subheader("Adjust truck type per load (optional)")
    st.caption(
        "Set a different truck type for as many loads as you like below, then click 'Rebuild with "
        "these truck types' once -- it re-simulates packing for every changed load and updates the "
        "table above and both downloads below in one go. Nothing changes until you click Rebuild. "
        "Every truck type is selectable here regardless of the fleet count set in the sidebar -- "
        "that count only controls how many of each type the automatic build assumes are available."
    )
    # Show/store DISPLAY NAMES in the editor (e.g. "34 Ton Flat Deck") but
    # translate back to the internal key ("FD") when applying the rebuild.
    key_to_display = {k: lb.truck_display_name(k) for k in lb.TRUCK_TYPES}
    display_to_key = {v: k for k, v in key_to_display.items()}
    truck_options = list(key_to_display.values())

    # IMPORTANT: st.data_editor's own "edited_rows" tracking is a ONE-SHOT
    # delta -- it reflects an edit only on the exact rerun where that edit
    # happened, and is silently cleared on any LATER rerun (even one with no
    # new interaction at all, e.g. clicking the Rebuild button below). If we
    # rebuilt the editor's input fresh from `loads` every run and only read
    # its return value when Rebuild was clicked, any selection made here
    # would appear to work, then do nothing once Rebuild was actually
    # pressed -- exactly what was reported. So we persist the user's choices
    # into our OWN session_state immediately, and feed that back in as the
    # editor's baseline on every run, instead of trusting the widget to
    # remember it for us.
    if "truck_overrides" not in st.session_state:
        st.session_state["truck_overrides"] = {}
    overrides = st.session_state["truck_overrides"]

    override_df = pd.DataFrame([
        {"Load ID": l["load_id"], "Site": l["site"],
         "Truck type": key_to_display[overrides.get(l["load_id"], l["truck_type"])]}
        for l in loads
    ])
    edited_df = st.data_editor(
        override_df,
        column_config={
            "Load ID": st.column_config.TextColumn("Load ID", disabled=True),
            "Site": st.column_config.TextColumn("Site", disabled=True),
            "Truck type": st.column_config.SelectboxColumn("Truck type", options=truck_options, required=True),
        },
        hide_index=True, width="stretch", key="truck_override_editor",
    )
    for _, r in edited_df.iterrows():
        overrides[r["Load ID"]] = display_to_key.get(r["Truck type"], r["Truck type"])

    if st.button("Rebuild with these truck types"):
        id_to_truck = dict(st.session_state["truck_overrides"])
        changed_ids = [lid for lid, t in id_to_truck.items() if t != next(
            (l["truck_type"] for l in loads if l["load_id"] == lid), t)]
        # Re-plan every CHANGED load against a shared unassigned pool, so a
        # load resized to a bigger truck can pick up nearby freight that an
        # earlier override just evicted (or that was already unassigned),
        # and a load resized to a smaller truck sheds its excess back into
        # that same pool for a later load to potentially pick up.
        shared_unassigned = list(res["unassigned"])
        moved_notes = []
        for load in loads:
            new_truck = id_to_truck.get(load["load_id"], load["truck_type"])
            if new_truck == load["truck_type"]:
                continue
            before_n, before_m3 = len(load["lines"]), load["total_m3"]
            lb.resize_load_for_truck(load, new_truck, shared_unassigned)
            moved_notes.append((load["load_id"], lb.truck_display_name(new_truck),
                                 before_n, len(load["lines"]), before_m3, load["total_m3"]))
        res["unassigned"] = shared_unassigned

        warnings = []
        for load in loads:
            packing, leftover = lb.pack_load(load)
            load["packing"] = packing
            load["pack_leftover"] = leftover
            if leftover:
                warnings.append((load["load_id"], lb.truck_display_name(load["truck_type"]), len(leftover)))
        st.session_state["result"]["loads"] = loads
        st.session_state["result"]["unassigned"] = shared_unassigned
        # NOTE: st.rerun() below immediately aborts this script run, so any
        # st.success/st.warning called here would never actually reach the
        # browser -- stash the message and show it after the rerun instead.
        msg_lines = []
        if moved_notes:
            details = "; ".join(
                f"{lid} -> {t}: {n0} to {n1} lines ({m0:.1f} to {m1:.1f} m3)"
                for lid, t, n0, n1, m0, m1 in moved_notes)
            msg_lines.append(f"Re-planned: {details}.")
        if warnings:
            details = "; ".join(f"{lid} on {t}: {n} bundle(s) would not fit" for lid, t, n in warnings)
            msg_lines.append(f"Some bundles no longer physically fit: {details}. "
                              "Consider a bigger truck for those loads, or move some freight to another load.")
            st.session_state["rebuild_message"] = ("warning", " ".join(msg_lines))
        else:
            msg_lines.append("Every bundle still physically fits with your chosen truck types.")
            st.session_state["rebuild_message"] = ("success", " ".join(msg_lines))
        st.rerun()

    wb = openpyxl.Workbook()
    write_loads_summary(wb, loads, unassigned)
    write_orders_summary(wb, loads, unassigned)
    if use_demand_buckets:
        write_transport_orders(wb, loads, sites, to_start, sched)
    xlsx_buf = io.BytesIO()
    wb.save(xlsx_buf)
    xlsx_buf.seek(0)

    with tempfile.NamedTemporaryFile(suffix=".pdf") as tmp:
        write_schematics_pdf(loads, tmp.name)
        tmp.seek(0)
        pdf_bytes = tmp.read()

    tabs_note = ("Loads Summary, Orders Line Summary, Transport Orders" if use_demand_buckets
                 else "Loads Summary, Orders Line Summary (no Transport Orders tab -- unticked above)")
    st.caption(f"Excel tabs included: {tabs_note}")

    dcol1, dcol2 = st.columns(2)
    dcol1.download_button("Download Load Building Output.xlsx", xlsx_buf,
                           file_name="Load Building Output.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    dcol2.download_button("Download Load Schematics.pdf", pdf_bytes,
                           file_name="Load Schematics.pdf", mime="application/pdf")
elif not have_file:
    st.info("Upload your workbook above to get started.")
