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

import load_builder as lb
from build_outputs import write_loads_summary, write_orders_summary, write_schematics_pdf

APP_VERSION = "v23 (17 Jul 2026)"

st.set_page_config(page_title="Load Builder", layout="wide")
st.title("Load Builder")
st.caption(f"Upload your orders/customers/SKU/truck workbook, set your parameters, and build loads.  \n**Build {APP_VERSION}** -- if this version number doesn't match the latest one Claude gave you, the app is running old code.")

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
    lb.MAX_DROPS_PER_LOAD = st.number_input(
        "Max customer drops per load", min_value=1, max_value=20, value=10,
        help=("Raising this lets adaptive mode pull more nearby orders into the same truck "
              "before calling it full -- directly raises volume utilisation. Tested against "
              "your real data: 5 drops = 9 loads (highest 64.2 m3), 10 drops = 7 loads "
              "(highest 64.2 m3, 2 loads over 50 m3), 12 drops = 6 loads (3 loads over 50 m3, "
              "one at 51.4 m3). Fewer, fuller trucks."))

    st.subheader("Fleet available (number of trucks)")
    lb.TRUCK_TYPES["34T"]["fleet_count"] = st.number_input("34 Ton Tautliner", min_value=0, value=10)
    lb.TRUCK_TYPES["30T"]["fleet_count"] = st.number_input("30 Ton Tri Axle Tautliner", min_value=0, value=0)
    lb.TRUCK_TYPES["14T"]["fleet_count"] = st.number_input("14 Ton Tautliner", min_value=0, value=0)
    lb.TRUCK_TYPES["8T"]["fleet_count"] = st.number_input("8 Ton Tautliner", min_value=0, value=0)

uploaded = st.file_uploader("Input workbook (.xlsx)", type=["xlsx"])

if uploaded is not None:
    if st.button("Build Loads", type="primary"):
        with st.spinner("Building loads..."):
            try:
                sites, customers, skus, orders, excluded = lb.load_workbook_data(uploaded)
            except ValueError as e:
                st.error(f"Problem with the uploaded workbook: {e}")
                st.stop()
            lines = lb.enrich_lines(orders, customers, skus, sites)
            loads, unassigned, fleet_left = lb.assemble_loads(lines)
            for load in loads:
                packing, leftover = lb.pack_load(load)
                load["packing"] = packing
                load["pack_leftover"] = leftover

        st.success(f"Built {len(loads)} loads from {len(lines)} delivery lines "
                   f"({excluded} collects excluded). {len(unassigned)} lines could not be placed.")

        col1, col2, col3 = st.columns(3)
        col1.metric("Loads built", len(loads))
        col2.metric("Unassigned lines", len(unassigned),
                     help="Delivery lines that couldn't be grouped into any viable load at all -- "
                          "usually because no truck type in the fleet met its minimum utilisation "
                          "for the freight left over, or the fleet ran out.")
        col3.metric("Fleet remaining", sum(fleet_left.values()))

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
                "Load ID": load["load_id"], "Site": load["site"], "Truck": load["truck_type"],
                "Group": lb.group_label(load["group"]), "m3": round(load["total_m3"], 1),
                "Vol Util %": round(load["total_m3"] / spec["cube_cap_m3"] * 100),
                "KG": round(load["total_kg"]),
                "Weight Util %": round(load["total_kg"] / (spec["payload_cap_t"] * 1000) * 100),
                "Drops": load["n_customers"], "Lines": len(load["lines"]),
                "Bundles placed": total_bundles - len(load["pack_leftover"]),
                "Not physically placed": len(load["pack_leftover"]),
            })
        st.dataframe(rows, use_container_width=True)

        wb = openpyxl.Workbook()
        write_loads_summary(wb, loads, unassigned)
        write_orders_summary(wb, loads, unassigned)
        xlsx_buf = io.BytesIO()
        wb.save(xlsx_buf)
        xlsx_buf.seek(0)

        with tempfile.NamedTemporaryFile(suffix=".pdf") as tmp:
            write_schematics_pdf(loads, tmp.name)
            tmp.seek(0)
            pdf_bytes = tmp.read()

        dcol1, dcol2 = st.columns(2)
        dcol1.download_button("Download Load Building Output.xlsx", xlsx_buf,
                               file_name="Load Building Output.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        dcol2.download_button("Download Load Schematics.pdf", pdf_bytes,
                               file_name="Load Schematics.pdf", mime="application/pdf")
else:
    st.info("Upload your workbook above to get started.")
