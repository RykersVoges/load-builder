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

st.set_page_config(page_title="Load Builder", layout="wide")
st.title("Load Builder")
st.caption("Upload your orders/customers/SKU/truck workbook, set your parameters, and build loads.")

with st.sidebar:
    st.header("Parameters")
    group_mode = st.radio("Group customers by", ["Direction (degrees)", "Province"], index=0)
    lb.GROUP_MODE = "province" if group_mode == "Province" else "direction"
    lb.DIRECTION_DEGREES = st.number_input("Direction bucket size (degrees)", min_value=5, max_value=180, value=30, step=5,
                                            disabled=(lb.GROUP_MODE == "province"))
    lb.MAX_DROPS_PER_LOAD = st.number_input("Max customer drops per load", min_value=1, max_value=20, value=5)

    st.subheader("Fleet available (number of trucks)")
    lb.TRUCK_TYPES["34T"]["fleet_count"] = st.number_input("34 Ton Tautliner", min_value=0, value=10)
    lb.TRUCK_TYPES["30T"]["fleet_count"] = st.number_input("30 Ton Tri Axle Tautliner", min_value=0, value=0)
    lb.TRUCK_TYPES["14T"]["fleet_count"] = st.number_input("14 Ton Tautliner", min_value=0, value=0)
    lb.TRUCK_TYPES["8T"]["fleet_count"] = st.number_input("8 Ton Tautliner", min_value=0, value=0)

uploaded = st.file_uploader("Input workbook (.xlsx)", type=["xlsx"])

if uploaded is not None:
    if st.button("Build Loads", type="primary"):
        with st.spinner("Building loads..."):
            sites, customers, skus, orders, excluded = lb.load_workbook_data(uploaded)
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
        col2.metric("Unassigned lines", len(unassigned))
        col3.metric("Fleet remaining", sum(fleet_left.values()))

        st.subheader("Loads")
        rows = []
        for load in loads:
            spec = lb.TRUCK_TYPES[load["truck_type"]]
            rows.append({
                "Load ID": load["load_id"], "Site": load["site"], "Truck": load["truck_type"],
                "Group": load["group"], "m3": round(load["total_m3"], 1),
                "Vol Util %": round(load["total_m3"] / spec["cube_cap_m3"] * 100),
                "KG": round(load["total_kg"]),
                "Weight Util %": round(load["total_kg"] / (spec["payload_cap_t"] * 1000) * 100),
                "Drops": load["n_customers"], "Lines": len(load["lines"]),
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
