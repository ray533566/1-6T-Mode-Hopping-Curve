import io
import zipfile
from datetime import date

import pandas as pd
import streamlit as st

import mh_core as core

st.set_page_config(page_title="Sian3L Mode Hopping Report", page_icon="📈", layout="wide")
st.title("📈 Sian3L Mode Hopping – Curve & PPT Generator")
st.caption("Upload WRP raw data (csv / xlsx / zip – any folder depth). Each SN needs one "
           "*Operational Current_* file and one *Maximum Current_* file.")

# ---------------------------------------------------------------- inputs
uploads = st.file_uploader("Raw data files", type=["csv", "xlsx", "xlsm", "zip"], accept_multiple_files=True)

c1, c2 = st.columns(2)
do_curve = c1.checkbox("① Curve Excel – post Raw Data to Sian3L_MH_Chart-Template.xlsx", value=True)
do_ppt = c2.checkbox("② PPT report – auto-generate from Sian3L_Mode_Hopping_Template.pptx", value=True)

with st.expander("Report settings", expanded=False):
    rc1, rc2 = st.columns(2)
    rep_title = rc1.text_input("Cover title", "1.6T Sian3L LRO Mode Hopping Test Report")
    rep_date = rc2.date_input("Report date", date.today())
    purpose = st.text_input("Purpose", "")
    condition = st.text_input("Test Condition", "")
    st.markdown("**Templates** (optional – leave empty to use the built-in ones)")
    tc1, tc2 = st.columns(2)
    xlsx_tpl = tc1.file_uploader("Chart template (.xlsx)", type=["xlsx"], key="xt")
    ppt_tpl = tc2.file_uploader("PPT template (.pptx)", type=["pptx"], key="pt")

if not uploads:
    st.info("Upload files to start.")
    st.stop()


@st.cache_data(show_spinner="Reading raw data…")
def parse(files):
    return core.load_inputs(files)


sns, warns = parse(tuple((f.name, f.getvalue()) for f in uploads))
for w in warns:
    st.warning(w)
complete = {k: v for k, v in sns.items() if "Operational" in v.raw and "Maximum" in v.raw}
if not complete:
    st.error("No SN with both Operational and Maximum raw data was found.")
    st.stop()

# ---------------------------------------------------------------- overview + editable conclusions
rows = []
for sn, d in complete.items():
    a = core.analyze(d)
    wl_c, dw_c = core.default_conclusions(d)
    dw = list(a["dwdt"].values())
    rows.append({
        "Include": True, "SN": sn, "Date": d.date,
        "Opr rows (CH1/CH5)": f'{len(d.channel("Operational", "1"))}/{len(d.channel("Operational", "5"))}',
        "Max rows (CH1/CH5)": f'{len(d.channel("Maximum", "1"))}/{len(d.channel("Maximum", "5"))}',
        "dWL/dT range": f"{min(v[0] for v in dw):.2f} ~ {max(v[1] for v in dw):.2f}" if dw else "",
        "dWL/dT": "PASS" if a["dwdt_pass"] else "FAIL",
        "WL & TxP conclusion": wl_c, "dWL/dT conclusion": dw_c,
    })
st.subheader(f"{len(rows)} SN found")
table = st.data_editor(
    pd.DataFrame(rows), hide_index=True, width="stretch",
    disabled=["SN", "Date", "Opr rows (CH1/CH5)", "Max rows (CH1/CH5)", "dWL/dT range", "dWL/dT"],
    column_config={"WL & TxP conclusion": st.column_config.TextColumn(width="large"),
                   "dWL/dT conclusion": st.column_config.TextColumn(width="large")})
st.caption("Conclusion columns are editable – they go to the 'Conclusion:' line of each slide and the Summary slide.")

chosen = [complete[sn] for sn in table.loc[table["Include"], "SN"]]
concl = {r["SN"]: (r["WL & TxP conclusion"], r["dWL/dT conclusion"]) for _, r in table.iterrows()}

# ---------------------------------------------------------------- run
if st.button("▶ Generate", type="primary", disabled=not (chosen and (do_curve or do_ppt))):
    st.session_state.pop("out", None)
    out = {}
    if do_curve:
        tpl = io.BytesIO(xlsx_tpl.getvalue()) if xlsx_tpl else core.XLSX_TEMPLATE
        files = {}
        with st.spinner("Building Curve workbooks…"):
            for d in chosen:
                if xlsx_tpl:
                    tpl.seek(0)
                files[core.curve_filename(d)] = core.build_curve_xlsx(d, tpl)
        if len(files) == 1:
            (name, data), = files.items()
            out["curve"] = (name, data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for n, b in files.items():
                    z.writestr(n, b)
            dates = sorted({d.date for d in chosen if d.date})
            out["curve"] = (f"Sian3L_MH_Chart_{dates[0] if len(dates) == 1 else rep_date.strftime('%m-%d')}.zip",
                            buf.getvalue(), "application/zip")
    if do_ppt:
        bar = st.progress(0.0, "Building PPT…")
        data = core.build_ppt(
            chosen, conclusions=concl, report_date=rep_date.strftime("%Y/%m/%d"), title=rep_title,
            purpose=purpose, condition=condition,
            template=io.BytesIO(ppt_tpl.getvalue()) if ppt_tpl else core.PPT_TEMPLATE,
            progress=lambda f, sn: bar.progress(f, f"Charts for {sn}"))
        bar.empty()
        out["ppt"] = (core.ppt_filename(rep_date.strftime("%Y%m%d")), data,
                      "application/vnd.openxmlformats-officedocument.presentationml.presentation")
    st.session_state["out"] = out

out = st.session_state.get("out")
if out:
    st.success("Done – download below.")
    d1, d2 = st.columns(2)
    if "curve" in out:
        n, b, m = out["curve"]
        d1.download_button(f"⬇ 1. Curve  ({n})", b, file_name=n, mime=m, width="stretch")
    if "ppt" in out:
        n, b, m = out["ppt"]
        d2.download_button(f"⬇ 2. PPT  ({n})", b, file_name=n, mime=m, width="stretch")
