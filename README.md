# Sian3L Mode Hopping – Curve & PPT Generator

Streamlit app that turns WRP temperature-sweep raw data into:

1. **Curve Excel** – raw data posted into `Operational_RawData` / `Maximum_RawData` of
   `Sian3L_MH_Chart-Template.xlsx`. The split sheets (1_/5_Operational, 1_/5_Maximum), the Curve
   sheet and its 8 charts update by formula. One workbook per SN:
   `<SN>_Sian3L_MH_Chart_<MM-DD>.xlsx` (several SNs → one zip).
2. **PPT report** – from `Sian3L_Mode_Hopping_Template.pptx`: cover, Summary, and two slides per SN
   (Wavelength & TxP, dWL/dT) at the template's picture positions, with auto conclusions.

## Input
- Upload `.csv`, `.xlsx` or `.zip` (sub-folders and zips inside zips are searched).
- Each SN needs one `*_Operational Current_*` and one `*_Maximum Current_*` file, e.g.
  `A3_Slot1_Operational Current_10-22_13-46_PASS.xlsx`. SN = text before `_Operational` / `_Maximum`.
- The raw table is found by its `[CHNumber]` header row (first sheet that has it for xlsx).

## Auto conclusion rules
- dWL/dT of CH1/CH5 (Operational + Maximum) must stay within **0 ~ 0.2** → otherwise "suspect mode hopping".
- Rows with `[CH_Pass_Fail] = Fail` are listed with their error message and temperature range.
- Both conclusion columns are editable in the app before generating the PPT.

## Run locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Deploy on Streamlit Community Cloud
1. Push this folder to a GitHub repo (keep the `templates/` folder).
2. https://share.streamlit.io → **Create app** → pick the repo, branch `main`, main file `streamlit_app.py` → **Deploy**.

## Files
| File | Purpose |
|---|---|
| `streamlit_app.py` | UI: upload, 2 checkboxes, editable conclusions, 2 downloads |
| `mh_core.py` | parsing, Excel filling (XML level, keeps charts), chart images, PPT build |
| `templates/` | default Excel and PPT templates (can be overridden in the app's *Report settings*) |
