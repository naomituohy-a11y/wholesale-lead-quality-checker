import io
import re
import unicodedata
from typing import Any, Dict, List, Tuple
import datetime

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ============================================================

# Streamlit config

# ============================================================

st.set_page_config(page_title="Wholesale Lead Quality Checker", page_icon="✅", layout="wide")

# ============================================================

# Excel colours

# ============================================================

HEADER_YELLOW = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")
CELL_GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
CELL_BLUE = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
CELL_AMBER = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
CELL_RED = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")

# ============================================================

# Normalisation

# ============================================================

def norm_text(v):
if v is None:
return ""
return re.sub(r"\s+", " ", str(v)).strip().lower()

def norm_key(v):
v = norm_text(v)
v = v.replace("&", " and ")
v = re.sub(r"[\/]", " ", v)
v = re.sub(r"[^a-z0-9\s]", " ", v)
return re.sub(r"\s+", " ", v).strip()

# ============================================================

# Template row detection

# ============================================================

def looks_like_template_row(row):
txt = " ".join([str(x).lower() for x in row])
markers = ["picklist", "leave blank", "integer", "text", "dd/mm/yyyy"]
return sum(1 for m in markers if m in txt) >= 2

# ============================================================

# Synonyms

# ============================================================

SYNONYMS = {
"us": ["usa", "united states"],
"uk": ["gb", "united kingdom"],
"it": ["information technology"],
"vp": ["vice president"],
}

def canonical(v):
k = norm_key(v)
for base, vals in SYNONYMS.items():
if k == base or k in vals:
return base
return k

# ============================================================

# Matching

# ============================================================

def match_value(v, allowed):
k = norm_key(v)
ck = canonical(v)

```
if k in allowed:
    return "Match", 100
if ck in allowed:
    return "Match", 95

m = process.extractOne(k, allowed.keys(), scorer=fuzz.token_sort_ratio)
if m:
    if m[1] >= 85:
        return "Match", m[1]
    elif m[1] >= 70:
        return "Review", m[1]

return "No Match", 0
```

# ============================================================

# UI

# ============================================================

st.title("Wholesale Lead Quality Checker")

master_file = st.file_uploader("Upload Master File", type=["xlsx"])
picklist_file = st.file_uploader("Upload Picklist File", type=["xlsx"])

if master_file and picklist_file:

```
df_master = pd.read_excel(master_file)
df_pick = pd.read_excel(picklist_file)

st.subheader("Preview")
st.dataframe(df_master.head())

if st.button("Run QA"):

    allowed = {}
    for col in df_pick.columns:
        allowed[col] = {}
        for val in df_pick[col]:
            if val and norm_key(val) not in ["picklist", "text", "integer"]:
                allowed[col][norm_key(val)] = val

    results = []

    for _, row in df_master.iterrows():

        if looks_like_template_row(row):
            continue

        country = row.get("Country", "")
        status, score = match_value(country, allowed.get("Country", {}))

        overall = "PASS" if score >= 85 else "REVIEW" if score >= 60 else "FAIL"

        results.append({
            "Country": country,
            "QA_Status": status,
            "QA_Score": score,
            "QA_Overall_Status": overall
        })

    df_out = pd.DataFrame(results)

    st.subheader("Results")
    st.dataframe(df_out)

    # ===== Dynamic filename =====
    original_name = master_file.name
    base_name = original_name.rsplit(".", 1)[0]
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")

    output_filename = f"{base_name}_match_results_{timestamp}.xlsx"

    st.download_button(
        "Download Results",
        df_out.to_csv(index=False),
        file_name=output_filename,
    )
```
