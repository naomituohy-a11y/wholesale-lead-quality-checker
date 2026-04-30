import io
import re
import unicodedata
from decimal import Decimal, InvalidOperation
from typing import Any, Dict

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process

# ============================================================

# Streamlit config

# ============================================================

st.set_page_config(page_title="Wholesale QA Tool", layout="wide")

# ============================================================

# Normalisation

# ============================================================

DASH_CHARS = "\u2010\u2011\u2012\u2013\u2014\u2212"

def norm_text(value):
if value is None:
return ""
s = str(value)
s = unicodedata.normalize("NFKC", s)
s = s.replace("\u00A0", " ")
s = re.sub(r"\s+", " ", s).strip()
return s.lower()

def norm_key(value):
x = norm_text(value)
x = x.replace("&", " and ")
for ch in DASH_CHARS:
x = x.replace(ch, " ")
x = x.replace("-", " ")
x = x.replace(",", " ")
x = x.replace("/", " ")
x = x.replace("\", " ")
x = re.sub(r"[^a-z0-9\s]", " ", x)
x = re.sub(r"\s+", " ", x).strip()
return x

# ============================================================

# Template row detection

# ============================================================

def looks_like_template_row(row):
values = " ".join([str(v).lower() for v in row if str(v).strip()])
markers = ["picklist", "leave blank", "integer", "text", "dd/mm/yyyy"]
return sum(1 for m in markers if m in values) >= 2

# ============================================================

# Synonyms

# ============================================================

SYNONYMS = {
"us": ["usa", "united states"],
"uk": ["gb", "united kingdom"],
"it": ["information technology"],
"vp": ["vice president"],
}

def canonical(value):
k = norm_key(value)
for base, vals in SYNONYMS.items():
if k == base or k in vals:
return base
return k

# ============================================================

# Matching

# ============================================================

def match_value(value, allowed):
key = norm_key(value)
ckey = canonical(value)

```
if key in allowed:
    return "Match"

if ckey in allowed:
    return "Match"

match = process.extractOne(key, allowed.keys(), scorer=fuzz.token_sort_ratio)
if match:
    if match[1] >= 85:
        return "Match"
    elif match[1] >= 70:
        return "Review"

return "No Match"
```

# ============================================================

# UI

# ============================================================

st.title("Wholesale Lead Quality Checker")

master = st.file_uploader("Upload Master File", type=["xlsx"])
picklist = st.file_uploader("Upload Picklist File", type=["xlsx"])

if master and picklist:

```
df_master = pd.read_excel(master)
df_pick = pd.read_excel(picklist)

st.subheader("Preview")
st.dataframe(df_master.head())

if st.button("Run QA"):

    # Build allowed values
    allowed = {}
    for col in df_pick.columns:
        allowed[col] = {}
        for val in df_pick[col]:
            if val and str(val).strip().lower() not in ["picklist", "text", "integer"]:
                allowed[col][norm_key(val)] = val

    results = []

    for _, row in df_master.iterrows():

        if looks_like_template_row(row):
            continue

        country = row.get("Country", "")
        status = match_value(country, allowed.get("Country", {}))

        results.append({
            "Country": country,
            "QA_Status": status
        })

    df_out = pd.DataFrame(results)

    st.subheader("Results")
    st.dataframe(df_out)

    st.download_button(
        "Download CSV",
        df_out.to_csv(index=False),
        file_name="qa_results.csv"
    )
```
