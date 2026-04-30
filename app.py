import io
import re
import unicodedata
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

import phonenumbers
from phonenumbers.phonenumberutil import NumberParseException

# ============================================================

# Streamlit page config

# ============================================================

st.set_page_config(
page_title="Wholesale Lead Quality Checker",
page_icon="✅",
layout="wide",
)

# ============================================================

# Excel colours

# ============================================================

HEADER_YELLOW = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")
CELL_GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
CELL_BLUE = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
CELL_AMBER = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
CELL_RED = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")

# ============================================================

# Normalisation helpers

# ============================================================

DASH_CHARS = "\u2010\u2011\u2012\u2013\u2014\u2212"

def norm_text(value):
if value is None or (isinstance(value, float) and pd.isna(value)):
return ""
s = str(value)
s = unicodedata.normalize("NFKC", s)
s = s.replace("\u00A0", " ")
s = re.sub(r"\s+", " ", s).strip()
return s.casefold()

def norm_key(value):
x = norm_text(value)
x = x.replace("&", " and ")
for ch in DASH_CHARS:
x = x.replace(ch, " ")
x = x.replace("-", " ")
x = x.replace(",", " ")
x = x.replace("/", " ")
x = x.replace("\", " ")
x = x.replace("’", "'")
x = re.sub(r"[^a-z0-9+#.\s]", " ", x)
x = re.sub(r"\s+", " ", x).strip()
return x

def clean_header(value):
h = "" if value is None else str(value)
h = unicodedata.normalize("NFKC", h)
h = h.replace("\u00A0", " ")
h = h.strip().lower()
h = h.replace("_", " ")
h = re.sub(r"[^a-z0-9]+", " ", h)
h = re.sub(r"\s+", " ", h).strip()
return h

# ============================================================

# Placeholder / template detection

# ============================================================

PLACEHOLDER_PATTERNS = [
r"^\s*$", r"^\s*picklist\s*$", r"^\s*leave blank\s*$", r"^\s*integer\s*$",
r"^\s*text\s*$", r"^\s*dd/mm/yyyy\s*$", r"^\s*https://", r"^\s*no toll",
r"^\s*all accepted", r"^\s*do not map"
]

def is_placeholder(value):
raw = "" if value is None else str(value).strip()
if not raw:
return True
s = norm_text(raw)
for pat in PLACEHOLDER_PATTERNS:
if re.match(pat, s):
return True
return False

def looks_like_template_row(row):
vals = [str(v).strip() for v in row.tolist() if str(v).strip()]
if not vals:
return True

```
joined = " | ".join(vals[:20]).lower()

markers = ["picklist", "leave blank", "integer", "text", "https://", "no toll", "dd/mm/yyyy"]
hits = sum(1 for m in markers if m in joined)

if hits >= 2:
    return True

return False
```

# ============================================================

# Synonyms

# ============================================================

SYNONYM_GROUPS = [
["us", "usa", "united states"],
["uk", "gb", "united kingdom"],
["it", "information technology"],
["vp", "vice president"],
["c level", "c-suite", "chief"],
]

def canonical_key(value):
k = norm_key(value)
for group in SYNONYM_GROUPS:
keys = [norm_key(x) for x in group]
if k in keys:
return keys[0]
return k

# ============================================================

# Matching logic

# ============================================================

def match_value(value, allowed_map):
raw = str(value or "").strip()
key = norm_key(raw)
ckey = canonical_key(raw)

```
if not raw:
    return "Review", "blank", 0

if key in allowed_map:
    return "Match", "exact", 100

if ckey in allowed_map:
    return "Match", "synonym", 95

match = process.extractOne(key, list(allowed_map.keys()), scorer=fuzz.token_sort_ratio)
if match:
    score = int(match[1])
    if score >= 85:
        return "Match", "fuzzy", score
    if score >= 70:
        return "Review", "weak fuzzy", score

return "No Match", "none", 0
```

# ============================================================

# Basic UI

# ============================================================

st.title("Wholesale Lead Quality Checker")

master_file = st.file_uploader("Upload Master File", type=["xlsx"])
picklist_file = st.file_uploader("Upload Picklist File", type=["xlsx"])

if master_file and picklist_file:

```
df_master = pd.read_excel(master_file)
df_pick = pd.read_excel(picklist_file)

st.write("Master Preview")
st.dataframe(df_master.head())

st.write("Picklist Preview")
st.dataframe(df_pick.head())

if st.button("Run QA"):

    allowed = {}

    for col in df_pick.columns:
        allowed[col] = {}
        for v in df_pick[col]:
            if not is_placeholder(v):
                allowed[col][norm_key(v)] = v

    results = []

    for i, row in df_master.iterrows():

        if looks_like_template_row(row):
            continue

        country = row.get("Country", "")
        status, reason, score = match_value(country, allowed.get("Country", {}))

        results.append({
            "Country": country,
            "QA_Country_Status": status,
            "QA_Country_Score": score
        })

    df_out = pd.DataFrame(results)

    st.write("Results")
    st.dataframe(df_out)

    st.download_button(
        "Download Results",
        df_out.to_csv(index=False),
        file_name="results.csv"
    )
```
