import io
import pandas as pd
import streamlit as st
from rapidfuzz import fuzz
import re

st.set_page_config(page_title="Wholesale QA Tool", layout="wide")

# -----------------------------
# Helpers
# -----------------------------
def norm(x):
    if pd.isna(x):
        return ""
    return str(x).strip().lower()

def extract_email_domain(email):
    if "@" in str(email):
        return email.split("@")[-1].lower()
    return ""

def score_text(a, b):
    return fuzz.token_sort_ratio(norm(a), norm(b))

# -----------------------------
# Column Detection (simple v1)
# -----------------------------
def detect_column(cols, keywords):
    for c in cols:
        for k in keywords:
            if k in c.lower():
                return c
    return None

# -----------------------------
# Picklist Extractor (basic)
# -----------------------------
def extract_picklist_values(df):
    values = {}
    for col in df.columns:
        vals = set()
        for v in df[col]:
            v = str(v).strip()
            if v and not any(x in v.lower() for x in ["picklist", "text", "integer", "leave blank"]):
                vals.add(v)
        if len(vals) > 2:
            values[col] = vals
    return values

# -----------------------------
# Validation Logic
# -----------------------------
def validate_row(row, picklists, cols):
    score = 0
    issues = []

    # Country
    if cols["country"]:
        val = row[cols["country"]]
        if any(score_text(val, x) > 85 for x in picklists.get("country", [])):
            score += 15
        else:
            issues.append("Country")

    # Industry
    if cols["industry"]:
        val = row[cols["industry"]]
        if any(score_text(val, x) > 80 for x in picklists.get("industry", [])):
            score += 15
        else:
            issues.append("Industry")

    # Domain check
    if cols["email"] and cols["company"]:
        domain = extract_email_domain(row[cols["email"]])
        comp = row[cols["company"]]
        if domain and comp and domain.split(".")[0] in comp.lower():
            score += 10
        else:
            issues.append("Domain")

    # Title relevance
    if cols["title"]:
        title = norm(row[cols["title"]])
        if any(x in title for x in ["it", "technology", "information"]):
            score += 15
        else:
            issues.append("Title")

    # Phone basic check
    if cols["phone"]:
        phone = str(row[cols["phone"]])
        if phone.startswith("1-8"):
            issues.append("Toll-Free")
        else:
            score += 5

    return score, issues

# -----------------------------
# UI
# -----------------------------
st.title("Wholesale Lead Quality Checker")

master_file = st.file_uploader("Upload Master File", type=["xlsx"])
picklist_file = st.file_uploader("Upload Picklist File", type=["xlsx"])

if master_file and picklist_file:

    master_df = pd.read_excel(master_file)
    pick_df = pd.read_excel(picklist_file)

    # Detect columns
    cols = {
        "company": detect_column(master_df.columns, ["company"]),
        "email": detect_column(master_df.columns, ["email"]),
        "country": detect_column(master_df.columns, ["country"]),
        "industry": detect_column(master_df.columns, ["industry"]),
        "title": detect_column(master_df.columns, ["title"]),
        "phone": detect_column(master_df.columns, ["phone"])
    }

    st.write("Detected columns:", cols)

    # Extract picklists
    picklists = extract_picklist_values(pick_df)

    # Run validation
    scores = []
    statuses = []
    issues_list = []

    for _, row in master_df.iterrows():
        score, issues = validate_row(row, picklists, cols)
        scores.append(score)
        issues_list.append(", ".join(issues))

        if score >= 40:
            statuses.append("PASS")
        elif score >= 20:
            statuses.append("REVIEW")
        else:
            statuses.append("FAIL")

    master_df["QA_Score"] = scores
    master_df["QA_Status"] = statuses
    master_df["QA_Issues"] = issues_list

    st.dataframe(master_df.head())

    # Download
    output = io.BytesIO()
    master_df.to_excel(output, index=False)
    st.download_button("Download Results", output.getvalue(), "wholesale_results.xlsx")
