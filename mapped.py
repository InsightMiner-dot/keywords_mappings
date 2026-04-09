import pandas as pd
import re
import os
from rapidfuzz import process, fuzz
import numpy as np

# =========================
# CONFIG
# =========================
EXTRACTION_FILE = "extraction.xlsx"
EXTRACTION_SHEET = "Sheet1"   # 👈 CHANGE THIS

MASTER_FILE = "master_category.xlsx"
MASTER_SHEET = "Sheet1"       # 👈 CHANGE IF NEEDED

OUTPUT_FOLDER = "output"

INVOICE_COL = "invoice_number"
GL_COL = "gldescription"

CATEGORY_COL = "Category"
KEYWORDS_COL = "Keywords"

# =========================
# CREATE OUTPUT FOLDER
# =========================
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# =========================
# LOAD DATA (WITH SHEET)
# =========================
df = pd.read_excel(EXTRACTION_FILE, sheet_name=EXTRACTION_SHEET)
master_df = pd.read_excel(MASTER_FILE, sheet_name=MASTER_SHEET)

# =========================
# CLEAN TEXT
# =========================
def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = re.sub(r'[^a-z0-9 ]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

df[GL_COL] = df[GL_COL].apply(clean_text)

# =========================
# BUILD KEYWORD MAP
# =========================
keyword_to_category = {}
keyword_list = []

for _, row in master_df.iterrows():
    category = str(row[CATEGORY_COL]).strip()
    keywords = str(row[KEYWORDS_COL]).split(",")

    for kw in keywords:
        kw_clean = clean_text(kw)
        if kw_clean:
            keyword_to_category[kw_clean] = category
            keyword_list.append(kw_clean)

keyword_list = list(set(keyword_list))

# =========================
# MATCH FUNCTION
# =========================
def match_category(text):
    if not text:
        return None, 0, None, None

    # 1. Exact match
    for kw in keyword_list:
        if kw in text:
            return keyword_to_category[kw], 100, kw, "exact"

    # 2. Fuzzy match
    match = process.extractOne(
        text,
        keyword_list,
        scorer=fuzz.token_set_ratio,
        score_cutoff=70
    )

    if match:
        best_kw, score, _ = match
        if score >= 75:
            return keyword_to_category[best_kw], score, best_kw, "fuzzy"

    return None, 0, None, None

# =========================
# APPLY MATCHING
# =========================
results = df[GL_COL].apply(match_category)

df["Temp_Category"] = results.apply(lambda x: x[0])
df["Confidence"] = results.apply(lambda x: x[1])
df["Matched_Keyword"] = results.apply(lambda x: x[2])
df["Match_Type"] = results.apply(lambda x: x[3])

# =========================
# INVOICE LEVEL LOGIC
# =========================
def assign_invoice_category(group):
    matched = group.dropna(subset=["Temp_Category"])

    if len(matched) > 0:
        best_row = matched.sort_values("Confidence", ascending=False).iloc[0]

        return pd.Series({
            "Final_Category": best_row["Temp_Category"],
            "Final_Confidence": best_row["Confidence"],
            "Final_Match_Type": best_row["Match_Type"],
            "Final_Keyword": best_row["Matched_Keyword"]
        })
    else:
        return pd.Series({
            "Final_Category": "Others",
            "Final_Confidence": 0,
            "Final_Match_Type": None,
            "Final_Keyword": None
        })

invoice_result = (
    df.groupby(INVOICE_COL)
      .apply(assign_invoice_category)
      .reset_index()
)

# =========================
# MERGE BACK
# =========================
final_df = df.merge(invoice_result, on=INVOICE_COL, how="left")

# =========================
# OUTPUT FILE NAME LOGIC
# =========================
base_name = os.path.basename(EXTRACTION_FILE)
file_name_without_ext = os.path.splitext(base_name)[0]

output_file_path = os.path.join(
    OUTPUT_FOLDER,
    f"{file_name_without_ext}_glcat.xlsx"
)

# =========================
# SAVE OUTPUT
# =========================
final_df.to_excel(output_file_path, index=False)

print(f"✅ Output saved at: {output_file_path}")
