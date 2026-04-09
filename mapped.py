import pandas as pd
import re
import os
from rapidfuzz import fuzz

# =========================
# CONFIG
# =========================
EXTRACTION_FILE = "extraction.xlsx"
EXTRACTION_SHEET = "Sheet1"

MASTER_FILE = "master_category.xlsx"
MASTER_SHEET = "Sheet1"

OUTPUT_FOLDER = "output"

INVOICE_COL = "invoice_number"
GL_COL = "gldescription"

CATEGORY_COL = "Category"
KEYWORDS_COL = "Keywords"

FUZZY_THRESHOLD = 70  # tune if needed

# =========================
# SETUP
# =========================
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

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
# CLEAN INVOICE (VERY IMPORTANT)
# =========================
df[INVOICE_COL] = df[INVOICE_COL].astype(str).str.strip()
df[INVOICE_COL] = df[INVOICE_COL].fillna("UNKNOWN_INVOICE")

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
# MATCH FUNCTION (NO MISS)
# =========================
def match_category(text):
    if not text:
        return None, 0, None, None

    best_score = 0
    best_keyword = None

    # 1. Exact match
    for kw in keyword_list:
        if kw in text:
            return keyword_to_category[kw], 100, kw, "exact"

    # 2. Full fuzzy scan (no miss)
    for kw in keyword_list:
        score = fuzz.token_set_ratio(text, kw)

        if score > best_score:
            best_score = score
            best_keyword = kw

    # 3. Assign category
    if best_score >= FUZZY_THRESHOLD:
        return (
            keyword_to_category[best_keyword],
            best_score,
            best_keyword,
            "fuzzy"
        )

    return None, best_score, best_keyword, "low_match"

# =========================
# APPLY MATCHING
# =========================
results = df[GL_COL].apply(match_category)

df["Temp_Category"] = results.apply(lambda x: x[0])
df["Confidence"] = results.apply(lambda x: x[1])
df["Matched_Keyword"] = results.apply(lambda x: x[2])
df["Match_Type"] = results.apply(lambda x: x[3])

# =========================
# INVOICE LEVEL (NO APPLY)
# =========================
invoice_result = (
    df.sort_values("Confidence", ascending=False)
      .groupby(INVOICE_COL, as_index=False)
      .first()
)

# Rename final columns
invoice_result = invoice_result.rename(columns={
    "Temp_Category": "Final_Category",
    "Confidence": "Final_Confidence",
    "Match_Type": "Final_Match_Type",
    "Matched_Keyword": "Final_Keyword"
})

# Fill no-match invoices
invoice_result["Final_Category"] = invoice_result["Final_Category"].fillna("Others")
invoice_result["Final_Confidence"] = invoice_result["Final_Confidence"].fillna(0)

# =========================
# MERGE BACK
# =========================
final_df = df.merge(
    invoice_result[[INVOICE_COL, "Final_Category", "Final_Confidence",
                    "Final_Match_Type", "Final_Keyword"]],
    on=INVOICE_COL,
    how="left"
)

# =========================
# OUTPUT NAMING
# =========================
base_name = os.path.basename(EXTRACTION_FILE)
file_name = os.path.splitext(base_name)[0]

output_path = os.path.join(
    OUTPUT_FOLDER,
    f"{file_name}_glcat.xlsx"
)

# =========================
# SAVE OUTPUT
# =========================
final_df.to_excel(output_path, index=False)

print(f"✅ Completed! Output saved at: {output_path}")
