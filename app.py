import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz

# =========================
# CONFIG
# =========================
MASTER_FILE = "master_category.xlsx"

INVOICE_COL = "Invoice Number"
GL_COL = "GL Description"

CATEGORY_COL = "Category"
KEYWORDS_COL = "GL Description"

FUZZY_THRESHOLD = 70

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

# =========================
# BUILD KEYWORD MAP
# =========================
def build_keyword_map(master_df):
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

    return keyword_to_category, list(set(keyword_list))

# =========================
# MATCH FUNCTION
# =========================
def match_category(text, keyword_list, keyword_to_category):
    if not text:
        return None, 0

    best_score = 0
    best_keyword = None

    # Exact
    for kw in keyword_list:
        if kw in text:
            return keyword_to_category[kw], 100

    # Fuzzy
    for kw in keyword_list:
        score = fuzz.token_set_ratio(text, kw)
        if score > best_score:
            best_score = score
            best_keyword = kw

    if best_score >= FUZZY_THRESHOLD:
        return keyword_to_category[best_keyword], best_score

    return None, best_score

# =========================
# STREAMLIT UI
# =========================
st.title("📊 GL Category Mapping Tool")

# Upload file
uploaded_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])

if uploaded_file:
    excel_file = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Select Input Sheet", excel_file.sheet_names)

    system_type = st.selectbox("Select Source System", ["netsuite", "sage"])

    if st.button("Run Mapping"):
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

        # Clean columns
        df[GL_COL] = df[GL_COL].apply(clean_text)
        df[INVOICE_COL] = df[INVOICE_COL].astype(str).str.strip().fillna("UNKNOWN")

        # Load master
        master_df = pd.read_excel(MASTER_FILE, sheet_name=system_type)

        keyword_to_category, keyword_list = build_keyword_map(master_df)

        # Apply matching
        results = df[GL_COL].apply(
            lambda x: match_category(x, keyword_list, keyword_to_category)
        )

        df["Temp_Category"] = results.apply(lambda x: x[0])
        df["Confidence"] = results.apply(lambda x: x[1])

        # Invoice level logic
        invoice_result = (
            df.sort_values("Confidence", ascending=False)
              .groupby(INVOICE_COL, as_index=False)
              .first()
        )

        invoice_result = invoice_result.rename(columns={
            "Temp_Category": "Final_Category",
            "Confidence": "Final_Confidence"
        })

        invoice_result["Final_Category"] = invoice_result["Final_Category"].fillna("Others")
        invoice_result["Final_Confidence"] = invoice_result["Final_Confidence"].fillna(0)

        final_output = invoice_result[[INVOICE_COL, "Final_Category", "Final_Confidence"]]

        st.success("✅ Mapping Completed!")

        st.dataframe(final_output)

        # Download button
        st.download_button(
            label="Download Result",
            data=final_output.to_csv(index=False),
            file_name="gl_category_output.csv",
            mime="text/csv"
        )
