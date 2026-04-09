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
# UI HEADER
# =========================
st.title("📊 GL Category Mapping Tool")

# =========================
# SIDEBAR
# =========================
st.sidebar.header("⚙️ Configuration")

uploaded_file = st.sidebar.file_uploader("Upload Input Excel", type=["xlsx"])

sheet_name = None
system_type = None

if uploaded_file:
    excel_file = pd.ExcelFile(uploaded_file)

    sheet_name = st.sidebar.selectbox(
        "Select Input Sheet",
        excel_file.sheet_names
    )

    system_type = st.sidebar.selectbox(
        "Select Source System",
        ["netsuite", "sage"]
    )

# =========================
# MAIN BUTTON
# =========================
run_button = st.sidebar.button("🚀 Run Mapping")

# =========================
# PROCESS
# =========================
if run_button and uploaded_file:

    progress_bar = st.progress(0)
    status_text = st.empty()

    # Step 1: Load data
    status_text.text("📥 Loading input data...")
    df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    progress_bar.progress(10)

    # Step 2: Clean data
    status_text.text("🧹 Cleaning data...")
    df[GL_COL] = df[GL_COL].apply(clean_text)
    df[INVOICE_COL] = df[INVOICE_COL].astype(str).str.strip().fillna("UNKNOWN")
    progress_bar.progress(25)

    # Step 3: Load master
    status_text.text("📚 Loading master data...")
    master_df = pd.read_excel(MASTER_FILE, sheet_name=system_type)
    keyword_to_category, keyword_list = build_keyword_map(master_df)
    progress_bar.progress(40)

    # Step 4: Matching
    status_text.text("🔍 Matching categories...")

    results = []
    total = len(df)

    for i, text in enumerate(df[GL_COL]):
        results.append(match_category(text, keyword_list, keyword_to_category))

        # Update progress dynamically
        if i % max(1, total // 20) == 0:
            progress_bar.progress(40 + int((i / total) * 40))

    df["Temp_Category"] = [r[0] for r in results]
    df["Confidence"] = [r[1] for r in results]

    progress_bar.progress(80)

    # Step 5: Invoice logic
    status_text.text("📊 Aggregating invoice results...")

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

    progress_bar.progress(100)
    status_text.text("✅ Completed!")

    # =========================
    # DISPLAY
    # =========================
    st.success("🎉 Mapping Completed Successfully!")

    st.dataframe(final_output, use_container_width=True)

    # =========================
    # DOWNLOAD
    # =========================
    st.download_button(
        label="⬇️ Download Result",
        data=final_output.to_csv(index=False),
        file_name="gl_category_output.csv",
        mime="text/csv"
    )

elif run_button and not uploaded_file:
    st.error("⚠️ Please upload a file first!")
