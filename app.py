import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
from io import BytesIO

# =========================
# CONFIG
# =========================
# Set page config for a wider layout and browser tab title
st.set_page_config(page_title="GL Mapping Tool", page_icon="📊", layout="wide")

MASTER_FILE = "master_category.xlsx"
CATEGORY_COL = "Category"
KEYWORDS_COL = "GL Description"
FUZZY_THRESHOLD = 70

# =========================
# UTIL FUNCTIONS
# =========================
def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = re.sub(r'[^a-z0-9 ]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

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

# Cache the master data loading so it doesn't read from disk on every interaction
@st.cache_data
def load_master_data(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return build_keyword_map(df)

def match_category(text, keyword_list, keyword_to_category):
    if not text:
        return None, 0

    best_score = 0
    best_keyword = None

    # Exact
    for kw in keyword_list:
        if kw in text:
            return keyword_to_category[kw], 100

    # Fuzzy (no miss)
    for kw in keyword_list:
        score = fuzz.token_set_ratio(text, kw)
        if score > best_score:
            best_score = score
            best_keyword = kw

    if best_score >= FUZZY_THRESHOLD:
        return keyword_to_category[best_keyword], best_score

    return None, best_score

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# =========================
# MAIN APP
# =========================
def main():
    st.title("📊 GL Category Mapping Tool")
    st.markdown("Easily map your invoice GL descriptions to standardized categories using fuzzy matching.")

    # =========================
    # SIDEBAR
    # =========================
    st.sidebar.header("⚙️ Configuration")

    with st.sidebar.container(border=True):
        st.markdown("**📂 Input Settings**")
        uploaded_file = st.file_uploader("Upload Input Excel", type=["xlsx"])

    sheet_name = None
    invoice_col = None
    gl_col = None
    master_sheet = None

    if uploaded_file:
        excel_file = pd.ExcelFile(uploaded_file)
        
        with st.sidebar.container(border=True):
            st.markdown("**📝 Column Mapping**")
            sheet_name = st.selectbox("Select Input Sheet", excel_file.sheet_names)
            temp_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, nrows=0) # Load just headers
            columns = list(temp_df.columns)

            invoice_col = st.selectbox("Select Invoice Column", columns)
            gl_col = st.selectbox("Select GL Description Column", columns)

        with st.sidebar.container(border=True):
            st.markdown("**📚 Master Data**")
            try:
                master_excel = pd.ExcelFile(MASTER_FILE)
                master_sheet = st.selectbox("Select Source System", master_excel.sheet_names)
            except FileNotFoundError:
                st.error(f"Missing '{MASTER_FILE}' in directory.")
                return

    run_button = st.sidebar.button("🚀 Run Mapping", use_container_width=True, type="primary")

    # =========================
    # PROCESS
    # =========================
    if run_button and uploaded_file:

        if invoice_col == gl_col:
            st.error("❌ Invoice and GL column cannot be the same.")
            return

        # 1. Use st.status for a clean loading experience
        with st.status("🚀 Processing Mapping...", expanded=True) as status:
            st.write("📥 Loading user data...")
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

            st.write("🧹 Cleaning text...")
            df[gl_col] = df[gl_col].apply(clean_text)
            df[invoice_col] = df[invoice_col].astype(str).str.strip().fillna("UNKNOWN")

            st.write("📚 Loading master dictionary...")
            keyword_to_category, keyword_list = load_master_data(MASTER_FILE, master_sheet)

            st.write("🔍 Performing fuzzy matching (this may take a moment)...")
            results = []
            for text in df[gl_col]:
                results.append(match_category(text, keyword_list, keyword_to_category))

            df["Temp_Category"] = [r[0] for r in results]
            df["Confidence"] = [r[1] for r in results]

            st.write("📊 Aggregating final results...")
            invoice_result = (
                df.sort_values("Confidence", ascending=False)
                .groupby(invoice_col, as_index=False)
                .first()
            )

            invoice_result = invoice_result.rename(columns={
                "Temp_Category": "Final_Category",
                "Confidence": "Final_Confidence"
            })

            invoice_result["Final_Category"] = invoice_result["Final_Category"].fillna("Others")
            invoice_result["Final_Confidence"] = invoice_result["Final_Confidence"].fillna(0)

            final_df = df.merge(
                invoice_result[[invoice_col, "Final_Category", "Final_Confidence"]],
                on=invoice_col,
                how="left"
            )

            st.write("📦 Preparing output file...")
            download_df = final_df.drop(columns=["Temp_Category", "Confidence"], errors="ignore")
            excel_data = to_excel(download_df)

            status.update(label="✅ Mapping Complete!", state="complete", expanded=False)

        # 2. Trigger non-intrusive toast notification
        st.toast('Mapping successful! You can now review and download the results.', icon='🎉')

        # 3. Display High-Level Metrics
        st.subheader("Results Snapshot")
        col1, col2, col3 = st.columns(3)
        
        total_invoices = len(download_df[invoice_col].unique())
        avg_confidence = download_df["Final_Confidence"].mean()
        unmapped = len(download_df[download_df["Final_Category"] == "Others"])

        col1.metric("Total Unique Invoices", f"{total_invoices:,}")
        col2.metric("Average Match Confidence", f"{avg_confidence:.1f}%")
        col3.metric("Unmapped (Others)", f"{unmapped:,}", delta="- Action Recommended" if unmapped > 0 else None, delta_color="inverse")
        
        st.divider()

        # =========================
        # TABS
        # =========================
        tab1, tab2 = st.tabs(["📄 Output Data", "📊 Summary Dashboard"])

        with tab1:
            st.download_button(
                label="⬇️ Download Excel File",
                data=excel_data,
                file_name="gl_category_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            st.write("### Data Preview")
            # 4. Use st.column_config for visual progress bars in the table
            st.dataframe(
                download_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Final_Confidence": st.column_config.ProgressColumn(
                        "Match Confidence",
                        help="Fuzzy match score from 0 to 100",
                        format="%d%%",
                        min_value=0,
                        max_value=100,
                    ),
                    "Final_Category": st.column_config.TextColumn(
                        "Mapped Category",
                        help="The assigned GL Category"
                    )
                }
            )

        with tab2:
            st.subheader("Category Distribution")
            summary = (
                download_df["Final_Category"]
                .value_counts()
                .reset_index()
            )
            summary.columns = ["Category", "Count"]

            st.bar_chart(summary.set_index("Category"))

    elif run_button and not uploaded_file:
        st.error("⚠️ Please upload a file first!")

# =========================
# ENTRY POINT
# =========================
if __name__ == "__main__":
    main()
