import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
from io import BytesIO
import os
from datetime import datetime

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="GL Mapping Tool", page_icon="📊", layout="wide")

MASTER_FILE = "master_category.xlsx"
CATEGORY_COL = "Category"
KEYWORDS_COL = "GL Description"

# Audit Config
AUDIT_DIR = "audit_logs"

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

@st.cache_data
def load_master_data(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    return build_keyword_map(df)

def match_category(text, keyword_list, keyword_to_category, user_threshold):
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

    if best_score >= user_threshold:
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
            temp_df = pd.read_excel(uploaded_file, sheet_name=sheet_name, nrows=0)
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
                
    with st.sidebar.container(border=True):
        st.markdown("**⚙️ Advanced Settings**")
        user_threshold = st.slider(
            "Fuzzy Match Threshold", 
            min_value=0, 
            max_value=100, 
            value=70, 
            step=1,
            help="Higher values demand more exact matches. Lower values catch more variations but may increase false positives."
        )

    run_button = st.sidebar.button("🚀 Run Mapping", use_container_width=True, type="primary")

    # =========================
    # PROCESS
    # =========================
    if run_button and uploaded_file:

        if invoice_col == gl_col:
            st.error("❌ Invoice and GL column cannot be the same.")
            return

        with st.status("🚀 Processing Mapping...", expanded=True) as status:
            st.write("📥 Loading user data...")
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

            st.write("🧹 Cleaning text...")
            df[gl_col] = df[gl_col].apply(clean_text)
            df[invoice_col] = df[invoice_col].astype(str).str.strip().fillna("UNKNOWN")

            st.write("📚 Loading master dictionary...")
            keyword_to_category, keyword_list = load_master_data(MASTER_FILE, master_sheet)

            st.write(f"🔍 Performing fuzzy matching (Threshold: {user_threshold}%)...")
            results = []
            for text in df[gl_col]:
                results.append(match_category(text, keyword_list, keyword_to_category, user_threshold))

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

            st.write("📦 Preparing output file and saving to audit...")
            download_df = final_df.drop(columns=["Temp_Category", "Confidence"], errors="ignore")
            excel_data = to_excel(download_df)

            # =========================
            # AUDIT TRAIL LOGIC & METRICS
            # =========================
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            os.makedirs(AUDIT_DIR, exist_ok=True)
            
            # 1. Save physical excel file backup
            audit_file_path = os.path.join(AUDIT_DIR, f"mapping_output_{timestamp}.xlsx")
            download_df.to_excel(audit_file_path, index=False)

            # 2. Calculate updated metrics (Unmapped is now Unique Invoices)
            total_invoices = len(download_df[invoice_col].unique())
            avg_confidence = download_df["Final_Confidence"].mean()
            
            # UPDATED: Count unique invoice IDs where the final category is "Others"
            unmapped_invoices = download_df[download_df["Final_Category"] == "Others"][invoice_col].nunique()

            # 3. Append metrics to master tracking CSV
            audit_csv_path = os.path.join(AUDIT_DIR, "audit_summary.csv")
            audit_record = pd.DataFrame([{
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Input_Filename": uploaded_file.name,
                "Fuzzy_Threshold": user_threshold,
                "Total_Invoices": total_invoices,
                "Avg_Confidence": round(avg_confidence, 2),
                "Unmapped_Unique_Invoices": unmapped_invoices,
                "Backup_File": f"mapping_output_{timestamp}.xlsx"
            }])

            if os.path.exists(audit_csv_path):
                audit_record.to_csv(audit_csv_path, mode='a', header=False, index=False)
            else:
                audit_record.to_csv(audit_csv_path, index=False)
            
            status.update(label="✅ Mapping Complete & Audit Saved!", state="complete", expanded=False)

        st.toast('Mapping successful! Results saved to audit log.', icon='🎉')

        # Updated KPI Display
        st.subheader("Results Snapshot")
        col1, col2, col3 = st.columns(3)
        
        col1.metric("Total Unique Invoices", f"{total_invoices:,}")
        col2.metric("Average Match Confidence", f"{avg_confidence:.1f}%")
        # Updated label to reflect unique invoices
        col3.metric("Unmapped Invoices (Others)", f"{unmapped_invoices:,}", delta="- Action Recommended" if unmapped_invoices > 0 else None, delta_color="inverse")
        
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
