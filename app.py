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

def load_audit_data():
    """Loads historical audit data safely."""
    audit_csv_path = os.path.join(AUDIT_DIR, "audit_summary.csv")
    if os.path.exists(audit_csv_path):
        df = pd.read_csv(audit_csv_path)
        df['Timestamp'] = pd.to_datetime(df['Timestamp'])
        return df
    return pd.DataFrame()

def match_category(text, keyword_list, keyword_to_category, user_threshold):
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
    st.markdown("Automate invoice GL mapping using fuzzy logic and track your processing history.")

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

    # Initialize variables to hold current run data
    download_df = None
    excel_data = None
    run_successful = False

    # =========================
    # PROCESS NEW FILE
    # =========================
    if run_button:
        if not uploaded_file:
            st.error("⚠️ Please upload a file first!")
        elif invoice_col == gl_col:
            st.error("❌ Invoice and GL column cannot be the same.")
        else:
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

                # --- AUDIT TRAIL LOGIC ---
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                os.makedirs(AUDIT_DIR, exist_ok=True)
                
                # 1. Save physical excel file backup
                audit_file_path = os.path.join(AUDIT_DIR, f"mapping_output_{timestamp}.xlsx")
                download_df.to_excel(audit_file_path, index=False)

                # 2. Calculate current run metrics
                total_invoices = len(download_df[invoice_col].unique())
                avg_confidence = download_df["Final_Confidence"].mean()
                unmapped_invoices = download_df[download_df["Final_Category"] == "Others"][invoice_col].nunique()

                # 3. Append to master tracking CSV
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
                run_successful = True

            st.toast('Mapping successful! Results saved to audit log.', icon='🎉')

            # --- CURRENT RUN KPI CARDS (BOX DESIGN) ---
            st.markdown("### ⚡ Current Extraction Results")
            c1, c2, c3 = st.columns(3)
            
            with c1:
                with st.container(border=True):
                    st.metric("Total Unique Invoices", f"{total_invoices:,}")
            with c2:
                with st.container(border=True):
                    st.metric("Average Match Confidence", f"{avg_confidence:.1f}%")
            with c3:
                with st.container(border=True):
                    # Show green check if all mapped, red warning if unmapped remain
                    delta_text = "- Action Recommended" if unmapped_invoices > 0 else "All Mapped! ✅"
                    delta_color = "inverse" if unmapped_invoices > 0 else "normal"
                    st.metric("Unmapped Invoices (Others)", f"{unmapped_invoices:,}", delta=delta_text, delta_color=delta_color)
            
            st.divider()

    # =========================
    # MAIN TABS DISPLAY
    # =========================
    tab1, tab2, tab3 = st.tabs(["📈 Historical Audit", "📄 Current Extraction", "📊 Summary Charts"])

    # --- TAB 1: HISTORICAL AUDIT ---
    with tab1:
        st.markdown("### 🏢 Enterprise Processing History")
        audit_df = load_audit_data()

        if audit_df.empty:
            st.info("ℹ️ No historical audit data found. Run your first mapping to generate logs.")
        else:
            # Historical KPI Cards
            h1, h2, h3 = st.columns(3)
            with h1:
                with st.container(border=True):
                    st.metric("Total Files Processed All-Time", f"{len(audit_df):,}")
            with h2:
                with st.container(border=True):
                    st.metric("Total Invoices Processed All-Time", f"{audit_df['Total_Invoices'].sum():,}")
            with h3:
                with st.container(border=True):
                    all_time_avg = audit_df['Avg_Confidence'].mean()
                    st.metric("All-Time Avg Confidence", f"{all_time_avg:.1f}%")

            # Historical Trend Charts
            st.markdown("#### 📉 Performance Trends")
            chart_col1, chart_col2 = st.columns(2)
            
            # Set index for plotting
            plot_df = audit_df.set_index("Timestamp")

            with chart_col1:
                with st.container(border=True):
                    st.markdown("**Average Confidence (%) Over Time**")
                    st.line_chart(plot_df["Avg_Confidence"], color="#28a745") # Green line

            with chart_col2:
                with st.container(border=True):
                    st.markdown("**Unmapped Invoices Over Time**")
                    st.line_chart(plot_df["Unmapped_Unique_Invoices"], color="#dc3545") # Red line
            
            # Raw Log Table
            st.markdown("#### 📋 Raw Audit Logs")
            st.dataframe(audit_df.sort_values("Timestamp", ascending=False), use_container_width=True, hide_index=True)

    # --- TAB 2: CURRENT EXTRACTION DATA ---
    with tab2:
        if run_successful and download_df is not None:
            st.download_button(
                label="⬇️ Download Output Excel File",
                data=excel_data,
                file_name="gl_category_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            st.write("### Output Data Preview")
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
        else:
            st.info("👈 Please configure settings and click 'Run Mapping' to view current extraction data.")

    # --- TAB 3: SUMMARY CHARTS ---
    with tab3:
        if run_successful and download_df is not None:
            st.subheader("Current Run: Category Distribution")
            summary = (
                download_df["Final_Category"]
                .value_counts()
                .reset_index()
            )
            summary.columns = ["Category", "Count"]

            # Wrap chart in a border
            with st.container(border=True):
                st.bar_chart(summary.set_index("Category"))
        else:
            st.info("👈 Please configure settings and click 'Run Mapping' to view current summary charts.")

# =========================
# ENTRY POINT
# =========================
if __name__ == "__main__":
    main()
