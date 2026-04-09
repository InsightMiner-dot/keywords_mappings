import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
from io import BytesIO
import os
from datetime import datetime

# =========================
# DYNAMIC PATH CONFIGURATION
# =========================
st.set_page_config(page_title="GL Mapping Tool", page_icon="📊", layout="wide")

# Automatically detect the folder where this script is running
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Define sub-folders
CONFIG_DIR = os.path.join(BASE_DIR, "config")
AUDIT_DIR = os.path.join(BASE_DIR, "audit_logs")

# Define master file path inside the config folder
MASTER_FILE = os.path.join(CONFIG_DIR, "master_category.xlsx")

CATEGORY_COL = "Category"
KEYWORDS_COL = "GL Description"

# =========================
# AUTO-SETUP FOLDERS
# =========================
def setup_directories():
    """Checks for required folders and creates them if they don't exist."""
    os.makedirs(CONFIG_DIR, exist_ok=True)
    os.makedirs(AUDIT_DIR, exist_ok=True)
    
    # If the master file doesn't exist at all, create a blank template
    if not os.path.exists(MASTER_FILE):
        df_template = pd.DataFrame({
            CATEGORY_COL: ["Example Category (Delete Me)"],
            KEYWORDS_COL: ["example keyword 1, example keyword 2"]
        })
        with pd.ExcelWriter(MASTER_FILE, engine='openpyxl') as writer:
            df_template.to_excel(writer, index=False, sheet_name="Master")

# Run the setup immediately when the script starts
setup_directories()

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

    for kw in keyword_list:
        if kw in text:
            return keyword_to_category[kw], 100

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

    # Show a warning if they are using the auto-generated template
    if os.path.exists(MASTER_FILE):
        temp_df = pd.read_excel(MASTER_FILE)
        if "Example Category (Delete Me)" in temp_df[CATEGORY_COL].values:
            st.warning(f"⚠️ **Action Required:** We created a blank template for your dictionary at `{MASTER_FILE}`. Please open it, add your real categories, and refresh this page.")

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
            except Exception as e:
                st.error(f"Error loading master file: {e}")
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

    run_button = st.sidebar.button("🚀 Run Fast Mapping", use_container_width=True, type="primary")

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

                st.write("📦 Preparing high-speed output file...")
                download_df = final_df.drop(columns=["Temp_Category", "Confidence"], errors="ignore")
                excel_data = to_excel(download_df)

                # --- AUDIT TRAIL LOGIC ---
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                # We know AUDIT_DIR exists because setup_directories() ran at the start
                audit_file_path = os.path.join(AUDIT_DIR, f"mapping_output_{timestamp}.xlsx")
                download_df.to_excel(audit_file_path, index=False)

                # Metrics
                total_invoices = len(download_df[invoice_col].unique())
                avg_confidence = download_df["Final_Confidence"].mean()
                unmapped_invoices = download_df[download_df["Final_Category"] == "Others"][invoice_col].nunique()

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
                
                status.update(label="✅ Mapping Complete!", state="complete", expanded=False)
                run_successful = True

            st.toast('Mapping successful! Results saved to audit log.', icon='🎉')

            # --- CURRENT RUN KPI CARDS ---
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
                    delta_text = "- Action Recommended" if unmapped_invoices > 0 else "All Mapped! ✅"
                    delta_color = "inverse" if unmapped_invoices > 0 else "normal"
                    st.metric("Unmapped Invoices (Others)", f"{unmapped_invoices:,}", delta=delta_text, delta_color=delta_color)
            
            st.divider()

    # =========================
    # MAIN TABS DISPLAY
    # =========================
    tab1, tab2, tab3 = st.tabs(["📈 Historical Audit", "📄 Current Extraction", "📊 Summary Charts"])

    with tab1:
        st.markdown("### 🏢 Enterprise Processing History")
        audit_df = load_audit_data()

        if audit_df.empty:
            st.info("ℹ️ No historical audit data found. Run your first mapping to generate logs.")
        else:
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

            st.markdown("#### 📉 Performance Trends")
            chart_col1, chart_col2 = st.columns(2)
            plot_df = audit_df.set_index("Timestamp")

            with chart_col1:
                with st.container(border=True):
                    st.markdown("**Average Confidence (%) Over Time**")
                    st.line_chart(plot_df["Avg_Confidence"], color="#28a745") 

            with chart_col2:
                with st.container(border=True):
                    st.markdown("**Unmapped Invoices Over Time**")
                    st.line_chart(plot_df["Unmapped_Unique_Invoices"], color="#dc3545") 
            
            st.markdown("#### 📋 Raw Audit Logs")
            st.dataframe(audit_df.sort_values("Timestamp", ascending=False), use_container_width=True, hide_index=True)

    with tab2:
        if run_successful and download_df is not None:
            st.download_button(
                label="⬇️ Download Output Excel File",
                data=excel_data,
                file_name="gl_category_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            st.write("### Target Sheet Data Preview")
            st.dataframe(
                download_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Final_Confidence": st.column_config.ProgressColumn(
                        "Match Confidence",
                        format="%d%%",
                        min_value=0,
                        max_value=100,
                    ),
                    "Final_Category": st.column_config.TextColumn("Mapped Category")
                }
            )
        else:
            st.info("👈 Please configure settings and click 'Run Fast Mapping' to view current extraction data.")

    with tab3:
        if run_successful and download_df is not None:
            st.subheader("Current Run: Category Distribution")
            summary = download_df["Final_Category"].value_counts().reset_index()
            summary.columns = ["Category", "Count"]

            with st.container(border=True):
                st.bar_chart(summary.set_index("Category"))
        else:
            st.info("👈 Please configure settings and click 'Run Fast Mapping' to view current summary charts.")

if __name__ == "__main__":
    main()
