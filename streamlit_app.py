import streamlit as st
import pandas as pd
import os
import tempfile
from pathlib import Path
import zipfile
import io

# Import your existing services
from app.services.excel_service import load_excel, detect_numeric_columns
from app.services.gin_service import list_sheets, load_sheet, auto_detect_columns, build_activity_lookup, merge_gin_with_lookup
from app.services.costhead_service import generate_costhead_report
from app.services.export_service import export_excel

# Configure Streamlit page
st.set_page_config(
    page_title="CostHead Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f538d;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #ffffff;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 0.375rem;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 0.375rem;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Main header
    st.markdown('<h1 class="main-header">📊 CostHead Report Generator</h1>', unsafe_allow_html=True)

    # Initialize session state
    if 'activities_df' not in st.session_state:
        st.session_state.activities_df = None
    if 'gin_df' not in st.session_state:
        st.session_state.gin_df = None
    if 'costhead_file' not in st.session_state:
        st.session_state.costhead_file = None
    if 'gst_file' not in st.session_state:
        st.session_state.gst_file = None

    # File upload section in main area
    st.markdown('<div class="section-header">📁 File Upload</div>', unsafe_allow_html=True)

    # Create four columns for file uploads
    upload_col1, upload_col2, upload_col3, upload_col4 = st.columns(4)

    with upload_col1:
        st.markdown("### 📊 Activities (Master)")
        activities_file = st.file_uploader(
            "Upload Activities Excel file",
            type=['xlsx', 'xls'],
            key=f"activities_upload_{st.session_state.get('clear_counter', 0)}",
            help="Upload the master activities Excel file"
        )

        if activities_file is not None:
            try:
                # Load the first sheet automatically
                activities_df = load_sheet(activities_file, list_sheets(activities_file)[0])
                st.session_state.activities_df = activities_df
                st.success(f"✅ Loaded {len(activities_df)} rows")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

    with upload_col2:
        st.markdown("### 📋 GIN (Good Issue Note)")
        gin_file = st.file_uploader(
            "Upload GIN Excel file",
            type=['xlsx', 'xls'],
            key=f"gin_upload_{st.session_state.get('clear_counter', 0)}",
            help="Upload the GIN Excel file"
        )

        if gin_file is not None:
            try:
                # Load the first sheet automatically
                gin_df = load_sheet(gin_file, list_sheets(gin_file)[0])
                st.session_state.gin_df = gin_df
                st.success(f"✅ Loaded {len(gin_df)} rows")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

    with upload_col3:
        st.markdown("### 🗂️ CostHead Mapping")
        costhead_file = st.file_uploader(
            "Upload CostHead mapping Excel file",
            type=['xlsx', 'xls'],
            key=f"costhead_upload_{st.session_state.get('clear_counter', 0)}",
            help="Upload the CostHead mapping Excel file"
        )

        if costhead_file is not None:
            st.session_state.costhead_file = costhead_file
            st.success("✅ CostHead mapping uploaded")

    with upload_col4:
        st.markdown("### 💰 GST Sheet (Optional)")
        gst_file = st.file_uploader(
            "Upload GST Excel file",
            type=['xlsx', 'xls'],
            key=f"gst_upload_{st.session_state.get('clear_counter', 0)}",
            help="Upload the GST sheet with item descriptions and tax slabs (optional)"
        )

        if gst_file is not None:
            st.session_state.gst_file = gst_file
            st.success("✅ GST sheet uploaded")

    st.markdown("<br>", unsafe_allow_html=True)  # Add space

    # Main content area - Data Overview hidden
    # st.markdown('<div class="section-header">📋 Data Overview</div>', unsafe_allow_html=True)

    # Display current status - hidden
    # status_col1, status_col2, status_col3 = st.columns(3)

    # with status_col1:
    #     if st.session_state.activities_df is not None:
    #         st.metric("Activities Rows", len(st.session_state.activities_df))
    #     else:
    #         st.metric("Activities Rows", "Not loaded")

    # with status_col2:
    #     if st.session_state.gin_df is not None:
    #         st.metric("GIN Rows", len(st.session_state.gin_df))
    #     else:
    #         st.metric("GIN Rows", "Not loaded")

    # with status_col3:
    #     if st.session_state.costhead_file is not None:
    #         st.metric("CostHead File", "✅ Ready")
    #     else:
    #         st.metric("CostHead File", "❌ Not loaded")

    # Generate report button - centered and prominent
    st.markdown("<br>", unsafe_allow_html=True)  # Add some space

    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        match_mode = "contains"  # Default to contains mode
        if st.button("🚀 Generate CostHead Report", type="primary", use_container_width=True, key="generate_btn"):
            generate_report(match_mode)

    # Clear button below generate button - only show if files are uploaded
    if (st.session_state.activities_df is not None or
        st.session_state.gin_df is not None or
        st.session_state.costhead_file is not None or
        st.session_state.gst_file is not None):

        st.markdown("<br>", unsafe_allow_html=True)  # Add some space
        clear_col1, clear_col2, clear_col3 = st.columns([1, 1, 1])
        with clear_col2:
            if st.button("🗑️ Clear All Data", use_container_width=True):
                clear_data()

    st.markdown("<br>", unsafe_allow_html=True)  # Add some space

    # Data preview section
    if st.session_state.activities_df is not None or st.session_state.gin_df is not None or st.session_state.gst_file is not None:
        st.markdown('<div class="section-header">👀 Data Preview</div>', unsafe_allow_html=True)

        tab1, tab2, tab3 = st.tabs(["Activities Data", "GIN Data", "GST Data"])

        with tab1:
            if st.session_state.activities_df is not None:
                st.dataframe(st.session_state.activities_df, use_container_width=True)
            else:
                st.info("Upload an activities file to see the preview")

        with tab2:
            if st.session_state.gin_df is not None:
                st.dataframe(st.session_state.gin_df, use_container_width=True)
            else:
                st.info("Upload a GIN file to see the preview")

        with tab3:
            if st.session_state.gst_file is not None:
                try:
                    gst_df = load_sheet(st.session_state.gst_file, list_sheets(st.session_state.gst_file)[0])
                    st.dataframe(gst_df, use_container_width=True)
                except Exception as e:
                    st.error(f"Error loading GST file: {str(e)}")
            else:
                st.info("Upload a GST file to see the preview")

def generate_report(match_mode):
    """Generate the CostHead report"""

    # Check if all required files are loaded
    if st.session_state.activities_df is None:
        st.error("❌ Please upload the Activities file first")
        return

    if st.session_state.gin_df is None:
        st.error("❌ Please upload the GIN file first")
        return

    if st.session_state.costhead_file is None:
        st.error("❌ Please upload the CostHead mapping file first")
        return

    try:
        with st.spinner("🔄 Processing data and generating report..."):
            # Auto-detect columns
            activities_cols = auto_detect_columns(st.session_state.activities_df, "master")
            gin_cols = auto_detect_columns(st.session_state.gin_df, "gin")

            # Validate detected columns
            if not all([activities_cols["code"], activities_cols["name"], activities_cols["wbs"]]):
                st.error("❌ Could not detect required columns in activities sheet")
                return

            if not gin_cols["code"]:
                st.error("❌ Could not detect activity code column in GIN sheet")
                return

            # Build lookup and merge
            activities_lookup = build_activity_lookup(
                st.session_state.activities_df,
                activities_cols["code"],
                activities_cols["wbs"],
                activities_cols["name"]
            )

            merged_df = merge_gin_with_lookup(
                st.session_state.gin_df,
                gin_cols["code"],
                activities_lookup
            )

            # Process GST if GST file is provided
            if st.session_state.gst_file is not None:
                from app.services.gst_service import process_gst_for_gin_mapped
                merged_df = process_gst_for_gin_mapped(merged_df, st.session_state.gst_file)

            # Create temporary directory for output
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save GIN_Mapped file
                gin_mapped_base = os.path.join(temp_dir, "GIN_Mapped")
                gin_mapped_path = export_excel(merged_df, gin_mapped_base)

                # Generate CostHead report
                result = generate_costhead_report(
                    gin_mapped_path,
                    st.session_state.costhead_file,
                    temp_dir,
                    match_mode,
                    st.session_state.gst_file
                )

                # Create a zip file with all outputs
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    # Add all files from the output directory
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, temp_dir)
                            zip_file.write(file_path, arcname)

                zip_buffer.seek(0)

                # Display success message and download button
                st.success("✅ CostHead report generated successfully!")

                st.download_button(
                    label="📥 Download CostHead Report (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="CostHead_Report.zip",
                    mime="application/zip",
                    use_container_width=True
                )

                # Show summary
                st.markdown("### 📊 Report Summary")
                col1, col2, col3 = st.columns(3)

                with col1:
                    st.metric("GIN Mapped Rows", len(merged_df))

                with col2:
                    st.metric("Matched Rows", len(merged_df[merged_df["ActivityName"] != ""]))

                with col3:
                    st.metric("Unmatched Rows", len(merged_df[merged_df["ActivityName"] == ""]))

                # Show preview of merged data
                with st.expander("Preview Merged Data"):
                    st.dataframe(merged_df.head(20), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error generating report: {str(e)}")
        st.exception(e)

def clear_data():
    """Clear all uploaded data"""
    # Clear all session state data
    st.session_state.activities_df = None
    st.session_state.gin_df = None
    st.session_state.costhead_file = None
    st.session_state.gst_file = None

    # Clear file uploader states by using unique keys
    if 'clear_counter' not in st.session_state:
        st.session_state.clear_counter = 0
    st.session_state.clear_counter += 1

    # Show success message
    st.success("✅ All data cleared successfully!")

    # Force a complete page refresh
    st.rerun()

# Clear button moved above - removed from here

if __name__ == "__main__":
    main()
