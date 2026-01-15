"""
Online Attendance Analysis Tool
Web-based application for PDF to Excel conversion
"""

import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import sys
import io
import tempfile
import os

# Import the existing processing classes
from pdf_extractor import AttendancePDFExtractor, RosterExtractor

# Page configuration
st.set_page_config(
    page_title="Attendance Analysis Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 1rem;
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
    }
    .stAlert {
        border-radius: 10px;
    }
    </style>
""", unsafe_allow_html=True)


def process_pdf_to_excel(uploaded_file, department='Stations', pdf_type='attendance'):
    """
    Process uploaded PDF file and return Excel data
    
    Args:
        uploaded_file: The uploaded PDF file
        department: 'Stations', 'OCC', or 'Train Operations'
        pdf_type: Type of PDF - 'attendance' for Stations/OCC/Roster, 'trip_chart' for Train Operations Trip Chart
    """
    try:
        # Save uploaded file to temporary location
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = tmp_file.name

        # Choose extractor based on department and pdf_type
        if pdf_type == 'trip_chart':
            # Trip Chart: use AttendancePDFExtractor (with Shift/Shift_Timings)
            extractor = AttendancePDFExtractor(tmp_path)
            extractor.extract_tables()
            df = extractor.create_dataframe()
            sheet_name = 'Trip Chart Data'
        else:
            # Stations, OCC, Train Ops Roster: use RosterExtractor (simple extraction)
            extractor = RosterExtractor(tmp_path)
            extractor.extract_tables()
            df = extractor.create_dataframe()
            sheet_name = 'Attendance Data'

        if df is None or df.empty:
            return None, "No data found in PDF. Please check the file format."

        # Create Excel file in memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Clean up temporary file
        os.unlink(tmp_path)

        return output.getvalue(), None

    except Exception as e:
        return None, f"Error processing PDF: {str(e)}"


@st.cache_data(show_spinner=False)
def extract_df_from_bytes(file_bytes: bytes, department='Stations', pdf_type: str = 'attendance'):
    """Extract DataFrame from uploaded PDF bytes. Cached to make preview responsive."""
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(file_bytes)
            tmp_path = tmp_file.name

        if pdf_type == 'trip_chart':
            # Trip Chart: use AttendancePDFExtractor (with Shift/Shift_Timings)
            extractor = AttendancePDFExtractor(tmp_path)
            extractor.extract_tables()
            df = extractor.create_dataframe()
        else:
            # Stations, OCC, Train Ops Roster: use RosterExtractor
            extractor = RosterExtractor(tmp_path)
            extractor.extract_tables()
            df = extractor.create_dataframe()

        try:
            os.unlink(tmp_path)
        except Exception:
            pass

        return df
    except Exception as e:
        return None


def main():
    """Main application function"""

    # Header
    st.markdown('<div class="main-header">📊 Attendance Analysis Tool</div>', unsafe_allow_html=True)

    # Sidebar
    st.sidebar.title("⚙️ Tool Controls")

    # Department Selection
    st.sidebar.header("1️⃣ Department")
    department = st.sidebar.selectbox(
        "Choose Department",
        ["Stations", "OCC", "Train Operations"],
        help="Select the department"
    )

    # Train Operations specific option
    train_ops_mode = None
    if department == "Train Operations":
        st.sidebar.header("1b️⃣ Train Operations Type")
        train_ops_mode = st.sidebar.radio(
            "Select conversion type",
            ["Trip Chart", "Roster"],
            help="Choose what type of PDF to convert"
        )

    # Time Period (for future use)
    st.sidebar.header("2️⃣ Time Period")
    st.sidebar.info("⏰ Coming Soon - Time period selection")

    # Analysis Options (for future use)
    st.sidebar.header("3️⃣ Analysis Options")
    st.sidebar.info("📊 Coming Soon - Analysis features")

    # Data Management
    st.sidebar.header("4️⃣ Data Management")
    st.sidebar.info("🔄 Coming Soon - Batch processing")

    # Main Content - PDF to Excel Converter
    st.markdown("## 📄 PDF to Excel Converter")
    
    # Determine conversion type based on department and mode
    if department == "Train Operations" and train_ops_mode == "Trip Chart":
        st.markdown("Upload your Trip Chart PDF file and download the converted Excel file.")
        file_type_hint = "Trip Chart PDF"
        conversion_type = "trip_chart"
    elif department == "Train Operations" and train_ops_mode == "Roster":
        st.markdown("Upload your Train Operations Roster PDF file and download the converted Excel file.")
        file_type_hint = "Roster PDF"
        conversion_type = "attendance"
    else:
        st.markdown("Upload your attendance PDF file and download the converted Excel file.")
        file_type_hint = "Attendance PDF"
        conversion_type = "attendance"

    # File upload
    uploaded_file = st.file_uploader(
        "Choose a PDF file",
        type=['pdf'],
        help=f"Upload a {file_type_hint} file"
    )

    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")

        # Preview (optional)
        if st.checkbox("👀 Preview Data"):
            with st.spinner("Extracting preview..."):
                df = extract_df_from_bytes(uploaded_file.getvalue(), department, conversion_type)

            if df is not None and not df.empty:
                st.markdown("### 📋 Data Preview")
                st.dataframe(df.head(20), use_container_width=True)
                st.info(f"📊 Total rows: {len(df)} | Total columns: {len(df.columns)}")
            else:
                st.warning("No data extracted for preview. Please check the PDF format.")

        # Process button
        if st.button("🔄 Convert to Excel", type="primary"):
            with st.spinner("Processing PDF... This may take a few moments."):
                excel_data, error = process_pdf_to_excel(uploaded_file, department, conversion_type)

            if error:
                st.error(f"❌ {error}")
            else:
                st.success("✅ PDF converted successfully!")

                # Download button
                if department == "Train Operations" and train_ops_mode == "Trip Chart":
                    file_name = f"{Path(uploaded_file.name).stem}_trip_chart.xlsx"
                elif department == "Train Operations" and train_ops_mode == "Roster":
                    file_name = f"{Path(uploaded_file.name).stem}_roster.xlsx"
                else:
                    file_name = f"{Path(uploaded_file.name).stem}_converted.xlsx"

                st.download_button(
                    label="📊 Download Excel File",
                    data=excel_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Download the converted Excel file"
                )

    # Instructions
    with st.expander("📖 How to Use"):
        st.markdown("""
        1. **Select Department**: Choose from Stations, OCC, or Train Operations
        2. **Upload PDF**: Select your attendance or roster PDF file
        3. **Preview** (Optional): Check the data preview to verify the extraction
        4. **Convert**: Click "Convert to Excel" button to process the PDF
        5. **Download**: Download the converted Excel file

        **Supported Formats**: PDF files containing attendance/roster tables
        **File Sizes**: Up to 200MB per file
        """)

    # Footer
    st.markdown("---")
    st.markdown(
        f"**Last Updated:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | "
        "**Version:** 1.0 | "
        "**Status:** Beta"
    )


if __name__ == "__main__":
    main()
