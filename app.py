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
import plotly.graph_objects as go
import plotly.express as px

# Import the existing processing classes
from pdf_extractor import AttendancePDFExtractor, RosterExtractor
from report_generator import TrainOperationsRosterReportGenerator

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


def display_roster_report(df, department, conversion_type):
    """Display roster analysis report for Train Operations Roster"""
    if department != "Train Operations" or conversion_type != "trip_chart":
        return
    
    try:
        # Generate report
        report = TrainOperationsRosterReportGenerator(df)
        
        # Display report title with date range
        if report.date_range and len(report.date_range) > 0:
            date_text = f" ({report.date_range[0]} to {report.date_range[-1]})"
        else:
            date_text = ""
        st.markdown(f"## 📊 Roster Analysis Report{date_text}")
        
        # Create tabs for different report sections
        tab1, tab2, tab3, tab4 = st.tabs(["Summary Statistics", "Daily Trends", "Shift Analysis", "Employee Details"])
        
        # Tab 1: Summary Statistics
        with tab1:
            st.markdown("### 📋 Summary Statistics")
            summary_df = report.generate_summary_table()
            st.dataframe(summary_df, use_container_width=True)
            
            # Display as metrics
            cols = st.columns(len(summary_df))
            for idx, row in summary_df.iterrows():
                with cols[idx]:
                    st.metric(row['Metric'], row['Value'])
        
        # Tab 2: Daily Trends
        with tab2:
            st.markdown("### 📈 Daily Trends")
            daily_df = report.generate_daily_trends()
            st.dataframe(daily_df, use_container_width=True)
            
            # Create visualization
            if 'Total Shifts' in daily_df.columns:
                fig = px.line(daily_df, x='Day', y='Total Shifts', 
                             markers=True, title='Daily Shift Assignments Trend')
                st.plotly_chart(fig, use_container_width=True)
        
        # Tab 3: Shift Analysis
        with tab3:
            st.markdown("### 🎯 Shift Analysis")
            
            col1, col2 = st.columns(2)
            
            shift_analysis, leave_analysis = report.generate_shift_analysis()
            
            # Shift distribution chart
            with col1:
                st.markdown("**Shift Distribution**")
                shift_df = pd.DataFrame(list(shift_analysis.items()), 
                                       columns=['Shift Type', 'Count'])
                fig1 = px.pie(shift_df, values='Count', names='Shift Type',
                             title='Working Shifts by Type')
                st.plotly_chart(fig1, use_container_width=True)
                st.dataframe(shift_df, use_container_width=True, hide_index=True)
            
            # Leave distribution chart
            with col2:
                st.markdown("**Leave Distribution**")
                leave_df = pd.DataFrame(list(leave_analysis.items()), 
                                       columns=['Leave Type', 'Count'])
                fig2 = px.pie(leave_df, values='Count', names='Leave Type',
                             title='Leaves by Type')
                st.plotly_chart(fig2, use_container_width=True)
                st.dataframe(leave_df, use_container_width=True, hide_index=True)
        
        # Tab 4: Employee Details
        with tab4:
            st.markdown("### 👥 Employee Details Shift Analysis")
            
            employee_df = report.generate_employee_details()
            
            # Display full table
            st.dataframe(employee_df, use_container_width=True)
            
            # Summary statistics by shift type
            st.markdown("**Summary Statistics**")
            summary_cols = ['Working Shift RRTS', 'Working Shift MRTS', 'CL', 'SL', 'WL', 'CO']
            available_cols = [col for col in summary_cols if col in employee_df.columns]
            
            if available_cols:
                summary_stats = employee_df[available_cols].describe().T
                st.dataframe(summary_stats, use_container_width=True)
            
            # Top employees by shifts
            st.markdown("**Top 10 Employees by Total Assignments**")
            top_employees = employee_df.nlargest(10, 'Total')[
                ['Employee', 'Personnel_Number', 'Working Shift RRTS', 
                 'Working Shift MRTS', 'Total']
            ]
            st.dataframe(top_employees, use_container_width=True, hide_index=True)
    
    except Exception as e:
        st.error(f"Error generating report: {str(e)}")


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
        conversion_type = "attendance"
    elif department == "Train Operations" and train_ops_mode == "Roster":
        st.markdown("Upload your Train Operations Roster PDF file and download the converted Excel file.")
        file_type_hint = "Roster PDF"
        conversion_type = "trip_chart"
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
                
                # Generate and display report for Train Operations Roster
                if department == "Train Operations" and train_ops_mode == "Roster":
                    st.markdown("---")
                    df = extract_df_from_bytes(uploaded_file.getvalue(), department, conversion_type)
                    if df is not None and not df.empty:
                        display_roster_report(df, department, conversion_type)

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
