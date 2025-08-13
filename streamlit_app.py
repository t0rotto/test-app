import streamlit as st
import os
import pandas as pd
from datetime import datetime
import pytz
import zipfile
import warnings
import logging
import tempfile
import shutil
from io import BytesIO
import base64

# Configure page
st.set_page_config(
    page_title="Dispatch Summary Analyzer",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Import our processing modules
from data_processor import DataProcessor
from excel_generator import ExcelGenerator
from utils import setup_logging, METRICS, ANALYSIS_VALUES

# === MAIN APP ===
def main():
    st.title("📦 Dispatch Summary Analyzer")
    st.markdown("""
    **Transform your logistics data into actionable insights!**
    
    Upload ZIP files containing dispatch summaries and get comprehensive analysis with:
    - 📊 Automated data extraction and processing
    - 📈 Multi-sheet Excel reports with pivot tables
    - 🕐 Time-based MDT analysis with charts
    - 💰 Cost impact calculations
    """)
    
    # Sidebar configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # File upload
        uploaded_files = st.file_uploader(
            "Upload ZIP files containing dispatch data",
            type=['zip'],
            accept_multiple_files=True,
            help="Select one or more ZIP files containing Excel/CSV dispatch summaries"
        )
        
        # Advanced options
        with st.expander("🔧 Advanced Options"):
            exclude_backhauls = st.checkbox(
                "Exclude Unplanned Backhauls", 
                value=True,
                help="Skip files containing 'Unplanned_Backhauls_Reason'"
            )
            
            annualization_weeks = st.number_input(
                "Custom weeks for annualization", 
                min_value=0, 
                max_value=52, 
                value=0,
                help="Leave 0 for auto-detection based on data"
            )
            
            cost_per_stop = st.number_input("Cost per Stop ($)", value=30.0, format="%.2f")
            cost_per_route = st.number_input("Cost per Route ($)", value=250.0, format="%.2f")
            cost_per_mile = st.number_input("Cost per Mile ($)", value=2.5, format="%.2f")
    
    # Main processing area
    if uploaded_files:
        st.header("🚀 Processing Pipeline")
        
        # Create processing button
        if st.button("🔄 Process Data", type="primary", use_container_width=True):
            process_uploaded_files(
                uploaded_files, 
                exclude_backhauls,
                annualization_weeks,
                (cost_per_stop, cost_per_route, cost_per_mile)
            )
    else:
        st.info("👆 Please upload ZIP files to begin processing")
        
        # Show sample data format
        with st.expander("📋 Expected Data Format"):
            st.markdown("""
            **Your ZIP files should contain Excel files with:**
            - Dispatch summary data with columns like Activity Start Time, Route Number, etc.
            - Files named with pattern: `Date_Report_StoreID_Country_Product_ID_Baseline.xlsx`
            - Data starting around row 12 (after headers)
            
            **The app will automatically:**
            - Extract and process all Excel/CSV files from ZIPs
            - Generate analysis sheets with pivot tables
            - Create time-based MDT analysis
            - Calculate cost impacts and routing efficiency
            """)

def process_uploaded_files(uploaded_files, exclude_backhauls, custom_weeks, cost_params):
    """Process the uploaded files and generate analysis"""
    try:
        # Setup logging for this session
        setup_logging()
        
        # Create temporary directories
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = os.path.join(temp_dir, "inputs")
            extracted_path = os.path.join(temp_dir, "extracted")
            os.makedirs(input_path, exist_ok=True)
            os.makedirs(extracted_path, exist_ok=True)
            
            # Progress tracking
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Step 1: Save uploaded files
            status_text.text("💾 Saving uploaded files...")
            for i, uploaded_file in enumerate(uploaded_files):
                file_path = os.path.join(input_path, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
            progress_bar.progress(20)
            
            # Step 2: Extract files
            status_text.text("📂 Extracting ZIP contents...")
            processor = DataProcessor(exclude_backhauls)
            processor.extract_files(input_path, extracted_path)
            progress_bar.progress(40)
            
            # Step 3: Process data
            status_text.text("🔄 Processing dispatch data...")
            df1, df2, df3 = processor.process_extracted_files(extracted_path)
            progress_bar.progress(70)
            
            # Step 4: Generate Excel report
            status_text.text("📊 Generating Excel report...")
            excel_generator = ExcelGenerator(cost_params, custom_weeks)
            output_buffer = excel_generator.create_report(df1, df2, df3)
            progress_bar.progress(100)
            
            status_text.text("✅ Processing complete!")
            
            # Display results
            show_results(df1, df2, df3, output_buffer)
            
    except Exception as e:
        st.error(f"❌ Error during processing: {str(e)}")
        logging.error(f"Processing error: {e}")

def show_results(df1, df2, df3, output_buffer):
    """Display processing results and download option"""
    st.success("🎉 Data processing completed successfully!")
    
    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📄 Total Records", len(df1))
    with col2:
        st.metric("🚛 Routes Processed", df1['Routes'].sum() if 'Routes' in df1.columns else 0)
    with col3:
        st.metric("📦 Total Pallets", df1['Pallets'].sum() if 'Pallets' in df1.columns else 0)
    with col4:
        st.metric("📊 Data Sources", df1['DC'].nunique() if 'DC' in df1.columns else 0)
    
    # Data preview tabs
    tab1, tab2, tab3 = st.tabs(["📋 Totals Summary", "🚛 MDT Data", "📦 Dispatch Details"])
    
    with tab1:
        if not df1.empty:
            st.subheader("Totals Summary Data")
            st.dataframe(df1.head(10), use_container_width=True)
            st.caption(f"Showing first 10 of {len(df1)} total records")
        else:
            st.warning("No totals data available")
    
    with tab2:
        if not df2.empty:
            st.subheader("MDT Analysis Data")
            st.dataframe(df2.head(10), use_container_width=True)
            st.caption(f"Showing first 10 of {len(df2)} MDT records")
            
            # Quick time distribution chart
            if 'Time Range' in df2.columns and 'Simulation' in df2.columns:
                time_dist = df2.groupby(['Time Range', 'Simulation']).size().unstack(fill_value=0)
                if not time_dist.empty:
                    st.subheader("🕐 Trailer Distribution by Time")
                    st.bar_chart(time_dist)
        else:
            st.warning("No MDT data available")
    
    with tab3:
        if not df3.empty:
            st.subheader("Dispatch Summary Data")
            st.dataframe(df3.head(10), use_container_width=True)
            st.caption(f"Showing first 10 of {len(df3)} dispatch records")
        else:
            st.warning("No dispatch data available")
    
    # Download section
    st.subheader("📥 Download Results")
    
    # Generate timestamp for filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"dispatch_analysis_{timestamp}.xlsx"
    
    st.download_button(
        label="📊 Download Excel Report",
        data=output_buffer.getvalue(),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.info("""
    📋 **Your Excel report contains:**
    - **Analysis**: Summary metrics with cost calculations
    - **Summary**: Pivot table analysis by DC and simulation
    - **Dispatch Summaries Raw**: Complete dispatch data
    - **MDT Raw**: Mobile Data Terminal records
    - **Totals Raw**: Aggregated totals data
    - **MDT Analysis**: Time-based analysis with charts
    """)

if __name__ == "__main__":
    main()
