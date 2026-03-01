"""
Excel Consolidator for Attendance Data
Extracts PDF data and consolidates into Excel with year-wise sheets
"""

import pandas as pd
from pathlib import Path
import re
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pdf_extractor import AttendancePDFExtractor
import warnings
warnings.filterwarnings('ignore')


class ExcelConsolidator:
    """Consolidate PDF attendance data into Excel files with year-wise sheets"""
    
    def __init__(self, department_folder):
        """
        Initialize consolidator for a department
        
        Args:
            department_folder: Path to department folder (Station/OCC/Train Operations)
        """
        self.department_folder = Path(department_folder)
        self.excel_path = self.department_folder / "reports.xlsx"
        self.pdf_files = list(self.department_folder.glob("*.pdf"))
        
    def parse_filename(self, pdf_path):
        """
        Extract month and year from various filename formats:
        - APR_2025_SO-SC.pdf
        - NOV_2025_BCC_ASC.pdf
        - Nov BCC-DC.pdf
        - november_2025_report.pdf
        """
        filename = Path(pdf_path).stem
        
        # Convert month names to number (both full and abbreviated)
        months = {
            'JAN': '01', 'JANUARY': '01',
            'FEB': '02', 'FEBRUARY': '02',
            'MAR': '03', 'MARCH': '03',
            'APR': '04', 'APRIL': '04',
            'MAY': '05',
            'JUN': '06', 'JUNE': '06',
            'JUL': '07', 'JULY': '07',
            'AUG': '08', 'AUGUST': '08',
            'SEP': '09', 'SEPT': '09', 'SEPTEMBER': '09',
            'OCT': '10', 'OCTOBER': '10',
            'NOV': '11', 'NOVEMBER': '11',
            'DEC': '12', 'DECEMBER': '12'
        }
        
        # Pattern 1: MONTH_YEAR or MONTH_YEAR_something (case insensitive)
        # Matches: APR_2025, NOV_2025_BCC, november_2025, etc.
        match = re.search(r'([A-Za-z]+)[_\s-](\d{4})', filename, re.IGNORECASE)
        if match:
            month_str = match.group(1).upper()
            year = match.group(2)
            
            # Try exact match first, then prefix match for abbreviated months
            month_num = months.get(month_str)
            if not month_num:
                # Try to find matching month by prefix (e.g., NOV matches NOVEMBER)
                for month_key, month_val in months.items():
                    if month_key.startswith(month_str) or month_str.startswith(month_key):
                        month_num = month_val
                        month_str = month_key
                        break
            
            if month_num:
                month_name = month_str.capitalize() if len(month_str) <= 3 else month_str.title()
                return {
                    'year': year,
                    'month': month_name,
                    'month_num': month_num,
                    'sort_key': f"{year}_{month_num}"
                }
        
        # Pattern 2: Month at start without year (e.g., "Nov BCC-DC.pdf")
        # Default to current year 2025
        match = re.search(r'^([A-Za-z]+)[_\s-]', filename, re.IGNORECASE)
        if match:
            month_str = match.group(1).upper()
            year = '2025'  # Default year
            
            # Try exact match or prefix match
            month_num = months.get(month_str)
            if not month_num:
                for month_key, month_val in months.items():
                    if month_key.startswith(month_str) or month_str.startswith(month_key):
                        month_num = month_val
                        month_str = month_key
                        break
            
            if month_num:
                month_name = month_str.capitalize() if len(month_str) <= 3 else month_str.title()
                return {
                    'year': year,
                    'month': month_name,
                    'month_num': month_num,
                    'sort_key': f"{year}_{month_num}"
                }
        
        return None
    
    def extract_pdf_to_dataframe(self, pdf_path):
        """Extract data from PDF and add month/year columns"""
        extractor = AttendancePDFExtractor(str(pdf_path))
        df = extractor.create_dataframe()
        
        if df.empty:
            return None
        
        # Parse filename for month/year
        date_info = self.parse_filename(pdf_path)
        if date_info:
            df.insert(0, 'Year', date_info['year'])
            df.insert(1, 'Month', date_info['month'])
            df.insert(2, 'Month_Num', date_info['month_num'])
            df.insert(3, 'Sort_Key', date_info['sort_key'])
        
        return df
    
    def consolidate_year_data(self, year):
        """Consolidate all data for a specific year"""
        year_data = []
        
        for pdf_path in self.pdf_files:
            date_info = self.parse_filename(pdf_path)
            
            if date_info and date_info['year'] == year:
                print(f"  Processing: {pdf_path.name}")
                df = self.extract_pdf_to_dataframe(pdf_path)
                
                if df is not None and not df.empty:
                    year_data.append(df)
                    print(f"    ✓ Extracted {len(df)} records")
        
        if not year_data:
            return None
        
        # Combine all months for this year
        combined_df = pd.concat(year_data, ignore_index=True)
        
        # Sort by month
        combined_df = combined_df.sort_values('Sort_Key')
        
        # Remove sort key column (internal use only)
        combined_df = combined_df.drop('Sort_Key', axis=1)
        
        return combined_df
    
    def update_excel(self, overwrite_existing=True):
        """
        Update Excel file with data from all PDFs
        
        Args:
            overwrite_existing: If True, overwrite existing month data; if False, skip
        """
        print(f"\n📊 Consolidating data for: {self.department_folder.name}")
        print(f"Found {len(self.pdf_files)} PDF files")
        
        # Get all unique years from PDF files
        years = set()
        for pdf_path in self.pdf_files:
            date_info = self.parse_filename(pdf_path)
            if date_info:
                years.add(date_info['year'])
        
        if not years:
            print("❌ No valid PDF files found")
            return False
        
        print(f"Years found: {sorted(years)}")
        
        # Create or load Excel file
        excel_data = {}
        
        if self.excel_path.exists():
            print(f"📖 Loading existing Excel: {self.excel_path.name}")
            try:
                # Load all existing sheets
                xl_file = pd.ExcelFile(self.excel_path)
                for sheet_name in xl_file.sheet_names:
                    excel_data[sheet_name] = pd.read_excel(xl_file, sheet_name=sheet_name)
            except Exception as e:
                print(f"⚠️  Could not load existing Excel: {e}")
        
        # Process each year
        for year in sorted(years):
            print(f"\n📅 Processing Year: {year}")
            year_df = self.consolidate_year_data(year)
            
            if year_df is not None:
                excel_data[year] = year_df
                print(f"  ✓ Consolidated {len(year_df)} total records for {year}")
        
        # Write to Excel
        print(f"\n💾 Writing to Excel: {self.excel_path}")
        
        with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='w') as writer:
            for year in sorted(excel_data.keys()):
                df = excel_data[year]
                df.to_excel(writer, sheet_name=str(year), index=False)
                print(f"  ✓ Sheet '{year}': {len(df)} records")
        
        print(f"✅ Excel consolidation complete: {self.excel_path}")
        return True
    
    def get_consolidated_data(self, year=None):
        """
        Read consolidated data from Excel
        
        Args:
            year: Specific year to read (None = all years)
        
        Returns:
            DataFrame or dict of DataFrames
        """
        if not self.excel_path.exists():
            print(f"❌ Excel file not found: {self.excel_path}")
            return None
        
        try:
            if year:
                # Read specific year
                df = pd.read_excel(self.excel_path, sheet_name=str(year))
                return df
            else:
                # Read all years
                xl_file = pd.ExcelFile(self.excel_path)
                data = {}
                for sheet_name in xl_file.sheet_names:
                    data[sheet_name] = pd.read_excel(xl_file, sheet_name=sheet_name)
                return data
        except Exception as e:
            print(f"❌ Error reading Excel: {e}")
            return None
    
    def get_available_years(self):
        """Get list of years available in Excel"""
        if not self.excel_path.exists():
            return []
        
        try:
            xl_file = pd.ExcelFile(self.excel_path)
            return sorted(xl_file.sheet_names)
        except:
            return []
    
    def get_available_months(self, year):
        """Get list of months available for a specific year"""
        df = self.get_consolidated_data(year)
        if df is None or df.empty:
            return []
        
        if 'Month' in df.columns:
            months = df['Month'].unique()
            # Sort by month number if available
            if 'Month_Num' in df.columns:
                month_order = df[['Month', 'Month_Num']].drop_duplicates()
                month_order = month_order.sort_values('Month_Num')
                return month_order['Month'].tolist()
            return sorted(months)
        
        return []


def consolidate_all_departments(base_folder="Report IVU"):
    """Consolidate data for all departments"""
    base_path = Path(base_folder)
    departments = ['Stations', 'OCC', 'Train Operations']
    
    results = {}
    
    for dept in departments:
        dept_path = base_path / dept
        if dept_path.exists():
            print(f"\n{'='*80}")
            print(f"  {dept.upper()}")
            print(f"{'='*80}")
            
            consolidator = ExcelConsolidator(dept_path)
            success = consolidator.update_excel()
            results[dept] = success
        else:
            print(f"\n⚠️  Department folder not found: {dept}")
            results[dept] = False
    
    return results


if __name__ == "__main__":
    # Test consolidation
    results = consolidate_all_departments()
    
    print("\n" + "="*80)
    print("  CONSOLIDATION SUMMARY")
    print("="*80)
    
    for dept, success in results.items():
        status = "✅ Success" if success else "❌ Failed"
        print(f"{dept}: {status}")
