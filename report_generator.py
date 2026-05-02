"""
Report Generator for Train Operations Roster Analysis
Generates Summary Statistics, Daily Trends, Shift Analysis, and Employee Details

EXCLUSIVE FOR: Train Operations - Roster Selection
NOTE: For Stations and OCC departments, a different report class will be implemented
"""

import pandas as pd
import re
from collections import Counter, defaultdict
from datetime import datetime


class TrainOperationsRosterReportGenerator:
    """
    Generate analysis reports from extracted Train Operations Roster data
    EXCLUSIVE FOR: Train Operations - Roster Selection
    """
    
    # Shift categorization - based on shift code prefixes
    RRTS_SHIFTS = {'SR', 'WDR', 'RR', 'HSB A', 'HSB B'}  # RRTS duties
    MRTS_SHIFTS = {'SM', 'WDM', 'HSB M'}  # MRTS duties
    LEAVE_CODES = {'CL', 'SL', 'WL', 'CO', 'EL', 'AP', 'LM', 'WO'}  # All leave types
    
    def __init__(self, df):
        """
        Initialize with extracted Train Operations Roster dataframe
        
        Args:
            df: DataFrame from PDF extraction containing:
                - Employee, Personnel_Number, Scheduling_Row
                - Daily columns with shift codes (e.g., "RR-01", "SR-14", "SM-05")
                - Paid_Time or similar total column
        """
        self.df = df
        self.shift_cols = self._identify_shift_columns()
        self.day_cols = self._identify_day_columns()
        self.date_range = self._extract_date_range()
        
    def _identify_shift_columns(self):
        """Identify shift columns - Shift (date) columns containing the actual shift/leave codes"""
        return sorted([col for col in self.df.columns if col.startswith('Shift (')],
                     key=lambda x: self._extract_date_from_col(x))
    
    def _identify_day_columns(self):
        """Identify daily attendance columns - same as shift columns for this data format"""
        # For this roster format, the shift codes are directly in "Shift (date)" columns
        return self.shift_cols
    
    def _extract_date_from_col(self, col):
        """Extract date from column name for sorting"""
        match = re.search(r'\(([0-9.]+)\)', col)
        if match:
            date_str = match.group(1)
            # Convert "18.01" to sortable format
            parts = date_str.split('.')
            if len(parts) == 2:
                return int(parts[0]) * 100 + int(parts[1])
        return 0
    
    def _extract_date_range(self):
        """Extract date range from shift column headers"""
        dates = []
        for col in self.df.columns:
            # Match date pattern like "(18.01)" or "(Sun. 18.01)"
            match = re.search(r'\(([A-Za-z.]*\s*\d{1,2}\.\d{2})\)', col)
            if match:
                date_str = match.group(1).strip()
                dates.append(date_str)
        
        # Return unique dates in order
        return list(dict.fromkeys(dates)) if dates else []
    
    def _extract_shift_code(self, shift_val):
        """
        Extract shift code from value
        E.g., "RR-01", "SR-14", "SM-05", "WDM-13", "HSB A-1"
        """
        if not shift_val or str(shift_val).strip() == '':
            return None
        
        shift_str = str(shift_val).strip()
        # Remove newlines and extra spaces
        shift_str = shift_str.replace('\n', ' ').replace('\r', '').strip()
        
        # Extract shift code (letters and numbers before timing)
        match = re.match(r'([A-Za-z\s]+-?\d+)', shift_str)
        if match:
            code = match.group(1).strip()
            # Handle special case: "HSB A-1" -> "HSB A"
            if code.startswith('HSB'):
                parts = code.split('-')
                return f"{parts[0][:5]}{parts[0][5:].strip()}"
            # Extract prefix (RR, SR, SM, WDR, WDM, HSB, etc.)
            prefix_match = re.match(r'([A-Z]+)', code)
            if prefix_match:
                return prefix_match.group(1)
        return None
    
    def _is_leave(self, code):
        """Check if attendance code is a leave type"""
        if not code or str(code).strip() == '':
            return False
        code_str = str(code).strip().upper()
        
        # Check exact matches
        if code_str in self.LEAVE_CODES:
            return True
        
        # Check partial matches for codes like "LMCL" (Limited Medical CL)
        for leave_code in self.LEAVE_CODES:
            if leave_code in code_str:
                return True
        
        return False
    
    def _categorize_shift(self, shift_code):
        """Categorize shift into RRTS, MRTS, or Other"""
        if not shift_code:
            return 'Other Shifts'
        
        if shift_code in self.RRTS_SHIFTS:
            return 'Working Shift RRTS'
        elif shift_code in self.MRTS_SHIFTS:
            return 'Working Shift MRTS'
        else:
            return 'Other Shifts'
    
    def _extract_leave_type(self, code):
        """Extract the actual leave code"""
        if not code or str(code).strip() == '':
            return 'Unknown'
        
        code_str = str(code).strip().upper()
        
        # Exact match
        if code_str in self.LEAVE_CODES:
            return code_str
        
        # Partial match - return the matched code
        for leave_code in sorted(self.LEAVE_CODES, key=len, reverse=True):
            if leave_code in code_str:
                return leave_code
        
        return code_str[:2].upper() if len(code_str) >= 2 else code_str
    
    def generate_summary_statistics(self):
        """Generate overall summary statistics"""
        summary = {}
        
        # Employee count
        summary['Total Employees'] = len(self.df)
        
        # Reporting period
        if self.date_range:
            summary['Period'] = f"{self.date_range[0]} to {self.date_range[-1]}"
        else:
            summary['Period'] = f"{len(self.day_cols)} days"
        
        # Count duties and leaves
        total_shifts = 0
        total_leaves = 0
        blank = 0
        
        for idx, row in self.df.iterrows():
            for col in self.day_cols:
                if col in row.index:
                    val = str(row[col]).strip()
                    if not val or val == '':
                        blank += 1
                    elif self._is_leave(val):
                        total_leaves += 1
                    else:
                        total_shifts += 1
        
        summary['Total Shifts'] = total_shifts
        summary['Total Leaves'] = total_leaves
        summary['Total Records'] = total_shifts + total_leaves + blank
        
        return summary
    
    def generate_daily_trends(self):
        """Generate daily trends of duties and leaves"""
        daily_data = []
        
        for day_col in self.day_cols:
            # Extract date from column name like "Shift (18.01)"
            match = re.search(r'\(([^)]+)\)', day_col)
            day_label = match.group(1) if match else day_col
            
            leaves = defaultdict(int)
            shifts = defaultdict(int)
            blank = 0
            
            for idx, row in self.df.iterrows():
                val = str(row[day_col]).strip() if day_col in row.index else ''
                
                if not val or val == '':
                    blank += 1
                elif self._is_leave(val):
                    leave_type = self._extract_leave_type(val)
                    leaves[leave_type] += 1
                else:
                    # It's a shift code
                    shifts[val] += 1
            
            day_record = {'Day': day_label, 'Total Shifts': sum(shifts.values())}
            
            # Add leave type counts
            for leave_code in sorted(leaves.keys()):
                day_record[leave_code] = leaves[leave_code]
            
            day_record['Blank'] = blank
            daily_data.append(day_record)
        
        return pd.DataFrame(daily_data)
    
    def generate_shift_analysis(self):
        """Generate shift analysis by category"""
        shift_analysis = {
            'Working Shift RRTS': 0,
            'Working Shift MRTS': 0,
            'Other Shifts': 0
        }
        
        leave_analysis = {}
        
        # Count shifts and leaves
        for idx, row in self.df.iterrows():
            for col in self.day_cols:
                if col in row.index:
                    val = str(row[col]).strip()
                    
                    if self._is_leave(val):
                        leave_type = self._extract_leave_type(val)
                        leave_analysis[leave_type] = leave_analysis.get(leave_type, 0) + 1
                    elif val and val != '':
                        shift_code = self._extract_shift_code(val)
                        category = self._categorize_shift(shift_code)
                        shift_analysis[category] += 1
        
        return shift_analysis, leave_analysis
    
    def generate_employee_details(self):
        """Generate per-employee shift and leave analysis"""
        employee_details = []
        
        # Get all unique leave types
        all_leave_types = set()
        for idx, row in self.df.iterrows():
            for col in self.day_cols:
                val = str(row[col]).strip() if col in row.index else ''
                if self._is_leave(val):
                    leave_type = self._extract_leave_type(val)
                    all_leave_types.add(leave_type)
        
        for idx, row in self.df.iterrows():
            emp_record = {
                'Employee': row.get('Employee', ''),
                'Personnel_Number': row.get('Personnel_Number', ''),
                'Working Shift RRTS': 0,
                'Working Shift MRTS': 0,
                'Other Shifts': 0
            }
            
            # Add leave type columns
            for leave_code in sorted(all_leave_types):
                emp_record[leave_code] = 0
            
            # Count for this employee
            for col in self.day_cols:
                if col in row.index:
                    val = str(row[col]).strip()
                    
                    if self._is_leave(val):
                        leave_type = self._extract_leave_type(val)
                        if leave_type in emp_record:
                            emp_record[leave_type] += 1
                    elif val and val != '':
                        shift_code = self._extract_shift_code(val)
                        category = self._categorize_shift(shift_code)
                        
                        if category == 'Working Shift RRTS':
                            emp_record['Working Shift RRTS'] += 1
                        elif category == 'Working Shift MRTS':
                            emp_record['Working Shift MRTS'] += 1
                        else:
                            emp_record['Other Shifts'] += 1
            
            # Calculate total
            total = (
                emp_record['Working Shift RRTS'] + 
                emp_record['Working Shift MRTS'] + 
                emp_record['Other Shifts'] +
                sum(emp_record.get(code, 0) for code in all_leave_types)
            )
            emp_record['Total'] = total
            
            employee_details.append(emp_record)
        
        return pd.DataFrame(employee_details)
    
    def generate_summary_table(self):
        """Generate summary statistics table"""
        summary = self.generate_summary_statistics()
        return pd.DataFrame(list(summary.items()), columns=['Metric', 'Value'])


# Placeholder classes for future implementation
class StationsRosterReportGenerator:
    """
    Report generator for Stations department roster analysis
    TODO: Implement Stations-specific report logic
    """
    def __init__(self, df):
        self.df = df
        raise NotImplementedError("Stations report generator not yet implemented")


class OCCRosterReportGenerator:
    """
    Report generator for OCC (Operations Control Center) roster analysis
    TODO: Implement OCC-specific report logic
    """
    def __init__(self, df):
        self.df = df
        raise NotImplementedError("OCC report generator not yet implemented")


if __name__ == "__main__":
    print("Report Generator module loaded successfully")
    print("Generators available:")
    print("  - TrainOperationsRosterReportGenerator (Train Operations - Roster)")
    print("  - StationsRosterReportGenerator (Coming Soon)")
    print("  - OCCRosterReportGenerator (Coming Soon)")
