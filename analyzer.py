"""
Attendance Data Analyzer
Performs comprehensive analysis on attendance data including statistics, trends, and patterns
"""

import pandas as pd
import numpy as np
from datetime import datetime
from collections import Counter
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')


def load_employee_designations():
    """Load employee designations from the combined sheet"""
    try:
        emp_details_path = Path("Report IVU") / "Employee details" / "employee_details.xlsx"
        if emp_details_path.exists():
            df = pd.read_excel(emp_details_path, sheet_name='Combined')
            # Create a dictionary mapping Emp Id to Designation
            designation_dict = dict(zip(df['Emp Id'].astype(str), df['Designation']))
            return designation_dict
        return {}
    except Exception as e:
        print(f"Warning: Could not load employee designations: {e}")
        return {}


def load_employee_details(department=None):
    """Load complete employee details including Station/Division/Crew Control and AM"""
    try:
        base_dir = Path(__file__).parent.resolve()
        
        if department == "Stations" or not department:
            emp_details_path = base_dir / "Employee details" / "station_employee_details.xlsx"
            if emp_details_path.exists():
                try:
                    df_sc = pd.read_excel(emp_details_path, sheet_name='SC')
                    df_efo = pd.read_excel(emp_details_path, sheet_name='EFO')
                    df = pd.concat([df_sc, df_efo], ignore_index=True)
                except Exception as e:
                    # Fallback if sheets don't exist
                    df = pd.read_excel(emp_details_path, sheet_name='Combined')
                # Clean up Emp ID for robust matching
                df['clean_id'] = df['Emp Id'].astype(str).str.split('.').str[0].str.strip()
                
                details = {
                    'designation': dict(zip(df['clean_id'], df['Designation'].fillna(''))),
                    'station': dict(zip(df['clean_id'], df['Station'].fillna(''))),
                    'am': dict(zip(df['clean_id'], df['AM'].fillna('')))
                }
                return details
                
        elif department == "OCC":
            emp_details_path = base_dir / "Employee details" / "occ_emp_details.xlsx"
            if emp_details_path.exists():
                df = pd.read_excel(emp_details_path, sheet_name=0)
                df['clean_id'] = df['Emp Id'].astype(str).str.split('.').str[0].str.strip()
                
                details = {
                    'designation': dict(zip(df['clean_id'], df['Designation'].fillna(''))),
                    'station': dict(zip(df['clean_id'], df['Division'].fillna(''))) if 'Division' in df.columns else {},
                    'am': dict(zip(df['clean_id'], df['AM'].fillna(''))) if 'AM' in df.columns else {}
                }
                return details
                
        elif department == "Train Operations":
            emp_details_path = base_dir / "Employee details" / "to_ta_emp_details.xlsx"
            if emp_details_path.exists():
                try:
                    df_to = pd.read_excel(emp_details_path, sheet_name='TO')
                    df_ta = pd.read_excel(emp_details_path, sheet_name='TA')
                    df = pd.concat([df_to, df_ta], ignore_index=True)
                except Exception as e:
                    df = pd.read_excel(emp_details_path, sheet_name=0)
                
                df['clean_id'] = df['Emp Id'].astype(str).str.split('.').str[0].str.strip()
                
                details = {
                    'designation': dict(zip(df['clean_id'], df['Designation'].fillna(''))) if 'Designation' in df.columns else {},
                    'to_ta': dict(zip(df['clean_id'], df['crew control'].fillna(''))) if 'crew control' in df.columns else {},
                    'am': dict(zip(df['clean_id'], df['am'].fillna(''))) if 'am' in df.columns else {}
                }
                return details
                
        return None
    except Exception as e:
        print(f"Warning: Could not load employee details: {e}")
        return None


class AttendanceAnalyzer:
    """Analyze attendance data and generate insights"""
    
    # Shift codes (Working/Present days)
    SHIFT_CODES = {
        'M': 'Morning Shift (07:00-15:00)',
        'E': 'Evening Shift (15:00-22:00)',
        'N': 'Night Shift (22:00-07:00)',
        'G': 'General Shift (09:00-17:00)'
    }
    
    # Leave codes (On Leave - excluded from calculation)
    LEAVE_CODES = {
        'SL': 'Sick Leave',
        'CL': 'Casual Leave',
        'EL': 'Earned Leave',
        'OH': 'Off/Holiday',
        'CO': 'Compensatory Off',
        'PH': 'Public Holiday'
    }
    
    # Other codes
    OTHER_CODES = {
        'WO': 'Weekly Off',
        'AB': 'Absent (Unauthorized)'
    }
    
    # Combined for reference
    ATTENDANCE_CODES = {**SHIFT_CODES, **LEAVE_CODES, **OTHER_CODES}
    
    def __init__(self, df, month_name=""):
        self.df = df.copy()
        self.month_name = month_name
        
        # Clean all attendance codes - remove newlines
        self.day_columns = [col for col in df.columns if col.startswith('Day_')]
        for col in self.day_columns:
            self.df[col] = self.df[col].apply(
                lambda x: str(x).replace('\n', '').replace('\r', '').strip() if x and not pd.isna(x) else x
            )
        
        self.total_days = len(self.day_columns)
        
        # Calculate actual total days based on months in the data
        self.actual_total_days = self._calculate_actual_total_days()
    
    def _calculate_actual_total_days(self):
        """Calculate total days based on Year, Month, and Month_Num columns in the data"""
        from calendar import monthrange
        
        if 'Year' not in self.df.columns or 'Month_Num' not in self.df.columns:
            # Fallback: count non-empty columns per row
            max_days = 0
            for _, row in self.df.iterrows():
                non_empty = sum(1 for col in self.day_columns if row[col] and not pd.isna(row[col]) and str(row[col]).strip())
                max_days = max(max_days, non_empty)
            return max_days if max_days > 0 else self.total_days
        
        # Get unique year-month combinations
        unique_months = self.df[['Year', 'Month_Num']].drop_duplicates()
        
        total_days = 0
        for _, row in unique_months.iterrows():
            try:
                year = int(row['Year'])
                month = int(row['Month_Num'])
                # Get number of days in this month
                days_in_month = monthrange(year, month)[1]
                total_days += days_in_month
            except:
                pass
        
        return total_days if total_days > 0 else self.total_days
    
    def parse_attendance_code(self, code_str):
        """
        Parse attendance codes from the cell.
        Format examples:
        - 'M' or 'E' or 'N' or 'G' = Shift code (working day)
        - 'M-NASH' = Morning shift at NASH station
        - 'N-RITH22:00-07:00' = Night shift at RITH station with timing
        - 'SL', 'CL', 'EL', 'OH', 'CO', 'PH' = Leave codes
        - 'WO' = Weekly Off
        - 'AB' = Absent
        
        Returns dict with: shift, station, is_leave, is_absent, is_weekly_off, raw_code
        """
        if not code_str or pd.isna(code_str):
            return {'shift': None, 'station': None, 'is_leave': False, 
                    'is_absent': False, 'is_weekly_off': False, 'raw_code': ''}
        
        # Strip ALL whitespace from beginning and end
        code_str = str(code_str).strip().upper()
        if not code_str:
            return {'shift': None, 'station': None, 'is_leave': False, 
                    'is_absent': False, 'is_weekly_off': False, 'raw_code': ''}
        
        # Check for weekly off - be more specific
        if code_str == 'WO' or code_str.startswith('WO-') or code_str.startswith('WO '):
            return {'shift': None, 'station': None, 'is_leave': False,
                    'is_absent': False, 'is_weekly_off': True, 'raw_code': code_str}
        
        # Check for absent
        if code_str == 'AB' or code_str.startswith('AB-') or code_str.startswith('AB '):
            return {'shift': None, 'station': None, 'is_leave': False,
                    'is_absent': True, 'is_weekly_off': False, 'raw_code': code_str}
        
        # Check for leave codes (check exact match first, then startswith)
        for leave_code in self.LEAVE_CODES.keys():
            if code_str == leave_code or (len(code_str) >= len(leave_code) and code_str[:len(leave_code)] == leave_code):
                return {'shift': None, 'station': None, 'is_leave': True,
                        'is_absent': False, 'is_weekly_off': False, 'raw_code': code_str}
        
        # Check for shift codes
        first_char = code_str[0]
        if first_char in self.SHIFT_CODES:
            # Extract station name (everything after first hyphen, before timing)
            station = None
            if '-' in code_str:
                parts = code_str.split('-', 1)
                if len(parts) > 1:
                    # Remove timing if present (contains ':')
                    station_part = parts[1]
                    if ':' in station_part:
                        # Extract station name before timing
                        station = station_part.split(':')[0].replace('00', '').replace('-', '')
                    else:
                        station = station_part
            
            return {'shift': first_char, 'station': station, 'is_leave': False,
                    'is_absent': False, 'is_weekly_off': False, 'raw_code': code_str}
        
        # Unknown code - treat as empty
        return {'shift': None, 'station': None, 'is_leave': False,
                'is_absent': False, 'is_weekly_off': False, 'raw_code': code_str}
    
    def is_working_day(self, code_str):
        """Check if a day is a working day (has shift code: M, E, N, G)"""
        parsed = self.parse_attendance_code(code_str)
        return parsed['shift'] is not None
    
    def is_absent(self, code_str):
        """Check if employee was absent (AB)"""
        parsed = self.parse_attendance_code(code_str)
        return parsed['is_absent']
    
    def is_on_leave(self, code_str):
        """Check if employee was on leave (SL, CL, EL, OH, CO, PH)"""
        parsed = self.parse_attendance_code(code_str)
        return parsed['is_leave']
    
    def is_weekly_off(self, code_str):
        """Check if it's weekly off (WO)"""
        parsed = self.parse_attendance_code(code_str)
        return parsed['is_weekly_off']
    
    def calculate_employee_stats(self, row):
        """
        Calculate statistics for a single employee
        
        Attendance % = Present Days / (Total Days - Weekly Offs - Leaves) × 100
        
        Where:
        - Present Days = Days with shift codes (M, E, N, G)
        - Leave Days = SL, CL, EL, OH, CO, PH (excluded from denominator)
        - Weekly Off = WO (excluded from denominator)
        - Absent = AB (included in denominator, but counts as 0 in numerator)
        """
        stats = {
            'Employee': row['Employee'],
            'Personnel_Number': row['Personnel_Number'],
            'Total_Days': 0,
            'Present_Days': 0,
            'Absent_Days': 0,
            'Leave_Days': 0,
            'Weekly_Off': 0,
            'Attendance_Rate': 0.0,
            'Shift_Distribution': Counter(),
            'Stations': set()
        }
        
        for day_col in self.day_columns:
            code = row[day_col]
            
            # Skip truly empty cells
            if pd.isna(code) or (isinstance(code, str) and code.strip() == ''):
                continue
            
            # Count this as a day with data
            stats['Total_Days'] += 1
            
            # Parse the code
            parsed = self.parse_attendance_code(code)
            
            # Categorize day
            if parsed['is_weekly_off']:
                stats['Weekly_Off'] += 1
            elif parsed['is_leave']:
                stats['Leave_Days'] += 1
            elif parsed['is_absent']:
                stats['Absent_Days'] += 1
            elif parsed['shift']:
                # Working day with shift code
                stats['Present_Days'] += 1
                stats['Shift_Distribution'][parsed['shift']] += 1
                
                # Track station
                if parsed['station']:
                    stats['Stations'].add(parsed['station'])
        
        # Calculate attendance rate
        # Denominator = Total Days - Weekly Offs - Leaves
        expected_working_days = stats['Total_Days'] - stats['Weekly_Off'] - stats['Leave_Days']
        
        if expected_working_days > 0:
            stats['Attendance_Rate'] = (stats['Present_Days'] / expected_working_days) * 100
        else:
            stats['Attendance_Rate'] = 0.0
        
        # Convert stations set to comma-separated string
        stats['Stations'] = ', '.join(sorted(stats['Stations'])) if stats['Stations'] else 'N/A'
        
        return stats
    
    def get_all_employee_stats(self):
        """Calculate statistics for all employees"""
        stats_list = []
        
        for idx, row in self.df.iterrows():
            stats = self.calculate_employee_stats(row)
            stats_list.append(stats)
        
        stats_df = pd.DataFrame(stats_list)
        
        # Convert shift distribution to readable format
        stats_df['Most_Common_Shift'] = stats_df['Shift_Distribution'].apply(
            lambda x: self.SHIFT_CODES.get(x.most_common(1)[0][0], 'N/A') if x else 'N/A'
        )
        
        # Get shift counts as separate columns
        for shift_code, shift_name in self.SHIFT_CODES.items():
            stats_df[f'{shift_code}_Shift_Days'] = stats_df['Shift_Distribution'].apply(
                lambda x: x.get(shift_code, 0)
            )
        
        # Drop the Counter object for cleaner output
        stats_df = stats_df.drop('Shift_Distribution', axis=1)
        
        # Add employee details (designation, station, AM)
        emp_details = load_employee_details()
        if emp_details:
            stats_df['Designation'] = stats_df['Personnel_Number'].astype(str).map(emp_details['designation']).fillna('N/A')
            stats_df['Station'] = stats_df['Personnel_Number'].astype(str).map(emp_details['station']).fillna('')
            stats_df['AM'] = stats_df['Personnel_Number'].astype(str).map(emp_details['am']).fillna('')
            # Reorder columns to put details after Personnel_Number
            cols = stats_df.columns.tolist()
            # Remove the detail columns from their current position
            for col in ['Designation', 'Station', 'AM']:
                if col in cols:
                    cols.remove(col)
            # Insert after Personnel_Number
            idx = cols.index('Personnel_Number') + 1
            for col in reversed(['AM', 'Station', 'Designation']):
                cols.insert(idx, col)
            stats_df = stats_df[cols]
        
        return stats_df
    
    def get_daily_attendance(self):
        """Calculate daily attendance statistics - only for days that have data"""
        daily_stats = []
        
        # Only include days that have at least some data
        for i, day_col in enumerate(self.day_columns, 1):
            # Check if this day has any non-empty values
            day_data = self.df[day_col]
            has_data = any(code and not pd.isna(code) and str(code).strip() for code in day_data)
            
            if not has_data:
                continue  # Skip days with no data
            
            present = sum(1 for code in day_data if self.is_working_day(code))
            absent = sum(1 for code in day_data if self.is_absent(code))
            on_leave = sum(1 for code in day_data if self.is_on_leave(code))
            weekly_off = sum(1 for code in day_data if self.is_weekly_off(code))
            
            total_employees = len(self.df)
            # Expected = Total - Weekly Off - Leave
            expected_working = total_employees - weekly_off - on_leave
            
            attendance_rate = (present / expected_working * 100) if expected_working > 0 else 0
            
            daily_stats.append({
                'Day': i,
                'Day_Column': day_col,
                'Total_Employees': total_employees,
                'Present': present,
                'Absent': absent,
                'On_Leave': on_leave,
                'Weekly_Off': weekly_off,
                'Expected_Working': expected_working,
                'Attendance_Rate': round(attendance_rate, 2)
            })
        
        return pd.DataFrame(daily_stats)
    
    def get_summary_insights(self):
        """Generate overall summary insights"""
        stats_df = self.get_all_employee_stats()
        
        insights = {
            'Month': self.month_name,
            'Total_Employees': len(self.df),
            'Average_Attendance_Rate': round(stats_df['Attendance_Rate'].mean(), 2),
            'Median_Attendance_Rate': round(stats_df['Attendance_Rate'].median(), 2),
            'Min_Attendance_Rate': round(stats_df['Attendance_Rate'].min(), 2),
            'Max_Attendance_Rate': round(stats_df['Attendance_Rate'].max(), 2),
            'Total_Present_Days': int(stats_df['Present_Days'].sum()),
            'Total_Absent_Days': int(stats_df['Absent_Days'].sum()),
            'Total_Leave_Days': int(stats_df['Leave_Days'].sum()),
            'Total_Weekly_Offs': int(stats_df['Weekly_Off'].sum()),
            'Employees_Perfect_Attendance': int((stats_df['Attendance_Rate'] == 100).sum()),
            'Employees_Below_80_Percent': int((stats_df['Attendance_Rate'] < 80).sum()),
            'Employees_Below_90_Percent': int((stats_df['Attendance_Rate'] < 90).sum()),
            'Employees_Above_95_Percent': int((stats_df['Attendance_Rate'] >= 95).sum()),
        }
        
        return insights
    
    def get_top_performers(self, n=10):
        """Get employees with highest attendance"""
        stats_df = self.get_all_employee_stats()
        return stats_df.nlargest(n, 'Attendance_Rate')[
            ['Employee', 'Personnel_Number', 'Attendance_Rate', 'Present_Days', 'Absent_Days']
        ]
    
    def get_bottom_performers(self, n=10):
        """Get employees with lowest attendance"""
        stats_df = self.get_all_employee_stats()
        return stats_df.nsmallest(n, 'Attendance_Rate')[
            ['Employee', 'Personnel_Number', 'Attendance_Rate', 'Present_Days', 'Absent_Days', 'Leave_Days']
        ]
    
    def get_frequent_absentees(self, n=10):
        """Get employees with most absences"""
        stats_df = self.get_all_employee_stats()
        absentees = stats_df[stats_df['Absent_Days'] > 0].nlargest(n, 'Absent_Days')[
            ['Employee', 'Personnel_Number', 'Absent_Days', 'Present_Days', 'Attendance_Rate']
        ]
        return absentees
    
    def get_shift_analysis(self):
        """Analyze shift distribution"""
        shift_counts = Counter()
        
        for day_col in self.day_columns:
            for code in self.df[day_col]:
                parsed = self.parse_attendance_code(code)
                if parsed['shift']:
                    shift_counts[parsed['shift']] += 1
        
        total_shifts = sum(shift_counts.values())
        
        shift_distribution = []
        for shift_code in ['M', 'E', 'N', 'G']:  # Order: Morning, Evening, Night, General
            count = shift_counts.get(shift_code, 0)
            if count > 0 or total_shifts > 0:  # Show all shifts
                shift_distribution.append({
                    'Shift_Code': shift_code,
                    'Shift_Name': self.SHIFT_CODES.get(shift_code, shift_code),
                    'Count': count,
                    'Percentage': round((count / total_shifts * 100), 2) if total_shifts > 0 else 0
                })
        
        return pd.DataFrame(shift_distribution)
    
    def get_day_wise_trends(self):
        """Identify patterns in attendance by day"""
        daily_df = self.get_daily_attendance()
        
        trends = {
            'Average_Daily_Attendance_Rate': round(daily_df['Attendance_Rate'].mean(), 2),
            'Best_Attendance_Day': int(daily_df.loc[daily_df['Attendance_Rate'].idxmax(), 'Day']),
            'Worst_Attendance_Day': int(daily_df.loc[daily_df['Attendance_Rate'].idxmin(), 'Day']),
            'Most_Absences_Day': int(daily_df.loc[daily_df['Absent'].idxmax(), 'Day']),
            'Highest_Attendance_Rate': round(daily_df['Attendance_Rate'].max(), 2),
            'Lowest_Attendance_Rate': round(daily_df['Attendance_Rate'].min(), 2),
        }
        
        return trends


def compare_multiple_months(month_dataframes):
    """Compare attendance across multiple months"""
    comparison = []
    
    for month_name, df in month_dataframes.items():
        analyzer = AttendanceAnalyzer(df, month_name)
        summary = analyzer.get_summary_insights()
        comparison.append(summary)
    
    return pd.DataFrame(comparison)


if __name__ == "__main__":
    # Test analysis
    from pdf_extractor import AttendancePDFExtractor
    
    pdf_path = r"Test_reports\OCT_2025_SO-SC.pdf"
    extractor = AttendancePDFExtractor(pdf_path)
    df = extractor.create_dataframe()
    
    analyzer = AttendanceAnalyzer(df, "October 2025")
    print(analyzer.get_summary_insights())
