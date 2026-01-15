"""
PDF Attendance Data Extractor
Extracts attendance data from PDF files containing employee attendance tables
- RosterExtractor: Simple attendance extraction (Stations/OCC/Train Ops Roster) - no Shift/Shift_Timings
- AttendancePDFExtractor: Trip Chart extraction with Shift/Shift_Timings/Paid_Time
"""

import pdfplumber
import pandas as pd
import re
from datetime import datetime
from pathlib import Path


class RosterExtractor:
    """Extract and parse roster data from PDF files (Stations/OCC/Train Ops Roster)"""
    
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.raw_tables = []
        self.df = None
        
    def extract_tables(self):
        """Extract all tables from PDF"""
        with pdfplumber.open(self.pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    self.raw_tables.extend(tables)
        return self.raw_tables
    
    def parse_date_headers(self, header_rows):
        """Parse date information from header rows"""
        dates = []
        if len(header_rows) >= 2:
            days = header_rows[0]
            date_nums = header_rows[1]
            
            for i in range(len(days)):
                if days[i] and date_nums[i]:
                    day = str(days[i]).strip()
                    date_num = str(date_nums[i]).strip()
                    if day and date_num and date_num.isdigit():
                        dates.append(f"{day}-{date_num}")
                    else:
                        dates.append(f"Day-{i}")
        return dates
    
    def parse_attendance_data(self):
        """Parse attendance data from extracted tables"""
        if not self.raw_tables:
            self.extract_tables()
        
        all_data = []
        
        for table in self.raw_tables:
            if len(table) < 3:
                continue
                
            header_start = 0
            data_start = None
            
            for i, row in enumerate(table):
                if row and len(row) > 0:
                    first_cell = str(row[0]).strip().upper()
                    if first_cell and first_cell not in ['EMPLOYEE', '', 'NONE']:
                        data_start = i
                        break
            
            if data_start is None:
                continue
            
            header_rows = table[max(0, data_start-2):data_start]
            
            for row in table[data_start:]:
                if not row or len(row) < 4:
                    continue
                    
                employee_name = str(row[0]).strip().replace('\n', ' ').replace('\r', ' ') if row[0] else ""
                employee_name = ' '.join(employee_name.split())
                
                personnel_num = str(row[1]).strip() if row[1] else ""
                scheduling_row = str(row[2]).strip() if row[2] else ""
                
                if not employee_name or employee_name.upper() in ['EMPLOYEE', '', 'NONE']:
                    continue
                
                # For roster, column 3 onwards are attendance codes (no Shift extraction)
                attendance_codes = []
                for i in range(3, len(row)):
                    code = str(row[i]).strip() if row[i] else ""
                    code = code.replace('\n', '').replace('\r', '').strip()
                    attendance_codes.append(code)
                
                record = {
                    'Employee': employee_name,
                    'Personnel_Number': personnel_num,
                    'Scheduling_Row': scheduling_row,
                    'Attendance_Codes': attendance_codes
                }
                all_data.append(record)
        
        return all_data
    
    def create_dataframe(self):
        """Create structured dataframe from attendance data"""
        data = self.parse_attendance_data()
        
        if not data:
            return pd.DataFrame()
        
        max_days = max(len(record['Attendance_Codes']) for record in data)
        
        is_paid_time_col = False
        if data:
            sample_last_values = [rec['Attendance_Codes'][-1] for rec in data[:5] if rec['Attendance_Codes']]
            time_pattern_count = sum(1 for val in sample_last_values if val and ':' in val and 
                                     any(c.isdigit() for c in val.split(':')[0]))
            if time_pattern_count >= 3:
                is_paid_time_col = True
                max_days -= 1
        
        rows = []
        for record in data:
            row_dict = {
                'Employee': record['Employee'],
                'Personnel_Number': record['Personnel_Number'],
                'Scheduling_Row': record['Scheduling_Row']
            }
            
            attendance_codes = record['Attendance_Codes']
            if is_paid_time_col and len(attendance_codes) > max_days:
                paid_time = attendance_codes[-1]
                attendance_codes = attendance_codes[:max_days]
                row_dict['Paid_Time'] = paid_time
            
            for i, code in enumerate(attendance_codes):
                row_dict[f'Day_{i+1}'] = code
            
            for i in range(len(attendance_codes), max_days):
                row_dict[f'Day_{i+1}'] = ""
            
            rows.append(row_dict)
        
        self.df = pd.DataFrame(rows)
        return self.df
    
    def get_month_from_filename(self):
        """Extract month/year from filename"""
        filename = Path(self.pdf_path).stem
        match = re.search(r'([A-Z]+)_(\d{4})', filename)
        if match:
            month = match.group(1)
            year = match.group(2)
            return f"{month}_{year}"
        return filename


class AttendancePDFExtractor:
    """Extract and parse attendance data from PDF files (Trip Chart only)
    
    Extracts Shift, Shift_Timings, and Paid_Time columns in addition to daily attendance.
    """

    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.raw_tables = []
        self.df = None

    def extract_tables(self):
        """Extract all tables from PDF"""
        with pdfplumber.open(self.pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    self.raw_tables.extend(tables)
        return self.raw_tables

    def parse_date_headers(self, header_rows):
        """Parse date information from header rows"""
        dates = []
        if len(header_rows) >= 2:
            days = header_rows[0]
            date_nums = header_rows[1]

            for i in range(len(days)):
                if days[i] and date_nums[i]:
                    day = str(days[i]).strip()
                    date_num = str(date_nums[i]).strip()
                    if day and date_num and date_num.isdigit():
                        dates.append(f"{day}-{date_num}")
                    else:
                        dates.append(f"Day-{i}")
        return dates

    def parse_attendance_data(self):
        """Parse attendance data from extracted tables"""
        if not self.raw_tables:
            self.extract_tables()

        all_data = []

        for table in self.raw_tables:
            if len(table) < 3:
                continue

            header_start = 0
            data_start = None

            for i, row in enumerate(table):
                if row and len(row) > 0:
                    first_cell = str(row[0]).strip().upper()
                    if first_cell and first_cell not in ['EMPLOYEE', '', 'NONE']:
                        data_start = i
                        break

            if data_start is None:
                continue

            header_rows = table[max(0, data_start-2):data_start]

            for row in table[data_start:]:
                if not row or len(row) < 4:
                    continue

                employee_name = str(row[0]).strip().replace('\n', ' ').replace('\r', ' ') if row[0] else ""
                employee_name = ' '.join(employee_name.split())

                personnel_num = str(row[1]).strip() if row[1] else ""
                scheduling_row = str(row[2]).strip() if row[2] else ""

                if not employee_name or employee_name.upper() in ['EMPLOYEE', '', 'NONE']:
                    continue

                # Extract column 4 (index 3) as Shift - unconditionally
                shift = ""
                shift_timing = ""
                if len(row) > 3 and row[3]:
                    col4 = str(row[3])
                    # First, try to extract timing pattern (e.g., "05:00-13:00")
                    time_match = re.search(r"(\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2})", col4)
                    if time_match:
                        shift_timing = time_match.group(1).strip()
                        # Shift is everything else
                        shift = col4.replace(shift_timing, '').strip()
                    else:
                        # No timing pattern found, entire col4 is shift
                        shift = col4.strip()
                else:
                    shift = ""

                # Attendance codes start from column 4 (index 4) onwards
                attendance_codes = []
                for i in range(4, len(row)):
                    code = str(row[i]).strip() if row[i] else ""
                    code = code.replace('\n', '').replace('\r', '').strip()
                    attendance_codes.append(code)

                record = {
                    'Employee': employee_name,
                    'Personnel_Number': personnel_num,
                    'Scheduling_Row': scheduling_row,
                    'Attendance_Codes': attendance_codes,
                    'Shift': shift,
                    'Shift_Timings': shift_timing
                }
                all_data.append(record)

        return all_data

    def create_dataframe(self):
        """Create structured dataframe from attendance data"""
        data = self.parse_attendance_data()

        if not data:
            return pd.DataFrame()

        max_days = max(len(record['Attendance_Codes']) for record in data)

        is_paid_time_col = False
        if data:
            sample_last_values = [rec['Attendance_Codes'][-1] for rec in data[:5] if rec['Attendance_Codes']]
            time_pattern_count = sum(1 for val in sample_last_values if val and ':' in val and
                                     any(c.isdigit() for c in val.split(':')[0]))
            if time_pattern_count >= 3:
                is_paid_time_col = True
                max_days -= 1

        rows = []
        for record in data:
            row_dict = {
                'Employee': record.get('Employee', ''),
                'Personnel_Number': record.get('Personnel_Number', ''),
                'Scheduling_Row': record.get('Scheduling_Row', ''),
                'Shift': record.get('Shift', ''),
                'Shift_Timings': record.get('Shift_Timings', '')
            }

            attendance_codes = record['Attendance_Codes'].copy() if record['Attendance_Codes'] else []
            paid_time = ""
            if is_paid_time_col and len(attendance_codes) > max_days:
                paid_time = attendance_codes[-1]
                attendance_codes = attendance_codes[:max_days]
            
            row_dict['Paid_Time'] = paid_time

            for i, code in enumerate(attendance_codes):
                row_dict[f'Day_{i+1}'] = code

            for i in range(len(attendance_codes), max_days):
                row_dict[f'Day_{i+1}'] = ""

            rows.append(row_dict)

        self.df = pd.DataFrame(rows)
        return self.df

    def get_month_from_filename(self):
        """Extract month/year from filename"""
        filename = Path(self.pdf_path).stem
        match = re.search(r'([A-Z]+)_(\d{4})', filename)
        if match:
            month = match.group(1)
            year = match.group(2)
            return f"{month}_{year}"
        return filename


if __name__ == "__main__":
    print("PDF Extractor module loaded successfully")
