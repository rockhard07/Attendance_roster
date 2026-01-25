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
            
            # Extract all date headers from first header row (columns 3 onwards)
            date_headers = {}
            date_pattern = re.compile(r'([A-Za-z]+\.?\s+\d{1,2}\.\d{2})')
            
            if header_rows and len(header_rows) > 0:
                for col_idx in range(3, len(header_rows[0])):
                    header_val = str(header_rows[0][col_idx]).strip()
                    if header_val:
                        header_val = header_val.replace('\n', ' ').replace('\r', ' ')
                        header_val = ' '.join(header_val.split())
                        # Check if header matches date pattern (e.g., "Tue. 13.01")
                        if date_pattern.match(header_val):
                            date_headers[col_idx] = header_val

            for row in table[data_start:]:
                if not row or len(row) < 4:
                    continue

                employee_name = str(row[0]).strip().replace('\n', ' ').replace('\r', ' ') if row[0] else ""
                employee_name = ' '.join(employee_name.split())

                personnel_num = str(row[1]).strip() if row[1] else ""
                scheduling_row = str(row[2]).strip() if row[2] else ""

                if not employee_name or employee_name.upper() in ['EMPLOYEE', '', 'NONE']:
                    continue

                # Extract shift data from column 4 (index 3) and other date columns
                shift_data = {}  # {col_idx: {'shift': ..., 'sign_on': ..., 'sign_off': ...}}
                
                for col_idx in range(3, len(row)):
                    if col_idx in date_headers or col_idx == 3:
                        col_val = str(row[col_idx]) if col_idx < len(row) and row[col_idx] else ""
                        
                        shift = ""
                        sign_on = ""
                        sign_off = ""
                        
                        if col_val:
                            # Extract timing pattern (e.g., "05:00-13:00")
                            time_match = re.search(r"(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})", col_val)
                            if time_match:
                                sign_on = time_match.group(1).strip()
                                sign_off = time_match.group(2).strip()
                                # Shift is everything else (remove timing)
                                shift = col_val
                                shift = re.sub(r"\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}", "", shift).strip()
                            else:
                                # No timing, entire value is shift
                                shift = col_val.strip()
                        
                        shift_data[col_idx] = {
                            'shift': shift,
                            'sign_on': sign_on,
                            'sign_off': sign_off
                        }

                # Attendance codes start from the next column after the last date column
                last_date_col = max(date_headers.keys()) if date_headers else 3
                attendance_codes = []
                for i in range(last_date_col + 1, len(row)):
                    code = str(row[i]).strip() if row[i] else ""
                    code = code.replace('\n', '').replace('\r', '').strip()
                    attendance_codes.append(code)

                record = {
                    'Employee': employee_name,
                    'Personnel_Number': personnel_num,
                    'Scheduling_Row': scheduling_row,
                    'Attendance_Codes': attendance_codes,
                    'Shift_Data': shift_data,
                    'Date_Headers': date_headers
                }
                all_data.append(record)

        return all_data

    def create_dataframe(self):
        """Create structured dataframe from attendance data with split Sign On/Off times"""
        data = self.parse_attendance_data()

        if not data:
            return pd.DataFrame()

        max_days = max(len(record['Attendance_Codes']) for record in data) if data else 0

        is_paid_time_col = False
        if data:
            sample_last_values = [rec['Attendance_Codes'][-1] for rec in data[:5] if rec['Attendance_Codes']]
            time_pattern_count = sum(1 for val in sample_last_values if val and ':' in val and
                                     any(c.isdigit() for c in val.split(':')[0]))
            if time_pattern_count >= 3:
                is_paid_time_col = True
                max_days -= 1

        # Get date headers from first record
        date_headers = data[0].get('Date_Headers', {}) if data else {}
        
        # Extract dates from headers (e.g., "13.01" from "Tue. 13.01")
        date_map = {}  # {col_idx: date_str}
        for col_idx, header in date_headers.items():
            match = re.search(r'(\d{1,2}\.\d{2})', header)
            if match:
                date_map[col_idx] = match.group(1)

        # Build column order with date columns first
        shift_col_indices = sorted([col_idx for col_idx in date_headers.keys()])
        
        rows = []
        for record in data:
            row_dict = {
                'Employee': record.get('Employee', ''),
                'Personnel_Number': record.get('Personnel_Number', ''),
                'Scheduling_Row': record.get('Scheduling_Row', '')
            }
            
            shift_data = record.get('Shift_Data', {})
            
            # Add shift columns for each date in sorted order
            for col_idx in shift_col_indices:
                if col_idx in shift_data:
                    shift_info = shift_data[col_idx]
                    date_str = date_map.get(col_idx, '')
                    
                    # Create Shift, Sign On and Sign Off columns for this date
                    shift_col = f"Shift ({date_str})" if date_str else "Shift"
                    sign_on_col = f"Sign On time ({date_str})" if date_str else "Sign On time"
                    sign_off_col = f"Sign Off time ({date_str})" if date_str else "Sign Off time"
                    
                    row_dict[shift_col] = shift_info.get('shift', '')
                    row_dict[sign_on_col] = shift_info.get('sign_on', '')
                    row_dict[sign_off_col] = shift_info.get('sign_off', '')

            # Add attendance codes (daily)
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
