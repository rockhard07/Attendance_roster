import pandas as pd
import numpy as np
import re
from pathlib import Path

def extract_times(cell_str):
    """
    Extract Sign-On and Sign-Off times from duty strings.
    Format: SHIFT-STATION HH:MM-HH:MM (e.g., N-RITH 22:00-07:00)
    Regex: (\\d{2}:\\d{2})-(\\d{2}:\\d{2})$
    """
    if not isinstance(cell_str, str) or not cell_str.strip():
        return None
    
    # Clean string and search for time pattern at the end
    match = re.search(r'(\d{2}:\d{2})-(\d{2}:\d{2})$', cell_str.strip())
    if match:
        return match.groups()  # (sign_on, sign_off)
    return None

def to_minutes(time_str):
    """Convert HH:MM string to minutes from start of day."""
    if not time_str:
        return -1
    try:
        h, m = map(int, time_str.split(':'))
        return h * 60 + m
    except (ValueError, AttributeError):
        return -1

def calculate_nda(sign_on_str, sign_off_str):
    """
    Calculate NDA category and amount based on business rules.
    Rules:
    N4 (175): Sign Off [01:00-06:00] OR Sign On [01:00-02:00]
    N3 (140): Sign Off [00:01-00:59] OR Sign On [02:01-03:00]
    N2 (120): Sign Off [23:01-23:59] OR Sign On [03:01-04:00]
    N1 (90):  Sign Off [22:00-22:59] OR Sign On [04:01-05:00]
    """
    on_m = to_minutes(sign_on_str)
    off_m = to_minutes(sign_off_str)
    
    if on_m == -1 or off_m == -1:
        return None, 0

    # Handle "Midnight Wrap"
    # If Sign-Off is 00:00 and Sign-On was late in the evening (e.g., after 12:00)
    # treat Sign-Off as 24:00 (1440 minutes).
    if off_m == 0 and on_m > 720:  # 12:00 = 720 mins
        off_m = 1440

    # Define Categories
    categories = [
        {'name': 'N4', 'value': 175, 
         'off_range': (60, 360), 'on_range': (60, 120)},   # 01:00-06:00, 01:00-02:00
        {'name': 'N3', 'value': 140, 
         'off_range': (0, 59), 'on_range': (121, 180)},    # 00:00-00:59, 02:01-03:00
        {'name': 'N2', 'value': 120, 
         'off_range': (1380, 1439), 'on_range': (181, 240)}, # 23:00-23:59, 03:01-04:00
        {'name': 'N1', 'value': 90, 
         'off_range': (1320, 1379), 'on_range': (241, 300)}  # 22:00-22:59, 04:01-05:00
    ]

    best_cat = None
    max_val = 0

    for cat in categories:
        # Check Sign Off
        off_match = cat['off_range'][0] <= off_m <= cat['off_range'][1]
        
        # Explicit check for 1440 (00:00 after midnight wrap) in N3
        if cat['name'] == 'N3' and off_m == 1440:
            off_match = True
            
        # Check Sign On
        on_match = cat['on_range'][0] <= on_m <= cat['on_range'][1]
        
        if off_match or on_match:
            if cat['value'] > max_val:
                max_val = cat['value']
                best_cat = cat['name']
                
    return best_cat, max_val

def generate_nda_report(df, designation_map=None):
    """
    Processes the DataFrame and generates the NDA report.
    """
    # Identify date columns (Day_1, Day_2, etc.)
    day_columns = [col for col in df.columns if col.startswith('Day_')]
    
    report_rows = []
    
    for _, row in df.iterrows():
        emp_name = str(row.get('Employee', '')).strip()
        personnel_num = str(row.get('Personnel_Number', '')).strip()
        
        # Skip header rows or empty rows
        if not emp_name or emp_name.upper() in ['EMPLOYEE', 'NONE', '']:
            continue
            
        # Get designation
        designation = 'N/A'
        if 'Designation' in row:
            designation = row['Designation']
        elif designation_map and personnel_num in designation_map:
            designation = designation_map[personnel_num]
            
        emp_record = {
            'Employee': emp_name,
            'Personnel Number': personnel_num,
            'Designation': designation
        }
        
        counts = {'N1': 0, 'N2': 0, 'N3': 0, 'N4': 0}
        total_allowance = 0
        
        for day_col in day_columns:
            cell_value = str(row.get(day_col, '')).strip()
            times = extract_times(cell_value)
            
            result_str = ""
            if times:
                sign_on, sign_off = times
                cat, cost = calculate_nda(sign_on, sign_off)
                if cat:
                    result_str = f"{cat} ({cost})"
                    counts[cat] += 1
                    total_allowance += cost
            
            emp_record[day_col] = result_str
            
        # Add summary columns
        for cat in ['N1', 'N2', 'N3', 'N4']:
            emp_record[f'{cat} Count'] = counts[cat]
            
        emp_record['Total Allowance (INR)'] = total_allowance
        report_rows.append(emp_record)
        
    return pd.DataFrame(report_rows)

def load_designations():
    """Helper to load designations from the known project structure."""
    try:
        # Check current directory and its parent for Report IVU
        paths_to_check = [
            Path("Report IVU") / "Employee details" / "employee_details.xlsx",
            Path("..") / "Report IVU" / "Employee details" / "employee_details.xlsx",
            Path(__file__).parent / "Report IVU" / "Employee details" / "employee_details.xlsx"
        ]
        
        for path in paths_to_check:
            if path.exists():
                print(f"Loading designations from {path}")
                df = pd.read_excel(path, sheet_name='Combined')
                # Clean up Emp Id to handle potential floats/whitespace
                df['Emp Id'] = df['Emp Id'].astype(str).str.split('.').str[0].str.strip()
                return dict(zip(df['Emp Id'], df['Designation']))
    except Exception as e:
        print(f"Warning: Could not load designations: {e}")
    return {}

def main():
    import sys
    
    # If a file is passed as an argument, process it
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        try:
            print(f"Processing {input_file}...")
            df = pd.read_excel(input_file)
            designations = load_designations()
            report_df = generate_nda_report(df, designation_map=designations)
            
            output_file = "NDA_Allowance_Report.xlsx"
            report_df.to_excel(output_file, index=False)
            print(f"Report generated successfully: {output_file}")
            print(report_df.head())
        except Exception as e:
            print(f"Error processing file: {e}")
        return

    # Default Example Usage / Test
    data = {
        'Employee': ['John Doe', 'Jane Smith', 'Bob Wilson'],
        'Personnel_Number': ['1001', '1002', '1003'],
        'Day_1': ['N-RITH 22:00-07:00', 'E-KASH 14:00-22:00', 'N-RITH 23:30-07:30'],
        'Day_2': ['N-RITH 22:00-00:00', 'N-KASH 01:30-09:30', 'OFF'],
        'Day_3': ['N-RITH 22:00-00:30', 'N-KASH 03:30-11:30', 'N-RITH 04:30-12:30']
    }
    df = pd.DataFrame(data)
    
    print("Running with sample data...")
    designations = {'1001': 'Train Operator', '1002': 'Train Operator', '1003': 'Shunter'}
    
    report_df = generate_nda_report(df, designation_map=designations)
    
    print("\nNDA Report Summary:")
    cols_to_show = ['Employee', 'Personnel Number', 'Total Allowance (INR)', 'N4 Count', 'N3 Count', 'N2 Count', 'N1 Count']
    print(report_df[cols_to_show].to_string(index=False))

if __name__ == "__main__":
    main()
