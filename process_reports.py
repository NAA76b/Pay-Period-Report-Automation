import os
import shutil
from pathlib import Path
import pandas as pd
from rapidfuzz import process, fuzz
from datetime import datetime
import glob
import calendar

# --- CONFIGURATION ---
LANDING_ZONE = Path(r"C:\Users\Nathan.Allen\OneDrive - FDA\Desktop\PayPeriodAttachments")
FINAL_REPORTS_ROOT = Path(r"C:\Users\Nathan.Allen\OneDrive - FDA\Desktop\CBER ITR Participation Clean Up")
PAY_PERIOD_SCHEDULE_CSV = Path("pay_periods.csv")

REPORT_1_PATTERN = "*Proxy_Part_Report_Hist_1of2.xlsx"
REPORT_2_PATTERN = "*Prev_Part_Report_EmpID_2of2.xlsx"
# ---

def load_schedule():
    """Loads the pay period schedule from CSV."""
    if not PAY_PERIOD_SCHEDULE_CSV.exists():
        print(f"Error: Pay Period schedule '{PAY_PERIOD_SCHEDULE_CSV}' not found.")
        return None
    df = pd.read_csv(PAY_PERIOD_SCHEDULE_CSV)
    date_cols = ['PP Start Date', 'PP End Date', 'Initial Pull (Wed)', 'Final Pull & Email (Mon)']
    for col in date_cols:
        df[col] = pd.to_datetime(df[col])
    return df

def get_pp_details_from_date(file_date, schedule_df):
    """Finds the Pay Period info and folder name for a given date."""
    for _, row in schedule_df.iterrows():
        if row['PP Start Date'].date() <= file_date.date() <= row['Initial Pull (Wed)'].date():
            # Determine if it's a Monday or Wednesday pull
            day_of_week = file_date.weekday() # Monday is 0, Wednesday is 2
            day_str = "MON" if day_of_week == 0 else "WED" if day_of_week == 2 else "PULL"
            
            # Format folder name
            start_str = f"{calendar.month_name[row['PP Start Date'].month]} {row['PP Start Date'].day}"
            end_str = f"{calendar.month_name[row['PP End Date'].month]} {row['PP End Date'].day}"
            year = row['PP Start Date'].year
            fy = year + 1 if row['PP Start Date'].month >= 10 else year
            fy_short = str(fy)[-2:]
            
            folder_name = f"FY{fy_short} PP{row['Pay Period']} - {start_str} - {end_str}"
            return {"pp_num": row['Pay Period'], "day": day_str, "folder": folder_name}
    return None

def find_header_row(df_no_header, expected_headers):
    """Finds the header row in a dataframe using fuzzy matching."""
    for i in range(min(20, len(df_no_header))):
        row_str = ' '.join(map(str, df_no_header.iloc[i].dropna().tolist()))
        if not row_str: continue
        avg_score = sum(fuzz.partial_ratio(h, row_str) for h in expected_headers) / len(expected_headers)
        if avg_score > 80:
            print(f"Header row found at index {i} with score {avg_score:.2f}%")
            return i
    print("Warning: Header row not found. Using default row 14.")
    return 14

def merge_reports(report1_path, report2_path, pp_details):
    """Loads, merges, and saves the two Excel reports with smart naming."""
    print("Starting merge process...")
    
    headers1 = ['Super Office', 'Time Sheet: Owner Name', 'Sum of Hours']
    df1 = pd.read_excel(report1_path, header=find_header_row(pd.read_excel(report1_path, header=None), headers1))
    
    headers2 = ['Time Sheet: Owner Name', 'FDA Employee Number', 'Sum of Hours']
    df2 = pd.read_excel(report2_path, header=find_header_row(pd.read_excel(report2_path, header=None), headers2))

    merged_data, mismatched_log = [], []
    df2_names = df2['Time Sheet: Owner Name'].dropna().astype(str).tolist()
    
    for _, row1 in df1.iterrows():
        name1 = str(row1.get('Time Sheet: Owner Name', ''))
        if not name1 or pd.isna(row1.get('Time Sheet: Owner Name')): continue
        new_row = row1.to_dict()
        best_match = process.extractOne(name1, df2_names, scorer=fuzz.token_set_ratio, score_cutoff=85)
        
        if best_match:
            row2 = df2[df2['Time Sheet: Owner Name'] == best_match[0]].iloc[0]
            new_row['FDA Employee Number'] = row2.get('FDA Employee Number')
        else:
            new_row['FDA Employee Number'] = 'NOT FOUND'
            mismatched_log.append(f"No match found for: '{name1}'")
        merged_data.append(new_row)

    merged_df = pd.DataFrame(merged_data)
    
    # Create smart destination folder and filenames
    destination_folder = FINAL_REPORTS_ROOT / pp_details["folder"]
    destination_folder.mkdir(exist_ok=True)
    
    final_filename = f"PP{pp_details['pp_num']} {pp_details['day']} - Participation Data Combined.xlsx"
    final_report_path = destination_folder / final_filename
    
    merged_df.to_excel(final_report_path, index=False)
    print(f"Successfully created combined report: {final_report_path}")

    if mismatched_log:
        log_path = destination_folder / f"PP{pp_details['pp_num']}_Mismatch_Log.txt"
        with open(log_path, 'w') as f: f.write('\n'.join(mismatched_log))
        print(f"Mismatch log created: {log_path}")
    return True

def main():
    """Main function to find and process report files."""
    schedule = load_schedule()
    if schedule is None: return

    print(f"Scanning landing zone: {LANDING_ZONE}")
    report1_files = glob.glob(str(LANDING_ZONE / REPORT_1_PATTERN))
    report2_files = glob.glob(str(LANDING_ZONE / REPORT_2_PATTERN))
            
    if report1_files and report2_files:
        report1_path = Path(report1_files[0])
        report2_path = Path(report2_files[0])
        
        # Get date from filename to find PP details
        date_str = report1_path.name.split('_')[0]
        file_date = datetime.strptime(date_str, '%Y-%m-%d')
        pp_details = get_pp_details_from_date(file_date, schedule)

        if not pp_details:
            print(f"Error: Could not find a Pay Period for date {date_str}. Please check pay_periods.csv.")
            return

        print(f"Found files for {pp_details['folder']}")
        
        try:
            if merge_reports(report1_path, report2_path, pp_details):
                print("Cleaning up landing zone...")
                os.remove(report1_path)
                os.remove(report2_path)
                print("Process complete.")
        except Exception as e:
            print(f"\n--- AN ERROR OCCURRED ---\nError during merge: {e}\nFiles left in landing zone.")
    else:
        print("\nDid not find both required report files in the landing zone.")

if __name__ == "__main__":
    main()
