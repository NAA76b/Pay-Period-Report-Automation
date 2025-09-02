import os
import glob
import calendar
from pathlib import Path
from datetime import datetime
import pandas as pd
from rapidfuzz import process, fuzz
import configparser

# --- FILENAME PATTERNS ---
REPORT_1_PATTERN = "*Proxy_Part_Report_Hist_1of2.xlsx"
REPORT_2_PATTERN = "*Prev_Part_Report_EmpID_2of2.xlsx"


def get_path(config, key, default_subpath):
    """Retrieve a path from config or fall back to a default under the user's Desktop."""
    config_path_str = config.get('Paths', key, fallback=None)
    if config_path_str:
        config_path = Path(config_path_str)
        if config_path.exists():
            print(f"Using configured path for '{key}': {config_path}")
            return config_path

    home_dir = Path.home()
    default_path = home_dir / default_subpath
    print(f"Warning: Configured path for '{key}' not found. Trying default: {default_path}")
    if default_path.exists():
        print(f"Success: Found default path for '{key}'.")
        return default_path

    print(f"FATAL: Could not find a valid path for '{key}'.")
    return None


def load_schedule(schedule_path: Path):
    """Load the pay period schedule CSV."""
    if not schedule_path or not schedule_path.exists():
        print(f"Error: Pay Period schedule not found at '{schedule_path}'.")
        return None
    df = pd.read_csv(schedule_path)
    date_cols = ['PP Start Date', 'PP End Date', 'Initial Pull (Wed)', 'Final Pull & Email (Mon)']
    for col in date_cols:
        df[col] = pd.to_datetime(df[col])
    return df


def get_pp_details_from_date(file_date, schedule_df):
    for _, row in schedule_df.iterrows():
        if row['PP Start Date'].date() <= file_date.date() <= row['Initial Pull (Wed)'].date():
            day_str = {0: "MON", 2: "WED"}.get(file_date.weekday(), "PULL")
            start_str = f"{calendar.month_name[row['PP Start Date'].month]} {row['PP Start Date'].day}"
            end_str = f"{calendar.month_name[row['PP End Date'].month]} {row['PP End Date'].day}"
            fy = row['PP Start Date'].year + 1 if row['PP Start Date'].month >= 10 else row['PP Start Date'].year
            folder_name = f"FY{str(fy)[-2:]} PP{row['Pay Period']} - {start_str} - {end_str}"
            return {"pp_num": row['Pay Period'], "day": day_str, "folder": folder_name}
    return None


def find_header_row(df_no_header, expected_headers):
    for i in range(min(20, len(df_no_header))):
        row_str = ' '.join(map(str, df_no_header.iloc[i].dropna().tolist()))
        if not row_str:
            continue
        avg_score = sum(fuzz.partial_ratio(h, row_str) for h in expected_headers) / len(expected_headers)
        if avg_score > 80:
            return i
    return 14


def merge_reports(report1_path: Path, report2_path: Path, pp_details: dict, final_reports_root: Path):
    headers1 = ['Super Office', 'Time Sheet: Owner Name', 'Sum of Hours']
    df1 = pd.read_excel(report1_path, header=find_header_row(pd.read_excel(report1_path, header=None), headers1))

    headers2 = ['Time Sheet: Owner Name', 'FDA Employee Number', 'Sum of Hours']
    df2 = pd.read_excel(report2_path, header=find_header_row(pd.read_excel(report2_path, header=None), headers2))

    merged_data, mismatched_log = [], []
    df2_names = df2['Time Sheet: Owner Name'].dropna().astype(str).tolist()

    for _, row1 in df1.iterrows():
        name1 = str(row1.get('Time Sheet: Owner Name', ''))
        if not name1 or pd.isna(row1.get('Time Sheet: Owner Name')):
            continue
        new_row = row1.to_dict()
        best_match = process.extractOne(name1, df2_names, scorer=fuzz.token_set_ratio, score_cutoff=85)
        if best_match:
            row2 = df2[df2['Time Sheet: Owner Name'] == best_match[0]].iloc[0]
            new_row['FDA Employee Number'] = row2.get('FDA Employee Number')
        else:
            new_row['FDA Employee Number'] = 'NOT FOUND'
            mismatched_log.append(f"No match for: '{name1}'")
        merged_data.append(new_row)

    merged_df = pd.DataFrame(merged_data)
    destination_folder = final_reports_root / pp_details["folder"]
    destination_folder.mkdir(exist_ok=True)
    final_filename = f"PP{pp_details['pp_num']} {pp_details['day']} - Participation Data Combined.xlsx"
    final_report_path = destination_folder / final_filename
    merged_df.to_excel(final_report_path, index=False)
    print(f"Successfully created: {final_report_path}")

    if mismatched_log:
        log_path = destination_folder / f"PP{pp_details['pp_num']}_Mismatch_Log.txt"
        with open(log_path, 'w') as f:
            f.write('\n'.join(mismatched_log))

    return True


def main():
    config = configparser.ConfigParser()
    config.read('config.ini')

    landing_zone = get_path(config, 'landing_zone', 'Desktop/PayPeriodAttachments')
    final_reports_root = get_path(config, 'final_reports_root', 'Desktop/CBER ITR Participation Clean Up')
    schedule_csv_path = get_path(config, 'pay_period_schedule', 'PayPeriodAutomation/pay_periods.csv')

    if not all([landing_zone, final_reports_root, schedule_csv_path]):
        print("Exiting due to missing paths. Please check your config.ini or folder structure.")
        return

    schedule = load_schedule(schedule_csv_path)
    if schedule is None:
        return

    print(f"Scanning landing zone: {landing_zone}")
    report1 = glob.glob(str(landing_zone / REPORT_1_PATTERN))
    report2 = glob.glob(str(landing_zone / REPORT_2_PATTERN))

    if report1 and report2:
        report1_path, report2_path = Path(report1[0]), Path(report2[0])
        date_str = report1_path.name.split('_')[0]
        pp_details = get_pp_details_from_date(datetime.strptime(date_str, '%Y-%m-%d'), schedule)
        if not pp_details:
            print(f"Error: No Pay Period found for date {date_str}.")
            return
        try:
            if merge_reports(report1_path, report2_path, pp_details, final_reports_root):
                os.remove(report1_path)
                os.remove(report2_path)
                print("Process complete.")
        except Exception as e:
            print(f"ERROR during merge: {e}")
    else:
        print("Did not find both required report files in landing zone.")


if __name__ == "__main__":
    main()
