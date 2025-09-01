# Pay-Period-Report-Automation
Desktop application to automate downloading, saving, and merging Excel reports from Outlook emails for pay period tracking

Pay Period Report Automation Application
🎯 Objective
This project is a Python desktop application designed to fully automate the workflow of processing recurring pay period reports received via email. It features a graphical user interface (GUI) for ease of use by non-technical staff.

✨ Core Features
User-Friendly GUI: A simple interface to start the process and view logs. No command-line interaction needed.

Configuration File: All critical settings (email, folders, report subjects) are stored in a simple config.ini file, allowing for easy changes without editing code.

Secure Password Entry: The application prompts for the email password at runtime and does not store it in the configuration file.

Automated Email Monitoring: Uses exchangelib to connect to an Outlook/Exchange inbox.

Fuzzy Subject Matching: Intelligently finds the correct report emails even if the subject lines are not exact matches, using rapidfuzz.

Dynamic Folder Creation: Automatically creates a structured folder hierarchy for each pay period based on the email's received date and a pay_periods.csv schedule file.

Intelligent Excel Parsing: Finds the correct header row in the Excel reports, even if it's not on the first line.

Automated Report Merging: Merges the two downloaded reports into a single, consolidated .xlsx file, using fuzzy name matching to align rows.

Error & Mismatch Logging: Generates a Mismatch_Log.txt for any names that couldn't be matched during the merge process and keeps a general automation_log.log.

Bundled Executable: Can be easily packaged into a single .exe file for distribution on Windows, hiding the console window for a seamless user experience.

📂 Project File Structure
The project must be created with the following file structure:

/PayPeriodAutomation/
├── app.py                     # The main GUI application (CustomTkinter)
├── automation_worker.py       # The core email/Excel processing logic
├── config_manager.py          # Handles loading/saving settings from config.ini
├── pay_period_manager.py      # Manages pay period logic based on the CSV schedule
├── config.ini                 # User-configurable settings
├── pay_periods.csv            # The pay period schedule data
└── requirements.txt           # List of required Python libraries
📄 File Contents
Below is the complete source code and content for each file in the project.

requirements.txt
customtkinter
exchangelib
pandas
openpyxl
rapidfuzz
pay_periods.csv
Code snippet

Pay Period,PP Start Date,PP End Date,Report Due Date,Reporting Deadline Text,PP Close Date,Initial Pull (Wed),Final Pull & Email (Mon),Is the Pay Period Complete?
16,2025-07-13,2025-07-26,2025-07-28,"ITR reporting due by 11:59 pm ET on Monday, July 28",2025-07-26,2025-07-30,2025-07-28,Yes
17,2025-07-27,2025-08-09,2025-08-11,"ITR reporting due by 11:59 pm ET on Monday, August 11",2025-08-09,2025-08-13,2025-08-11,No
18,2025-08-10,2025-08-23,2025-08-25,"ITR reporting due by 11:59 pm ET on Monday, August 25",2025-08-23,2025-08-27,2025-08-25,Not Started
19,2025-08-24,2025-09-06,2025-09-08,"ITR reporting due by 11:59 pm ET on Monday, September 08",2025-09-06,2025-09-10,2025-09-08,Not Started
config.ini
Ini, TOML

[Email]
address = nathan.allen@fda.hhs.gov
server = outlook.office365.com
# IMPORTANT: Password will be requested securely when the app runs.
# It is NOT saved here for security reasons.

[Folders]
root_path = C:\Users\Nathan.Allen\OneDrive - FDA\Desktop\CBER ITR Participation Clean Up
pay_period_schedule_csv = pay_periods.csv
log_file = automation_log.log

[Reports]
report1_subject = Proxy Part. Report - Hist (1 of 2)
report2_subject = Prev Part. Report - Emp ID (2 of 2)

[FuzzyLogic]
subject_match_threshold = 75
name_match_threshold = 85

[Application]
theme = Dark-Blue
# Possible themes: Dark-Blue, Blue, Green
config_manager.py
Python

import configparser

class ConfigManager:
    """Handles loading and saving application settings from config.ini."""
    def __init__(self, path='config.ini'):
        self.path = path
        self.config = configparser.ConfigParser()
        self.load_config()

    def load_config(self):
        """Loads the configuration from the file."""
        self.config.read(self.path)

    def get(self, section, key):
        """Gets a value from the config."""
        return self.config.get(section, key)

    def set(self, section, key, value):
        """Sets a value in the config."""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, value)

    def save_config(self):
        """Saves the current configuration to the file."""
        with open(self.path, 'w') as configfile:
            self.config.write(configfile)
pay_period_manager.py
Python

import pandas as pd
from datetime import datetime
import calendar
import os

class PayPeriodManager:
    """Manages the pay period schedule and folder structure logic."""
    def __init__(self, schedule_path):
        self.schedule_path = schedule_path
        self.schedule_df = None
        self.load_schedule()

    def load_schedule(self):
        """Loads the pay period schedule from a CSV file."""
        if not os.path.exists(self.schedule_path):
            raise FileNotFoundError(f"Pay period schedule not found at: {self.schedule_path}")
            
        self.schedule_df = pd.read_csv(self.schedule_path)
        # Convert date columns to datetime objects for comparison
        date_cols = ['PP Start Date', 'PP End Date', 'Initial Pull (Wed)', 'Final Pull & Email (Mon)']
        for col in date_cols:
            self.schedule_df[col] = pd.to_datetime(self.schedule_df[col])
        
        # Ensure 'Is the Pay Period Complete?' is treated as a string for 'Yes' comparison
        self.schedule_df['Is the Pay Period Complete?'] = self.schedule_df['Is the Pay Period Complete?'].astype(str).str.strip()

    def find_pp_for_date(self, received_date: datetime):
        """Finds the correct pay period for a given email received date."""
        # Ensure the received_date is timezone-naive for comparison
        received_date = received_date.replace(tzinfo=None)
        
        for index, row in self.schedule_df.iterrows():
            # A report belongs to a PP if it's received after the PP starts and before the final pull date + a buffer
            start = row['PP Start Date']
            # We look up to the *next* PP's start date to define the window
            end = self.schedule_df.loc[index + 1, 'PP Start Date'] if index + 1 < len(self.schedule_df) else row['Final Pull & Email (Mon)'] + pd.Timedelta(days=7)
            
            if start <= received_date < end:
                if row['Is the Pay Period Complete?'].lower() == 'yes':
                    return None, "Pay period is already marked as complete."
                return row, None
        return None, "No active pay period found for this date."

    def get_folder_details(self, pp_row, received_date: datetime):
        """Determines the full path and subfolder for saving an attachment."""
        received_date = received_date.replace(tzinfo=None)

        # Format Month and Day for folder name
        start_day_str = f"{calendar.month_name[pp_row['PP Start Date'].month]} {pp_row['PP Start Date'].day}"
        end_day_str = f"{calendar.month_name[pp_row['PP End Date'].month]} {pp_row['PP End Date'].day}"
        
        # Determine Fiscal Year (starts in October)
        year = pp_row['PP Start Date'].year
        fiscal_year_start_month = 10
        fy = year + 1 if pp_row['PP Start Date'].month >= fiscal_year_start_month else year
        fy_short = str(fy)[-2:]

        pp_folder_name = f"FY{fy_short} PP{pp_row['Pay Period']} - {start_day_str} - {end_day_str}"

        # Determine subfolder
        if received_date.date() <= pp_row['Final Pull & Email (Mon)'].date():
            subfolder = "Final Pull & Email (Mon)"
        else:
            subfolder = "Initial Pull (Wed)"
            
        return pp_folder_name, subfolder
automation_worker.py
Python

import os
import logging
from pathlib import Path
from datetime import datetime
import time

import pandas as pd
from rapidfuzz import fuzz, process
from exchangelib import Credentials, Account, Configuration, DELEGATE

from pay_period_manager import PayPeriodManager

class AutomationWorker:
    def __init__(self, config, password, status_queue):
        self.config = config
        self.password = password
        self.status_queue = status_queue
        self.pp_manager = PayPeriodManager(config.get('Folders', 'pay_period_schedule_csv'))
        
        # Set up logging
        log_file = self.config.get('Folders', 'log_file')
        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s',
                            handlers=[logging.FileHandler(log_file), logging.StreamHandler()])

    def log_and_update(self, message, level="info"):
        """Logs a message and sends it to the GUI queue."""
        if level == "info":
            logging.info(message)
        elif level == "warning":
            logging.warning(message)
        elif level == "error":
            logging.error(message)
        self.status_queue.put(message)
        time.sleep(0.1) # Prevents GUI from lagging on rapid updates

    def find_header_row(self, df, expected_headers, threshold=80):
        """Finds the header row in a dataframe using fuzzy matching."""
        for i in range(min(20, len(df))):
            row_str = ' '.join(map(str, df.iloc[i].tolist()))
            combined_score = 0
            for header in expected_headers:
                score, _, _ = process.extractOne(header, [row_str], scorer=fuzz.partial_ratio)
                combined_score += score
            
            avg_score = combined_score / len(expected_headers)
            if avg_score > threshold:
                return i
        return None

    def process_attachments(self, found_reports):
        """Processes and merges reports for each completed pay period."""
        self.log_and_update("Starting file processing and merging...")
        for pp_id, reports in found_reports.items():
            if len(reports) == 2:
                self.log_and_update(f"Found both reports for Pay Period {pp_id}. Starting merge.")
                
                try:
                    self._merge_reports(pp_id, reports)
                except Exception as e:
                    self.log_and_update(f"ERROR merging reports for PP {pp_id}: {e}", "error")
            else:
                self.log_and_update(f"Skipping merge for PP {pp_id}: did not find both reports.", "warning")

    def _merge_reports(self, pp_id, reports):
        """Merges two Excel reports into a single consolidated file."""
        report_paths = {r['type']: r['path'] for r in reports}
        pp_folder = Path(report_paths[1]).parent.parent # The main PP folder
        
        headers1 = ['Super Office', 'Division', 'Time Sheet: Owner Name', 'Sum of Pay Period Number', 'Sum of Hours']
        headers2 = ['Super Office', 'Time Sheet: Owner Name', 'FDA Employee Number', 'Sum of Pay Period Number', 'Sum of Hours']

        df1_raw = pd.read_excel(report_paths[1], engine='openpyxl')
        header_row1 = self.find_header_row(df1_raw, headers1)
        if header_row1 is None:
            self.log_and_update(f"Could not find header row in {os.path.basename(report_paths[1])}", "error")
            return
        df1 = pd.read_excel(report_paths[1], engine='openpyxl', header=header_row1).rename(columns=lambda x: x.strip())
        
        df2_raw = pd.read_excel(report_paths[2], engine='openpyxl')
        header_row2 = self.find_header_row(df2_raw, headers2)
        if header_row2 is None:
            self.log_and_update(f"Could not find header row in {os.path.basename(report_paths[2])}", "error")
            return
        df2 = pd.read_excel(report_paths[2], engine='openpyxl', header=header_row2).rename(columns=lambda x: x.strip())

        merged_data = []
        mismatched_log = []
        df2_names = df2['Time Sheet: Owner Name'].tolist()
        name_match_threshold = int(self.config.get('FuzzyLogic', 'name_match_threshold'))

        for index, row1 in df1.iterrows():
            name1 = row1['Time Sheet: Owner Name']
            best_match = process.extractOne(name1, df2_names, score_cutoff=name_match_threshold)
            
            new_row = row1.to_dict()
            if best_match:
                match_name, score, _ = best_match
                row2 = df2[df2['Time Sheet: Owner Name'] == match_name].iloc[0]
                new_row['FDA Employee Number'] = row2.get('FDA Employee Number', 'N/A')
            else:
                new_row['FDA Employee Number'] = 'NOT FOUND'
                mismatched_log.append(f"No matching Employee ID found for: '{name1}'")
            merged_data.append(new_row)

        merged_df = pd.DataFrame(merged_data)
        final_columns = ['Super Office', 'Division', 'Time Sheet: Owner Name', 'FDA Employee Number', 'Sum of Pay Period Number', 'Sum of Hours', 'Sum of Tour of Duty Hours', 'Compliance % by User/ Div/SuperOffice']
        merged_df = merged_df[[col for col in final_columns if col in merged_df.columns]]
        
        output_filename = pp_folder / f"PP{pp_id} - Combined Report.xlsx"
        merged_df.to_excel(output_filename, index=False)
        self.log_and_update(f"Successfully created combined report: {output_filename}")
        
        if mismatched_log:
            log_filename = pp_folder / f"PP{pp_id} - Mismatch_Log.txt"
            with open(log_filename, 'w') as f:
                f.write('\n'.join(mismatched_log))
            self.log_and_update(f"Mismatch log created: {log_filename}")

    def run(self):
        """Main execution function to run the automation workflow."""
        try:
            email_addr = self.config.get('Email', 'address')
            server = self.config.get('Email', 'server')
            credentials = Credentials(email_addr, self.password)
            config = Configuration(server=server, credentials=credentials)
            account = Account(primary_smtp_address=email_addr, config=config, autodiscover=False, access_type=DELEGATE)
            self.log_and_update(f"Successfully connected to {email_addr}.")

            target_subjects = {1: self.config.get('Reports', 'report1_subject'), 2: self.config.get('Reports', 'report2_subject')}
            subject_threshold = int(self.config.get('FuzzyLogic', 'subject_match_threshold'))
            root_path = Path(self.config.get('Folders', 'root_path'))
            found_reports = {}

            self.log_and_update("Searching inbox for report emails...")
            for item in account.inbox.all().order_by('-datetime_received')[:200]:
                for report_type, target_subject in target_subjects.items():
                    score = fuzz.partial_ratio(target_subject.lower(), item.subject.lower())
                    if score >= subject_threshold:
                        self.log_and_update(f"Found potential match (Score: {score}%): '{item.subject}'")
                        pp_row, error = self.pp_manager.find_pp_for_date(item.datetime_received.astimezone())
                        if error:
                            self.log_and_update(f"Skipping email '{item.subject}': {error}", "warning")
                            continue
                        
                        pp_id = pp_row['Pay Period']
                        pp_folder_name, subfolder = self.pp_manager.get_folder_details(pp_row, item.datetime_received.astimezone())
                        save_path = root_path / pp_folder_name / subfolder
                        save_path.mkdir(parents=True, exist_ok=True)
                        
                        for attachment in item.attachments:
                            if attachment.name.lower().endswith(('.xlsx', '.xls')):
                                filepath = save_path / attachment.name
                                with open(filepath, 'wb') as f:
                                    f.write(attachment.content)
                                self.log_and_update(f"Downloaded '{attachment.name}' to '{subfolder}'.")
                                if pp_id not in found_reports:
                                    found_reports[pp_id] = []
                                if not any(r['type'] == report_type for r in found_reports[pp_id]):
                                    found_reports[pp_id].append({'type': report_type, 'path': filepath})
                                break
                        break
            
            if not found_reports:
                self.log_and_update("No new reports found in recent emails.")
            else:
                self.process_attachments(found_reports)
            self.log_and_update("Automation run complete.", "info")
        except Exception as e:
            self.log_and_update(f"An unexpected error occurred: {e}", "error")
        finally:
            self.status_queue.put("DONE")
app.py
Python

import customtkinter as ctk
from tkinter import filedialog, messagebox, simpledialog
import os
import queue
import threading

from config_manager import ConfigManager
from automation_worker import AutomationWorker

class ToolTip(ctk.CTkToplevel):
    def __init__(self, widget, text):
        super().__init__(widget)
        self.widget = widget
        self.text = text
        self.withdraw()
        self.overrideredirect(True)
        self.label = ctk.CTkLabel(self, text=self.text, corner_radius=5, fg_color="#404040", text_color="white", padx=10, pady=5)
        self.label.pack()
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event):
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.geometry(f"+{x}+{y}")
        self.deiconify()

    def hide_tip(self, event):
        self.withdraw()

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Pay Period Report Automation")
        self.geometry("800x650")

        self.config = ConfigManager()
        self.status_queue = queue.Queue()
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme(self.config.get('Application', 'theme').lower())
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.settings_frame = ctk.CTkFrame(self)
        self.settings_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.settings_frame.grid_columnconfigure(1, weight=1)
        self.create_settings_widgets()
        
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.create_log_widgets()
        
        self.load_settings()

    def create_settings_widgets(self):
        ctk.CTkLabel(self.settings_frame, text="Email Address:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.email_entry = ctk.CTkEntry(self.settings_frame, width=300)
        self.email_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ToolTip(self.email_entry, "The email address of the inbox to monitor.")
        
        ctk.CTkLabel(self.settings_frame, text="Root Folder Path:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.root_path_entry = ctk.CTkEntry(self.settings_frame, width=300)
        self.root_path_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        self.browse_button = ctk.CTkButton(self.settings_frame, text="Browse...", command=self.browse_folder)
        self.browse_button.grid(row=1, column=2, padx=10, pady=5)
        ToolTip(self.root_path_entry, "The main folder where all Pay Period subfolders will be created.")

        ctk.CTkLabel(self.settings_frame, text="Pay Period File:").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.pp_file_entry = ctk.CTkEntry(self.settings_frame, width=300)
        self.pp_file_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")
        self.browse_pp_button = ctk.CTkButton(self.settings_frame, text="Browse...", command=self.browse_pp_file)
        self.browse_pp_button.grid(row=2, column=2, padx=10, pady=5)
        ToolTip(self.pp_file_entry, "The CSV file containing the pay period schedule.")

        self.save_button = ctk.CTkButton(self.settings_frame, text="Save Settings", command=self.save_settings)
        self.save_button.grid(row=3, column=1, columnspan=2, padx=10, pady=10, sticky="e")

    def create_log_widgets(self):
        self.start_button = ctk.CTkButton(self.log_frame, text="Start Automation", command=self.start_automation, height=40, font=ctk.CTkFont(size=14, weight="bold"))
        self.start_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        ToolTip(self.start_button, "Connects to the email inbox and starts processing reports.")
        self.log_textbox = ctk.CTkTextbox(self.log_frame, state="disabled", wrap="word")
        self.log_textbox.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

    def log_message(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        self.log_textbox.configure(state="disabled")
        self.log_textbox.see("end")

    def browse_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path: self.root_path_entry.delete(0, "end"); self.root_path_entry.insert(0, folder_path)
    
    def browse_pp_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path: self.pp_file_entry.delete(0, "end"); self.pp_file_entry.insert(0, file_path)

    def load_settings(self):
        self.email_entry.insert(0, self.config.get('Email', 'address'))
        self.root_path_entry.insert(0, self.config.get('Folders', 'root_path'))
        self.pp_file_entry.insert(0, self.config.get('Folders', 'pay_period_schedule_csv'))

    def save_settings(self):
        self.config.set('Email', 'address', self.email_entry.get())
        self.config.set('Folders', 'root_path', self.root_path_entry.get())
        self.config.set('Folders', 'pay_period_schedule_csv', self.pp_file_entry.get())
        self.config.save_config()
        messagebox.showinfo("Success", "Settings have been saved successfully!")

    def start_automation(self):
        self.save_settings()
        password = simpledialog.askstring("Password", "Please enter the email password:", show='*')
        if not password:
            messagebox.showwarning("Cancelled", "Password not provided. Automation cancelled.")
            return
        self.log_textbox.configure(state="normal"); self.log_textbox.delete("1.0", "end"); self.log_textbox.configure(state="disabled")
        self.start_button.configure(state="disabled", text="Running...")
        self.worker = AutomationWorker(self.config, password, self.status_queue)
        self.thread = threading.Thread(target=self.worker.run, daemon=True)
        self.thread.start()
        self.after(100, self.check_queue)

    def check_queue(self):
        try:
            message = self.status_queue.get_nowait()
            if message == "DONE":
                self.start_button.configure(state="normal", text="Start Automation")
            else:
                self.log_message(message)
            self.after(100, self.check_queue)
        except queue.Empty:
            if self.thread.is_alive():
                self.after(100, self.check_queue)
            else: # Thread finished but no DONE message might indicate a crash
                self.start_button.configure(state="normal", text="Start Automation")


if __name__ == "__main__":
    if not os.path.exists('config.ini'):
         messagebox.showerror("Error", "config.ini not found! Please create it before running the application.")
    else:
        from datetime import datetime
        app = App()
        app.mainloop()

🚀 How to Set Up and Run
Prerequisites: Ensure Python 3.8+ is installed.

Clone Repository: Download all the files from this repository into a single folder.

Install Dependencies: Open a terminal or command prompt in the project folder and run:

Bash

pip install -r requirements.txt
Configure:

Open config.ini and verify the address and root_path are correct.

Open pay_periods.csv and ensure the schedule is up-to-date.

Run the Application:

Bash

python app.py
📦 How to Package for Distribution
To create a single .exe file for Windows that can be shared with users who do not have Python installed:

Install PyInstaller:

Bash

pip install pyinstaller
Run the Build Command: In the project's terminal, execute the following command:

Bash

pyinstaller --name "PayPeriodAutomation" --onefile --windowed --add-data "pay_periods.csv;." --add-data "config.ini;." app.py
Find the Executable: The final PayPeriodAutomation.exe file will be located in the newly created dist folder. Distribute this .exe file along with the config.ini and pay_periods.csv files.
