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
