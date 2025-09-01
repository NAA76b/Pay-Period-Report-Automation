# Pay Period Report Automation

## Objective
Automates processing of recurring pay period reports using a hybrid Outlook VBA and Python workflow.

## Architecture
- **Outlook VBA Macro**: Runs inside Outlook and watches incoming emails. It forwards messages, saves specific Excel attachments to a local landing zone, and prefixes each file with the current date.
- **Python Script (`process_reports.py`)**: Run manually from the command line. It scans the landing zone for the two required reports, uses `pay_periods.csv` to determine the correct pay period, creates a smartly named destination folder, merges the data, writes a combined report, logs unmatched names, and cleans up the processed files.

## File Structure
```
Pay-Period-Report-Automation/
├── process_reports.py     # Main Python script for merging reports
├── requirements.txt       # Required Python libraries
└── pay_periods.csv        # Pay period schedule
```

## Setup
1. Install Python 3.8 or later.
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Ensure the Outlook VBA macro is configured to save report attachments to the landing zone specified in `process_reports.py`.

## Usage
Run the processing script after the Outlook macro has downloaded the two Excel reports:
```bash
python process_reports.py
```
The script generates a combined report inside the destination folder for the pay period and removes the processed files from the landing zone.

