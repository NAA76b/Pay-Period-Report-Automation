# Pay Period Report Automation
Automates processing of recurring pay period reports using an Outlook VBA macro and a Python script.

## Configuration
Edit `config.ini` to set the landing zone, output folder, and schedule CSV. If a configured path is missing, the script falls back to a similarly named folder on your Desktop.

## How to Run
1. Ensure the VBA macro is active in Outlook and saving attachments to the landing zone.
2. Install dependencies: `pip install -r requirements.txt`
3. Run the script: `python process_reports.py`
