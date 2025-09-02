# Pay Period Report Automation
Automates processing of recurring pay period reports using an Outlook VBA macro and a companion Python script.

## How to Run
1. Ensure the Outlook VBA macro is active and saving attachments to `C:\Users\Nathan.Allen\OneDrive - FDA\Desktop\PayPeriodAttachments`.
2. Install dependencies with `pip install -r requirements.txt`.
3. Run `python process_reports.py` from this directory. The script merges the two report files and writes a consolidated workbook to `C:\Users\Nathan.Allen\OneDrive - FDA\Desktop\CBER ITR Participation Clean Up`.
