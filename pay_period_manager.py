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
