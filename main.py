import shutil
import pandas as pd
from datetime import datetime, timedelta
import schedule
import time
import pywhatkit

import tkinter as tk
from tkinter import filedialog
import json
import os
import sys

CONFIG_FILE = "config.json"


def get_excel_file_path():
    # Check if a file was already selected previously and saved in the config
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                saved_path = config.get("excel_path")
                # Ensure the file still exists at the saved path
                if saved_path and os.path.exists(saved_path):
                    return saved_path
        except Exception as e:
            print(f"Error reading config: {e}")

    # If there is no config file or the file was deleted, prompt the user
    root = tk.Tk()
    root.withdraw()  # Hide the main empty tkinter window

    # Open the file selection dialog
    file_path = filedialog.askopenfilename(
        title="Select the Excel file for badge tracking",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if file_path:
        # Save the selected path for future runs
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump({"excel_path": file_path}, f)
        return file_path
    else:
        print("No Excel file selected. The program will exit.")
        sys.exit()  # Terminate the script if no file is chosen

class ExcelBadgeChecker:
    """
    Responsible for reading data from the Excel file and filtering employees
    who haven't returned their temporary badge.
    """

    def __init__(self, file_path, date_formats=None):
        self.file_path = file_path
        if date_formats is None:
            self.date_formats = [
                "%d.%m.%y",
                "%d.%m.%Y",
                "%d/%m/%y",
                "%d/%m/%Y",
                "%d-%m-%y",
                "%d-%m-%Y",
                "%d %m %y",
                "%d %m %Y",
                "%d.%m",
                "%d/%m"
            ]
        else:
            self.date_formats = date_formats

    def get_unreturned_yesterday(self):
        yesterday = datetime.now() - timedelta(days=1)
        # Fix 1: Append the suffix to the end of the path to avoid invalid paths
        temp_file_path = self.file_path + '.temp.xlsx'
        target_sheet = None
        # Fix 3: Initialize df to None to prevent UnboundLocalError
        df = None

        try:
            # Copy the file to the temporary path (for case the original file is locked)
            shutil.copy2(self.file_path, temp_file_path)
            with pd.ExcelFile(temp_file_path) as xl_file:
                existing_sheet_names = xl_file.sheet_names

                for fmt in self.date_formats:
                    possible_name = yesterday.strftime(fmt)
                    if possible_name in existing_sheet_names:
                        target_sheet = possible_name
                        break

                if not target_sheet:
                    print("Could not find a sheet for yesterday in any of the known formats.")
                    return []

                # Read the data from the matched sheet
                df = pd.read_excel(xl_file, sheet_name=target_sheet)

        except Exception as e:
            print(f"Error processing the Excel file: {e}")
            return []

        finally:
            # Make sure to delete the temporary file when done
            if os.path.exists(temp_file_path):
                try:
                    os.remove(temp_file_path)
                except Exception as e:
                    # Fix 2: Only print a warning, do not stop the function execution with a return statement
                    print(f"Warning: Could not delete temporary file {temp_file_path}: {e}")

        # If df is somehow empty or not created, stop here
        if df is None or df.empty:
            return []

        df.columns = df.columns.str.strip()

        status_cols = ['הוחזר/לא הוחזר', 'הוחזר / לא הוחזר', 'הוחזר', 'סטטוס']
        phone_cols = ['מס טלפון', 'טלפון', 'טלפון נייד', 'מס פלאפון', 'פלאפון']
        name_cols = ['שם מלא', 'שם']

        actual_status_col = next((col for col in status_cols if col in df.columns), None)
        actual_phone_col = next((col for col in phone_cols if col in df.columns), None)
        actual_name_col = next((col for col in name_cols if col in df.columns), None)

        if not actual_status_col:
            print(f"Could not find a status column. Looked for: {status_cols}")
            return []

        # Filter: Not 'כן' and not 'הוחזר'
        status_col = df[actual_status_col].fillna('').astype(str).str.strip()
        unreturned_mask = ~status_col.isin(['כן', 'הוחזר'])

        unreturned_df = df[unreturned_mask]

        unreturned_employees = []
        for index, row in unreturned_df.iterrows():
            # Extract phone and name based on the found columns
            phone = str(row.get(actual_phone_col, '')) if actual_phone_col else ''
            phone = phone.strip()

            name = row.get(actual_name_col, 'Employee') if actual_name_col else 'Employee'

            # Format phone number for WhatsApp (adding +972 for Israel)
            phone = phone.replace("-", "").replace(" ", "")

            if phone and phone.lower() != 'nan':
                # If it ends with '.0', remove the extension
                phone = phone.split('.')[0]
                # If it starts with 0, replace with +972
                if phone.startswith('0'):
                    phone = '+972' + phone[1:]
                elif phone.startswith('5'):
                    phone = '+972' + phone
                # If it doesn't have a plus, assume it needs a plus
                elif not phone.startswith('+'):
                    phone = '+' + phone

                unreturned_employees.append({
                    'name': name,
                    'phone': phone,
                    'badge_number': row.get('תג', 'Unknown')
                })

        return unreturned_employees


class WhatsAppNotifier:
    """
    Responsible for receiving the list of employees and sending WhatsApp notifications.
    """

    def __init__(self):
        pass  # No special authentication needed, it uses the active WhatsApp Web session

    def send_notifications(self, unreturned_list):
        if not unreturned_list:
            print("No employees forgot their badges yesterday. Great job!")
            return

        for employee in unreturned_list:
            name = employee['name']
            phone = employee['phone']
            badge = employee['badge_number']

            message = (
                f"שלום {name},\n"
                f"נראה כי שכחת להחזיר את תג מספר {badge} שלקחת אתמול. "
                f"אנא דאג/י להחזיר אותו בהקדם.\n"
                f"תודה ויום טוב!"
            )

            self._send_single_whatsapp(phone, message)

    def _send_single_whatsapp(self, phone, message):
        print(f"Preparing to send WhatsApp message to {phone}...")
        try:
            # sendwhatmsg_instantly parameters:
            # phone_no, message, wait_time (seconds), tab_close (True/False), close_time (seconds)
            # We wait 15 seconds for the browser to open and load the chat, then close the tab after 3 seconds.
            pywhatkit.sendwhatmsg_instantly(phone, message, wait_time=15, tab_close=True, close_time=3)
            print(f"Successfully sent to: {phone}")

            # Wait a few seconds between messages to avoid WhatsApp spam filters
            time.sleep(5)
        except Exception as e:
            print(f"Failed to send WhatsApp to {phone}: {e}")


def main_job():
    print(f"Starting scan... [{datetime.now().strftime('%d/%m/%Y %H:%M')}]")

    excel_path = "tags.xlsx"
    excel_path = get_excel_file_path()

    checker = ExcelBadgeChecker(file_path=excel_path, date_formats=None)
    notifier = WhatsAppNotifier()

    unreturned = checker.get_unreturned_yesterday()
    notifier.send_notifications(unreturned)


main_job()

schedule.every().day.at("08:00").do(main_job)

if __name__ == "__main__":
    print("WhatsApp badge tracking system is running and waiting for 08:00...")

    # You MUST have WhatsApp Web logged in on your default browser for this to work!
    # Uncomment to test immediately:
    # main_job()


    exit(1)

    while True:
        schedule.run_pending()
        time.sleep(60)