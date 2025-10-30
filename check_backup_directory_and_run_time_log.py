
# The purpose of this file is to verify if a backup directory and run time log for the current month is present in the appropriate 
# directories.

from datetime import datetime
import os

# Function used for assigning a month key to the current date. It uses the current date in mm-dd-YYYY format as input.
def create_month_key(current_date):

    # A dictionary is created to assign month identifiers to a month string. This is used to create a month title for the current 
    # month's backup and run time log directory.
    month_dict = {
                  "01": "Jan",
                  "02": "Feb",
                  "03": "Mar",
                  "04": "Apr",
                  "05": "May",
                  "06": "Jun",
                  "07": "Jul",
                  "08": "Aug",
                  "09": "Sep",
                  "10": "Oct",
                  "11": "Nov",
                  "12": "Dec"
                }
    
    # The current date is split into three components.
    month, day, year = current_date.split("-")

    # The function checks if the month component is a key in the month dictionary. (This is done to prevent an error if somehow an 
    # invalid date input is given.)
    if month in month_dict:

        # The month key is created by concatenating the current month's key value with the year component. The month key is then returned.
        month_key = month_dict[month] + " " + year
        return month_key


# Function used for verifying if the current month already has a backup directory and run time log present in the appropriate directories.
def check_backup_directory_and_run_time_log():

    # The current date is gathered and saved to a variable in mm-dd-YYYY format.
    current_date = str(datetime.today().strftime('%m-%d-%Y'))

    # The month key is generated from the current date.
    month_key = create_month_key(current_date)

    # The function checks if the current month already has a backup directory present. If not, then one is created.
    if os.path.isdir(f"Backups/{month_key}"):
        pass
    else:
        os.mkdir(f"Backups/{month_key}")
    
    # The function checks if the current month already has a run time log present. If not, then one is created.
    if os.path.isdir(f"Run Times/{month_key}.txt"):
        pass
    else:
        with open(f'Run Times\\{month_key}.txt', 'w') as f:
            f.write(f"Run Time Archive for {month_key}")

    # The current date and generated month key are generated.
    return current_date, month_key

