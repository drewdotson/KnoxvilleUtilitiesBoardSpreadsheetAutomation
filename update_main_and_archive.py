# The purpose of this file is to update the main installation job and 90 day archive sheets. Each sheet is edited based off
# of which rows in the main data sheet contain jobs that have been completed for at least 90 days. If a job meets the criteria,
# then the row containing the data is added to the archive sheet and then subsequentially removed from the main sheet.
# Since Pandas does not support advanced Excel formatting, Openpyxl will be used in another file to restore the formatting of
# each sheet.
# NOTE: Since this file is so short, every piece of functionality related to specifically updating the data in the main and 
#       archive sheets is stored in this file only.

import pandas as pd
from datetime import datetime, timedelta
from time import sleep

# FUNCTIONS 👇 ---------------------------------------------------------------------------------------------------#

# Function used for formatting neccessary columns of a specified sheet for modification. It uses a sheet name as input.
def format_sheet(sheet):

    # Format each column that contains dates to clear any cell that does not contain a date value.
    sheet['Drop Installation Date'] = pd.to_datetime(sheet['Drop Installation Date'], errors='coerce')
    sheet['811 Called in Date'] = pd.to_datetime(sheet['811 Called in Date'], errors='coerce')
    sheet['Begin Work On'] = pd.to_datetime(sheet['Begin Work On'], errors='coerce')
    sheet['811 Marked Date'] = pd.to_datetime(sheet['811 Marked Date'], errors='coerce')
    sheet['Ticket Expiration Date'] = pd.to_datetime(sheet['Ticket Expiration Date'], errors='coerce')
    sheet['Job Completed Date'] = pd.to_datetime(sheet['Job Completed Date'], errors='coerce')

    # # Drop columns that contain legend info for cell fill colors. This will be restored once OpenpyXL is used.
    # sheet.drop(sheet.columns[[14, 15]], axis=1, inplace=True)


# Function used to check if a specified installation job has been completed. It uses a row number as input.
def status_check(row):

    # The value of the "Status" column in the specified row is saved to a variable
    status = row[14]

    # If the value's data type is a string and the status is "Completed" (case-insensitive), then the function returns True.
    if type(status) == str and status.lower() == "completed":
        return True

    # If the value does not match at least one of the two conditions, the function returns False.
    else:
        return False


# Function used to check if a specified installation job has been completed for at least 90 days. It uses a row number and
# the current date on the computer's internal clock as input.
def time_check(row, current_date):
    
    # The value of the "Job Completed Date" column in the specified row is saved as a variable.
    completion_date = row[6]

    # If the value is not null, then the elapsed time for the installation job is calculated by taking
    # the difference of the computer's current date and the completion date of the installation job.
    if pd.isnull(completion_date) != True:
            difference = current_date - completion_date.date()

            # If the difference is greater than or equal to 90 days, the function returns True
            if difference >= timedelta(days = 90):
                return True
            
            # If the difference is less than 90 days, the Functions returns False.
            else:
                return False


# This function creates two variables which will each be their own DataFrame copy of the main sheet.
# One of the DataFrames will remove every row that will be imported to the archive sheet, while the other will remove 
# every row that will be kept in the main sheet. The two DataFrames are then returned from the function.
def sort_data(main_sheet, import_list, keep_list):
    main_import = main_sheet.drop(index=keep_list)
    main_update = main_sheet.drop(index=import_list)
    return main_import, main_update

# MAIN FUNCTIONALITY 👇 -----------------------------------------------------------------------------------------#

def update_main_and_archive(main_sheet, archive_sheet):

    # Initialize the main and archive sheets 
    format_sheet(main_sheet)
    format_sheet(archive_sheet)

    # Gather the current date which will be used in the time check funtion. This is done here and not in the
    # function so the program does not have to retrieve the date numerous times.
    current_date = datetime.now().date()

    # Gather the first and last row indexes in the main sheet and save the values as variables
    main_first_row = main_sheet.first_valid_index()
    main_last_row = main_sheet.last_valid_index()

    # Create two lists that will be used the hold the row indexes of every row that will be either
    # imported to the archive sheet or kept in the main sheet. This will allow the program to import and delete
    # every necessary row all at once.
    rows_to_import = []
    rows_to_keep = []

    print(f"\nAmount of rows in main sheet pre-update: {main_sheet.last_valid_index() + 2}")
    print(f"Amount of rows in archive sheet pre-update: {archive_sheet.last_valid_index() + 2}")
    sleep(2)
    
    # The program iterates over every row in the main sheet and performs the time and status checks on each row.
    for row in main_sheet[main_first_row:main_last_row + 1].itertuples():
        check1 = status_check(row)
        check2 = time_check(row, current_date)
        
        # If the row fails either of the two checks, then its index is added to the keep list.
        if check1 == False or check2 == False:
            rows_to_keep.append(row.Index)

        # If the row passes both checks, then its index is added to the import list.
        else:
            rows_to_import.append(row.Index)

    print(f"\nAmount of rows to keep in main sheet: {len(rows_to_keep)}")
    print(f"Amount of rows to import to archive sheet: {len(rows_to_import)}")
    sleep(2)

    # The main sheet is split into two DataFrames, one which contains the rows to keep, and one which contains which
    # rows to import.
    main_import, main_update = sort_data(main_sheet, rows_to_import, rows_to_keep)

    archive_update = pd.concat([archive_sheet, main_import], ignore_index=True)

    print(f"\nAmount of rows in main sheet post-update: {main_last_row - len(rows_to_import) + 2}")
    print(f"Amount of rows in archive sheet post-update: {archive_update.last_valid_index() + 2}")

    # The main function returns the updated main and archive sheets. 
    return main_update, archive_update
