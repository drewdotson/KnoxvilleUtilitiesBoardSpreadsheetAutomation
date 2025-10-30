
# The purpose of this file is to update the main installation job and 90 day archive sheets. Each sheet is edited based off of which rows in the main 
# data sheet contain jobs that have been completed for at least 90 days. If a job meets the criteria, then the row containing the data is added to the 
# archive sheet and then subsequentially removed from the main sheet.

# Since Pandas does not support advanced Excel formatting, Openpyxl will be used in another file to restore the formatting of both sheets.

# NOTE: Since this file is so short, every piece of functionality related to specifically updating the data in the main and archive sheets is stored 
#       in this file only.

import pandas as pd
from datetime import datetime, timedelta
from time import sleep

# FUNCTIONS 👇 ---------------------------------------------------------------------------------------------------------------------------------------#

# Function used for formatting date-containing columns of the main sheet for modification. It uses the main sheet as input.
def format_main_sheet_dates(main_sheet):

    # Format each column that contains dates to date values. However, if a specific cell contains a value that can not be converted to a date 
    # value (such as a ? or a note) then then it is ignored. This is due to some jobs that have important notes in these columns prior to completion.
    main_sheet['Drop Installation Date'] = main_sheet.apply(lambda row: pd.to_datetime(row['Drop Installation Date'], errors='ignore'), axis=1)
    main_sheet['811 Called in Date'] = main_sheet.apply(lambda row: pd.to_datetime(row['811 Called in Date'], errors='ignore'), axis=1)
    main_sheet['Begin Work On'] = main_sheet.apply(lambda row: pd.to_datetime(row['Begin Work On'], errors='ignore'), axis=1)
    main_sheet['811 Marked Date'] = main_sheet.apply(lambda row: pd.to_datetime(row['811 Marked Date'], errors='ignore'), axis=1)
    main_sheet['Ticket Expiration Date'] = main_sheet.apply(lambda row: pd.to_datetime(row['Ticket Expiration Date'], errors='ignore'), axis=1)
    main_sheet['Job Completed Date'] = main_sheet.apply(lambda row: pd.to_datetime(row['Job Completed Date'], errors='ignore'), axis=1) 


# Function used for formatting date-containing columns of the archive sheet of the sheet for modification. It uses the archive sheet as input.
def format_archive_sheet_dates(archive_sheet):

    # Format each column that contains dates to date values. However, if a specific cell contains a value that can not be converted to a date value 
    # then its value is cleared out.
    archive_sheet['Drop Installation Date'] = archive_sheet.apply(lambda row: pd.to_datetime(row['Drop Installation Date'], errors='coerce'), axis=1)
    archive_sheet['811 Called in Date'] = archive_sheet.apply(lambda row: pd.to_datetime(row['811 Called in Date'], errors='coerce'), axis=1)
    archive_sheet['Begin Work On'] = archive_sheet.apply(lambda row: pd.to_datetime(row['Begin Work On'], errors='coerce'), axis=1)
    archive_sheet['811 Marked Date'] = archive_sheet.apply(lambda row: pd.to_datetime(row['811 Marked Date'], errors='coerce'), axis=1)
    archive_sheet['Ticket Expiration Date'] = archive_sheet.apply(lambda row: pd.to_datetime(row['Ticket Expiration Date'], errors='coerce'), axis=1)
    archive_sheet['Job Completed Date'] = archive_sheet.apply(lambda row: pd.to_datetime(row['Job Completed Date'], errors='coerce'), axis=1)  


# Function used to check if a specified installation job has been completed for at least 90 days. It uses a row number and the current date on the
# computer's internal clock as input.
def completion_check(row, current_date):
    
    # The value of the "Status" column in the specified row is saved to a variable
    status = row[14]
    # The value of the "Job Completed Date" column in the specified row is saved as a variable.
    completion_date = row[6]

    # The "Status" column's value is checked to see if the value equals the string "completed" (case-insensitive).
    if type(status) == str and status.lower() == "completed":

        # If the row has a completed job status, the "Job Completed Date" column's data type is checked to see if it is a pandas timestamp or 
        # datetime object.
        if type(completion_date) == pd._libs.tslibs.timestamps.Timestamp or type(completion_date) == datetime:
                
                # If the column's data type is one of the two, then the elapsed time is calculated and saved to a variable.
                difference = current_date - completion_date.date()

                # If the difference is greater than or equal to 90 days, the function returns True
                if difference >= timedelta(days = 90):
                    return True
                
                # If the difference is less than 90 days, the Functions returns False.
                else:
                    return False
                
    # If the "Status" column's value does not equal the string "completed" (case-insensitive), the funtion returns False.
    else:
                    return False


# This function creates two variables which will each be their own DataFrame copy of the main sheet.
# One of the DataFrames will remove every row that will be imported to the archive sheet, while the other will remove 
# every row that will be kept in the main sheet. The two DataFrames are then returned from the function.
def sort_data(main_sheet, import_list, keep_list):
    main_import = main_sheet.drop(index=keep_list)
    main_update = main_sheet.drop(index=import_list)
    return main_import, main_update

# MAIN FUNCTIONALITY 👇 ------------------------------------------------------------------------------------------------------------------------------#

def update_main_and_archive(main_sheet, archive_sheet):

    # Correct the main sheet's date-contatining column formats. 
    format_main_sheet_dates(main_sheet)

    # Gather the current date which will be used in the completion check funtion. This is done here and not in the function so the program does not 
    # have to retrieve the date numerous times.
    current_date = datetime.now().date()

    # Gather the first and last row indexes in the main sheet and save the values as variables
    main_first_row = main_sheet.first_valid_index()
    main_last_row = main_sheet.last_valid_index()

    # Create two lists that will be used the hold the row indexes of every row that will be either imported to the archive sheet or kept in the main 
    # sheet. This will allow the program to import and delete every necessary row all at once.
    rows_to_import = []
    rows_to_keep = []

    print(f"\nAmount of rows in main sheet pre-update: {main_last_row + 2}")
    print(f"Amount of rows in archive sheet pre-update: {archive_sheet.last_valid_index() + 2}")
    sleep(1)
    
    # The program iterates over every row in the main sheet and performs the completion check on each row.
    for row in main_sheet[main_first_row:main_last_row + 1].itertuples():
        check = completion_check(row, current_date)
        
        # If the row fails the check, then its index is added to the keep list.
        if check == False:
            rows_to_keep.append(row.Index)

        # If the row passes the check, then its index is added to the import list.
        else:
            rows_to_import.append(row.Index)

    print(f"\nAmount of rows to keep in main sheet: {len(rows_to_keep) + 1}")
    print(f"Amount of rows to import to archive sheet: {len(rows_to_import)}")
    sleep(1)

    # The main sheet is split into two DataFrames, one which contains the rows to keep, and one which contains whichrows to import.
    main_import, main_update = sort_data(main_sheet, rows_to_import, rows_to_keep)

    # The updated archive sheet is saved to a new DataFrame.
    archive_update = pd.concat([archive_sheet, main_import], ignore_index=True)

    # Correct the archive sheet's date-contatining column formats.
    format_archive_sheet_dates(archive_update)

    print(f"\nAmount of rows in main sheet post-update: {main_last_row - len(rows_to_import) + 2}")
    print(f"Amount of rows in archive sheet post-update: {archive_update.last_valid_index() + 2}")
    sleep(1)

    # The main function returns the updated main and archive sheets. 
    return main_update, archive_update
