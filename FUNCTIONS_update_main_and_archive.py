
# The purpose of this file is to contain the functions used for the update_main_and_archive file.

import pandas as pd
from datetime import datetime, timedelta


# FUNCTIONS 👇 ---------------------------------------------------------------------------------------------------------------------------------------#

# Function used for formatting date-containing columns of the main sheet for modification. It uses the main sheet as input.
def format_main_sheet_dates(main_sheet):

    # Format each column that contains dates to date values. However, if a specific cell contains a value that can not be converted to a date 
    # value (such as a ? or a note) then then it is ignored. This is due to some jobs that have important notes in these columns prior to completion.
    main_sheet['Drop Installation Date'] = main_sheet.apply(lambda row: pd.to_datetime(row['Drop Installation Date'], errors='ignore'), axis=1)
    main_sheet['811 Called in Date'] = main_sheet.apply(lambda row: pd.to_datetime(row['811 Called in Date'], errors='ignore'), axis=1)
    main_sheet['Begin Work On'] = main_sheet.apply(lambda row: pd.to_datetime(row['Begin Work On'], errors='ignore'), axis=1)
    main_sheet['811 Marked Date'] = main_sheet.apply(lambda row: pd.to_datetime(row['811 Marked Date'], errors='ignore'), axis=1)
    main_sheet['Update By'] = main_sheet.apply(lambda row: pd.to_datetime(row['Update By'], errors='ignore'), axis=1)
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
    archive_sheet['Update By'] = archive_sheet.apply(lambda row: pd.to_datetime(row['Update By'], errors='coerce'), axis=1)
    archive_sheet['Ticket Expiration Date'] = archive_sheet.apply(lambda row: pd.to_datetime(row['Ticket Expiration Date'], errors='coerce'), axis=1)
    archive_sheet['Job Completed Date'] = archive_sheet.apply(lambda row: pd.to_datetime(row['Job Completed Date'], errors='coerce'), axis=1)  


# Function used to check if a specified installation job has been completed for at least 90 days. It uses a row number and the current date on the
# computer's internal clock as input.
def completion_check(row, current_date):
    
    # The value of the "Status" column in the specified row is saved to a variable
    status = row[15]
    # The value of the "Job Completed Date" column in the specified row is saved as a variable.
    completion_date = row[7]

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
