
# The purpose of this file is to update the main installation job and 90 day archive sheets. Each sheet is edited based off of which rows in the main 
# data sheet contain jobs that have been completed for at least 90 days. If a job meets the criteria, then the row containing the data is added to the 
# archive sheet and then subsequentially removed from the main sheet.

# Since Pandas does not support advanced Excel formatting, Openpyxl will be used in another file to restore the formatting of both sheets.

# NOTE: Every piece of functionality related to specifically updating the data in the main and archive sheets is stored in this file and its corresponding
#       function file.

import pandas as pd
from datetime import datetime
from time import sleep

import FUNCTIONS_update_main_and_archive as function


# MAIN FUNCTIONALITY 👇 ------------------------------------------------------------------------------------------------------------------------------#

def update_main_and_archive(main_sheet, archive_sheet):

    # Correct the main sheet's date-contatining column formats. 
    function.format_main_sheet_dates(main_sheet)

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
        check = function.completion_check(row, current_date)
        
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
    main_import, main_update = function.sort_data(main_sheet, rows_to_import, rows_to_keep)

    # The updated archive sheet is saved to a new DataFrame.
    archive_update = pd.concat([archive_sheet, main_import], ignore_index=True)

    # Correct the archive sheet's date-contatining column formats.
    function.format_archive_sheet_dates(archive_update)

    print(f"\nAmount of rows in main sheet post-update: {main_last_row - len(rows_to_import) + 2}")
    print(f"Amount of rows in archive sheet post-update: {archive_update.last_valid_index() + 2}")
    sleep(1)

    # The main function returns the updated main and archive sheets. 
    return main_update, archive_update
