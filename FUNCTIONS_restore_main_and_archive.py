
# The purpose of this file is to contain the functions used for the restore_main_and_archive file.

from openpyxl.styles import Alignment, Font, Border, Side
import FUNCTIONS_cell_color as colors


# FUNCTIONS 👇 ---------------------------------------------------------------------------------------------------#

# This function is used to return the first row of a specified sheet that does not have installation job data.
# It takes a sheet name as input.
def find_row_limit(sheet_name):

    # Iterate over the "Status" column in every row.
    for cell in sheet_name['N']:

        # When the first row that does not have a value in the "Status" column is found, the function will return 
        # that row's number.
        if cell.value is None:
            return(cell.row)
        else:
            continue
    
    # If there are no rows in the specified sheet without an empty "Status" column, return the first empty row's 
    # number
    return sheet_name.max_row + 1


# Function used for restoring each row's height and each column's width to their original values. It takes a 
# specified sheet name and its row limit as input.
def restore_height_and_width(sheet_name, row_limit):

    # The header row is given a height value of 35
    sheet_name.row_dimensions[1].height = 35

    # Each non-header row is given a height value of 30
    for row in range(2, row_limit):
        sheet_name.row_dimensions[row].height = 30

    # Each column is given its own specific width value
    sheet_name.column_dimensions['A'].width = 32
    sheet_name.column_dimensions['B'].width = 34
    sheet_name.column_dimensions['C'].width = 30
    sheet_name.column_dimensions['D'].width = 32
    sheet_name.column_dimensions['E'].width = 40
    sheet_name.column_dimensions['F'].width = 37
    sheet_name.column_dimensions['G'].width = 60
    sheet_name.column_dimensions['H'].width = 15
    sheet_name.column_dimensions['I'].width = 25
    sheet_name.column_dimensions['J'].width = 25
    sheet_name.column_dimensions['K'].width = 40
    sheet_name.column_dimensions['L'].width = 50
    sheet_name.column_dimensions['M'].width = 70
    sheet_name.column_dimensions['N'].width = 35
    sheet_name.column_dimensions['O'].width = 9
    sheet_name.column_dimensions['P'].width = 25


# Function used to restore font details of each column. It takes a specified sheet name and its row limit as input.
def restore_font_details(sheet_name, row_limit):

    # Restore font details of the header row.
    for row in sheet_name.iter_rows(max_row=1, max_col=14):
        for cell in row:
            cell.font = Font(name='Arial', size=11, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Restore each date value to mm-dd-yyyy format.
    for row in sheet_name.iter_rows(min_row=2, max_row=(row_limit - 1), max_col=6):
        for cell in row:
            cell.number_format = 'mm-dd-yyyy'

    # Restore font details of columns A-L
    for row in sheet_name.iter_rows(min_row=2, max_row=(row_limit - 1), max_col=12):
        for cell in row:
            cell.font = Font(name='Arial', size=16)

    # Restore font details of the "Notes" column.
    for row in sheet_name.iter_rows(min_row=2, max_row=(row_limit - 1), min_col=13, max_col=13):
        for cell in row:
            cell.font = Font(name='Arial', size=11, bold=True)

    # Restore font details of the "Status" Column.
    for row in sheet_name.iter_rows(min_row=2, max_row=(row_limit - 1), min_col=14, max_col=14):
        for cell in row:
            cell.font = Font(name='Arial', size=16, bold=True)


# Function used for restoring the cell fill color to each cell in the "Job Completed Date" and "Address" columns. It
# takes a sheet name and its row limit as input.
def restore_column_color(sheet_name, row_limit):

    # Iterate over the 'Status' column in each row of the specified sheet.
    for row in range(2, row_limit):

        # If the row has a status of "COMPLETED", the two cells will be filled with blue.
        if sheet_name['N' + str(row)].value == 'COMPLETED':
            sheet_name['F' + str(row)].fill = colors.blue
            sheet_name['G' + str(row)].fill = colors.blue
        
        # If the row has a status of "HIGH PRIORITY", the two cells will be filled with magenta.
        elif sheet_name['N' + str(row)].value == 'HIGH PRIORITY':
            sheet_name['F' + str(row)].fill = colors.magenta
            sheet_name['G' + str(row)].fill = colors.magenta

        # If the row has a status of "KUB/GLOBAL", the two cells will be filled with cyan.
        elif sheet_name['N' + str(row)].value == 'KUB/GLOBAL':
            sheet_name['F' + str(row)].fill = colors.cyan
            sheet_name['G' + str(row)].fill = colors.cyan

        # If the row has a status of "NEED TO CALL 811", the two cells will be filled with plum.
        elif sheet_name['N' + str(row)].value == 'NEED TO CALL 811':
            sheet_name['F' + str(row)].fill = colors.plum
            sheet_name['G' + str(row)].fill = colors.plum

        # If the row has a status of "WAITING ON 811", the two cells will be filled with yellow.
        elif sheet_name['N' + str(row)].value == 'WAITING ON 811':
            sheet_name['F' + str(row)].fill = colors.yellow
            sheet_name['G' + str(row)].fill = colors.yellow

        # If the row has a status of "NOTES", the two cells will be filled with gold.
        elif sheet_name['N' + str(row)].value == 'NOTES':
            sheet_name['F' + str(row)].fill = colors.gold
            sheet_name['G' + str(row)].fill = colors.gold

        # If the row has a status of "ON HOLD/WAITING ON CUST TO CALL", the two cells will be filled with olive.
        elif sheet_name['N' + str(row)].value == 'ON HOLD/WAITING ON CUST TO CALL':
            sheet_name['F' + str(row)].fill = colors.olive
            sheet_name['G' + str(row)].fill = colors.olive

        # If the row has a status of "READY TO BURY", the two cells will be filled with yellow green.
        elif sheet_name['N' + str(row)].value == 'READY TO BURY':
            sheet_name['F' + str(row)].fill = colors.yellow_green
            sheet_name['G' + str(row)].fill = colors.yellow_green

        # If the row has a status of "SCHEDULED", the two cells will be filled with red.
        elif sheet_name['N' + str(row)].value == 'SCHEDULED':
            sheet_name['F' + str(row)].fill = colors.red
            sheet_name['G' + str(row)].fill = colors.red

        # If the row has a status of "CANCELLED", the two cells will be filled with black and the text font color
        # will be changed to white and have a strikethrough effect.
        elif sheet_name['N' + str(row)].value == 'CANCELLED':
            sheet_name['F' + str(row)].fill = colors.black
            sheet_name['F' + str(row)].font = Font(color='ffffff', strike=True)
            sheet_name['G' + str(row)].fill = colors.black
            sheet_name['G' + str(row)].font = Font(color='ffffff', strike=True)

        # If the "Notes" column has a value, then the cell is filled with gold.
        if sheet_name['M' + str(row)].value is not None:
            sheet_name['M' + str(row)].fill = colors.gold


# Function used for restoring borders to cells. It uses a specified sheet name and its row limit as input.
def restore_borders(sheet_name, row_limit):

    # Create the border.
    border = Border(
                    top=Side(style='thin'),
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    bottom=Side(style='thin'),
                    )
    
    # Iterate over every row in the specifed sheet and apply the border to all cells in every used column.
    for row in sheet_name.iter_rows(max_row=(row_limit - 1), max_col=14):
        for cell in row:
            cell.border = border


# Function used for restoring the filter tabs in each column. It uses a specified sheet name and its row limit as
# input.
def restore_filters(sheet_name, row_limit):
    sheet_name.auto_filter.ref = f"A1:N{row_limit}"


# Function used for restoring the Legend of the main sheet. It takes the main sheet name as input.
def restore_legend(main_sheet):

    # Restore the legend colors
    main_sheet['O3'].fill = colors.blue
    main_sheet['O4'].fill = colors.magenta
    main_sheet['O5'].fill = colors.cyan
    main_sheet['O6'].fill = colors.plum
    main_sheet['O7'].fill = colors.yellow
    main_sheet['O8'].fill = colors.gold
    main_sheet['O9'].fill = colors.olive
    main_sheet['O10'].fill = colors.yellow_green
    main_sheet['O11'].fill = colors.red
    main_sheet['O12'].fill = colors.black

    # Restore the legend labels
    main_sheet['P3'].value = "Completed"
    main_sheet['P4'].value = "High Priority"
    main_sheet['P5'].value = "KUB / Global"
    main_sheet['P6'].value = "Need to Call 811"
    main_sheet['P7'].value = "Waiting on 811"
    main_sheet['P8'].value = "Notes"
    main_sheet['P9'].value = "On Hold / Waiting on Customer to Call"
    main_sheet['P10'].value = "Ready to Bury"
    main_sheet['P11'].value = "Scheduled"
    main_sheet['P12'].value = "Cancelled"

    # Restore font details for the legend labels.
    for row in main_sheet.iter_rows(min_row=3, max_row=9, min_col=16, max_col=16):
        for cell in row:
            cell.font = Font(name='Calibri', size=12)


# Function used for restoring the filter tabs in each column. It uses a specified sheet name and its row limit as
# input.
def restore_filters(sheet_name, row_limit):
    sheet_name.auto_filter.ref = f"A1:N{row_limit}"
