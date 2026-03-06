
# The purpose of this file is to restore the pre-update format of the main and archive sheets.

import FUNCTIONS_restore_main_and_archive as function


# MAIN FUNCTIONALITY 👇 -----------------------------------------------------------------------------------------#

def restore_main_and_archive(main_sheet, archive_sheet, workbook):

    # Restore the row height and column width of both sheets.
    function.restore_height_and_width(main_sheet)
    function.restore_height_and_width(archive_sheet)

    # Restore the font details of both sheets.
    function.restore_font_details(main_sheet)
    function.restore_font_details(archive_sheet)

    function.restore_phone_numbers(main_sheet)
    function.restore_phone_numbers(archive_sheet)

    # Restore the cell fill colors of both sheets.
    function.restore_column_color(main_sheet)
    function.restore_column_color(archive_sheet)

    # Restore the cell borders of both sheets.
    function.restore_borders(main_sheet)
    function.restore_borders(archive_sheet)

    # Restore the filter tabs in both sheets.
    function.restore_filters(main_sheet)
    function.restore_filters(archive_sheet)

    function.restore_data_validation(main_sheet, workbook)
    function.restore_data_validation(archive_sheet, workbook)

    function.restore_conditional_formatting(main_sheet)
    function.restore_conditional_formatting(archive_sheet)
    
    # Restore the cell fill legend of the main sheet.
    function.restore_legend(main_sheet)
