
# The purpose of this file is to restore the pre-update format of the main and archive sheets.

import FUNCTIONS_restore_main_and_archive as function


# MAIN FUNCTIONALITY 👇 -----------------------------------------------------------------------------------------#

def restore_main_and_archive(main_sheet, archive_sheet):

    # Find the row limit of both the main and archive sheets and assign them to variables.
    main_row_limit = function.find_row_limit(main_sheet)
    archive_row_limit = function.find_row_limit(archive_sheet)

    # Restore the row height and column width of both sheets.
    function.restore_height_and_width(main_sheet,main_row_limit)
    function.restore_height_and_width(archive_sheet, archive_row_limit)

    # Restore the font details of both sheets.
    function.restore_font_details(main_sheet, main_row_limit)
    function.restore_font_details(archive_sheet, archive_row_limit)

    # Restore the cell fill colors of both sheets.
    function.restore_column_color(main_sheet, main_row_limit)
    function.restore_column_color(archive_sheet, archive_row_limit)

    # Restore the cell borders of both sheets.
    function.restore_borders(main_sheet, main_row_limit)
    function.restore_borders(archive_sheet, archive_row_limit)

    # Restore the cell borders of both sheets.
    function.restore_borders(main_sheet, main_row_limit)
    function.restore_borders(archive_sheet, archive_row_limit)

    # Restore the filter tabs in both sheets.
    function.restore_filters(main_sheet, main_row_limit)
    function.restore_filters(archive_sheet, archive_row_limit)

    # Restore the cell fill legend of the main sheet.
    function.restore_legend(main_sheet)
