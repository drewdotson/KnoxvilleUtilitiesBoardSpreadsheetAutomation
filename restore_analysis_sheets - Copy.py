
# The purpose of this file is to restore the pre-update format of the area metrics and month-by-month metrics 
# analysis sheets.

import FUNCTIONS_restore_analysis_sheets as function


# MAIN FUNCTIONALITY 👇 -----------------------------------------------------------------------------------------#

def restore_analysis_sheets(area_metrics_sheet, month_metrics_sheet):

    # Find the row limit of both the main and archive sheets and assign them to variables.
    area_row_limit = function.find_row_limit(area_metrics_sheet)
    month_row_limit = function.find_row_limit(month_metrics_sheet)

    # Restore the row height and column width of both sheets.
    function.restore_height_and_width(area_metrics_sheet,area_row_limit)
    function.restore_height_and_width(month_metrics_sheet, month_row_limit)

    # Restore the font details of both sheets.
    function.restore_font_details(area_metrics_sheet, area_row_limit)
    function.restore_font_details(month_metrics_sheet, month_row_limit)

    # Apply the cell fill colors of both sheets.
    function.apply_column_color(area_metrics_sheet, area_row_limit)
    function.apply_column_color(month_metrics_sheet, month_row_limit)

    # Restore the cell borders of both sheets.
    function.restore_borders(area_metrics_sheet, area_row_limit)
    function.restore_borders(month_metrics_sheet, month_row_limit)

    # Restore the filter tabs in both sheets.
    function.restore_filters(area_metrics_sheet, area_row_limit)
    function.restore_filters(month_metrics_sheet, month_row_limit)
