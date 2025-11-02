
# The purpose of this file is to update the month-by-month metrics analysis sheet. This sheet is modified based off the installation job metrics that 
# are pulled in from the main and archive sheets. The metrics are created from scratch upon each iteration of this program.

# Since Pandas does not support advanced Excel formatting, Openpyxl will be used in another file to restore the formatting of each sheet.

# NOTE: every piece of functionality for specifically modifying the data present in the month metrics sheet (besides the time check functions) are
#       contained within this file and its correspsonding function file.

import pandas as pd

import FUNCTIONS_update_month_metrics as function

# Analysis Metrics

# Average Drop completion time (Start to finish)
# Average 811 call to 811 mark time
# Average 811 mark to completion time
# Amount of jobs completed in less than 10 days
# Amount of jobs completed in between 10 and 16 days
# Amount of jobs completed in more than 16 days
# Amount of jobs completed on time (<= 16 days)
# Percentage of jobs completed on time (<= 16 days)


# MAIN FUNCTIONALITY 👇 ------------------------------------------------------------------------------------------------------------------------------#

def update_month_metrics(main_sheet, archive_sheet):

    # Create a dictionary in which the keys are every month ID present in the main and archive sheets.
    month_dict = function.merge_and_sort_sheet_dicts(main_sheet, archive_sheet)

    # Create a dictionary that associates all months with the time total time each installation job took to be 
    # completed (from start to finish) during each month.
    start_to_complete_times_dict = function.create_total_start_to_complete_dict(month_dict, main_sheet, archive_sheet)

    # Create a dictionary that associates all months with the time total time each 811 call took to be marked during each 
    # month.
    call811_to_mark811_times_dict = function.create_total_call811_to_mark811_dict(month_dict, main_sheet, archive_sheet)

    # Create a dictionary that associates all months with the time total time each installation job that was 811 marked 
    # took to be completed after the mark was placed during each month.
    mark811_to_complete_time_dict = function.create_total_mark811_to_complete_dict(month_dict, main_sheet, archive_sheet)

    # Create a dictionary that associates all months with the average time an installation job takes to be completed 
    # (from start to finish) during each month.
    average_start_to_complete_time = function.create_average_start_to_complete_time_dict(start_to_complete_times_dict)

    # Create a dictionary that associates all months with the average time an 811 call on a job took to be marked during 
    # each month.
    average_call_to_mark_time = function.create_average_call811_to_mark811_time_dict(call811_to_mark811_times_dict)

    # Create a dictionary that associates all months with the average time an installation job took to be completed
    # after being 811 marked during each month.
    average_mark_to_complete_time = function.create_average_mark811_to_complete_time_dict(mark811_to_complete_time_dict)

    # Create three dictionaries that each associate all months with the amount of completed installation jobs in each
    # for specific time ranges.
    jobs_under_10, jobs_10_to_16, jobs_above_16 = function.create_amount_jobs_per_time_range_dicts(start_to_complete_times_dict)

    # Create a dictionary that associates all months with the percentage of installation jobs that were completed
    # on time during each month.
    percent_jobs_on_time = function.create_percent_jobs_on_time_dict(start_to_complete_times_dict, jobs_under_10, jobs_10_to_16)

    # Create the DataFrame that will be saved into the "Month-by-Month Metrics" sheet of the Excel file.
    updated_month_metrics = pd.DataFrame({
        'Month': [],
        'Average Completion Time (Start to Finish)': [],
        'Average 811 Call to 811 Mark Time': [],
        'Average 811 Mark to Completion Time': [],
        '% of Jobs Completed on Time (<= 16 Days)': [],
        '# of Jobs (< 10 Days)': [],
        '# of Jobs ( 10 <= x <= 16 Days)': [],
        '# of Jobs (> 16 Days)': []
    })

    # Import the values of each metric-containing dictionary into the corrseponding column in the DataFrame.
    updated_month_metrics['Month'] = list(month_dict.keys())
    updated_month_metrics['Average Completion Time (Start to Finish)'] = list(average_start_to_complete_time.values())
    updated_month_metrics['Average 811 Call to 811 Mark Time'] = list(average_call_to_mark_time.values())
    updated_month_metrics['Average 811 Mark to Completion Time'] = list(average_mark_to_complete_time.values())
    updated_month_metrics['% of Jobs Completed on Time (<= 16 Days)'] = list(percent_jobs_on_time.values())
    updated_month_metrics['# of Jobs (< 10 Days)'] = list(jobs_under_10.values())
    updated_month_metrics['# of Jobs ( 10 <= x <= 16 Days)'] = list(jobs_10_to_16.values())
    updated_month_metrics['# of Jobs (> 16 Days)'] = list(jobs_above_16.values())

    # Once the DataFrame is complete, it is returned.
    return updated_month_metrics
