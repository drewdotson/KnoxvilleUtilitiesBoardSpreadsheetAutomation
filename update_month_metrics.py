# The purpose of this file is to update the month-by-month metrics analysis sheet. This sheet is modified based 
# off the installation job metrics that are pulled in from the main and archive sheets. The metrics are created 
# from scratch upon each iteration of this program.
# Since Pandas does not support advanced Excel formatting, Openpyxl will be used in another file to restore the 
# formatting of each sheet.
# NOTE: This file contains every piece of functionality for specifically modifying the data present in the area
#       sheet besides the time check functions as these functions are also used by the area metrics updating file.

import checks
import copy
import pandas as pd
import re

# Analysis Metrics

# Average Drop completion time (Start to finish)
# Average 811 call to 811 mark time
# Average 811 mark to completion time
# Amount of jobs completed in less than 10 days
# Amount of jobs completed in between 10 and 16 days
# Amount of jobs completed in more than 16 days
# Amount of jobs completed on time (<= 16 days)
# Percentage of jobs completed on time (<= 16 days)

# FUNCTIONS FOR GATHERING DATA 👇 --------------------------------------------------------------------------------#

# Function used for creating dictionary that assigns all present month and year combinations in the dataset for a 
# specified Excel sheet as keys. It takes the sheet name as input.
def months(sheet):

    # An empty dictionary is assigned as the month + year ID dictionary.
    month_year_dict = {}

    # Two dictionaries are created to assign month and year identifiers to integers. This will be used to verify 
    # proper date formats and to assist in organizing the month + year ID dictionary in the merge dictionaries 
    # function later.
    month_dict = {
                  "01": 1,
                  "02": 2,
                  "03": 3,
                  "04": 4,
                  "05": 5,
                  "06": 6,
                  "07": 7,
                  "08": 8,
                  "09": 9,
                  "10": 10,
                  "11": 11,
                  "12": 12
                }
    year_dict = {
                  "2023": 100,
                  "2024": 200,
                  "2025": 300,
                  "2026": 400,
                  "2027": 500,
                  "2028": 600,
                  "2029": 700,
                  "2030": 800,
                  "2031": 900,
                  "2032": 1000,
                  "2033": 1100,
                  "2034": 1200,
                  "2035": 1300,
                  "2036": 1400,
                  "2037": 1500,
                  "2038": 1600,
                  "2039": 1700,
                  "2040": 1800
                }

    # The function iterates over every row in the Excel sheet.
    for row in sheet[sheet.first_valid_index():sheet.last_valid_index() + 1].itertuples():
        
        # The value in the first column (which contains the date that the installation job began) is split into four 
        # different parts (year, month, day, and time respectively), converted to a string, and assigned as as a 
        # variable.
        row_date = re.split(r'[-\s]+', str(row[1]))

        # If the variable's length is not equal to one (a length of one means the value is "NaT"), then the function
        # checks if the second value in the string matches a key in the month dictionary.
        if len(row_date) != 1:

            # If the month value in the string does match a key in the month dictionary and the year value in the 
            # string does match a key in the year dictionary, then the following two actions are performed 👇
            if row_date[1] in month_dict and row_date[0] in year_dict:

                # 1. The value of the two keys are added together and assigned to a variable.
                month_year_value = month_dict[row_date[1]] + year_dict[row_date[0]]
                # 2. The keys are concatenated together and assigned to a variable to be considered as a key into 
                # the month + year ID dictionary.
                month_year_key = f"{row_date[1]}-{row_date[0]}"

                # If this key is not already in the month + year ID dictionary, then the key is added and the value 
                # variable is assigned as the key's value
                if month_year_key not in month_year_dict:
                    month_year_dict[month_year_key] = month_year_value

    # After the function iterates over every row in the Excel sheet, the month + year ID dictionary is returned.
    return month_year_dict


# Function used for merging the month + year ID dictionaries from both the main and archive Excel sheets into one
# dictionary. It takes the main sheet name and archive sheet name as inputs.
def merge_and_sort_sheet_dicts(main_sheet, archive_sheet):

    # The month + year ID dictionaries for both Excel sheets are assigned as variables
    main_months = months(main_sheet)
    archive_months = months(archive_sheet)
    
    # Both dictionaries are merged together into one dictionary and assigned to a variable.
    # To ensure proper month and year order in the dictionary, the dictionary is sorted by the each key's value
    # in ascending order.
    # NOTE: There will be no overlapping keys as by using this merge method, overlapping values from the second
    #       dictionary, archive_areas, will overwrite overlapping values from the first dictionary, main_areas. This 
    #       will not lead to improper sorting as duplicate keys will have the same value.
    month_dict = dict(sorted((main_months | archive_months).items(), key=lambda item: item[1]))

    # After the keys in the month + year dictionary are sorted, each key's value will be replaced with an empty list.
    # NOTE: from here on out the month + year ID dictionary will be referred to as the month ID dictionary.
    for key in month_dict:
        month_dict[key] = []

    # After the merge and sort process is complete, the new month ID dictionary is returned.
    return month_dict


# Function used for adding completion times (from start to finish) to the installation job times dictionary. The 
# function takes the month ID dictionary and the Excel sheet name as inputs.
def add_start_to_complete_times_dict(month_dict, sheet):
    
    # A copy of the month ID dictionary is created for the purpose of storing each month's completed installation job 
    # times (from start to finish).
    start_to_complete_times_dict = copy.deepcopy(month_dict)

    # The first and last row in the Excel sheet are saved as variables.
    first_row = sheet.first_valid_index()
    last_row = sheet.last_valid_index()

    # The function iterates over every row in the Excel sheet.
    # NOTE: if a row does not pass every check, then no job time value for the row is appended to the job times 
    #       dictionary.
    for row in sheet[first_row:last_row + 1].itertuples():

        # The job completion time for the specified row is assigned to a variable.
        job_time = checks.start_to_complete_time_check(row)

        # The month ID for the specfied row is generated in the same fashion as the months() function.
        row_date = re.split(r'[-\s]+', str(row[1]))

        # The job start date column is check to see if its value is a date.
        if len(row_date) != 1:
            month_key = f"{row_date[1]}-{row_date[0]}"

            # The job time value is check to see if there is an actual job time value present
            if job_time is not None:

                # The value of the month ID is checked to see if it matches a key present in the job times dictionary.
                if month_key in start_to_complete_times_dict:

                    # If there is a present key in the dictionary, the job time is appended to its list.
                    start_to_complete_times_dict[month_key].append(job_time)

    # After the function iterates over every row in the Excel sheet, the job times dictionary containing the
    # corresponding time values is returned.
    return start_to_complete_times_dict


# Function used to gather all completed job time values from both the main Excel sheet and the archive Excel sheets into 
# one dictionary. It takes the month ID dictionary, main sheet name, and archive sheet name as inputs.
def create_total_start_to_complete_dict(month_dict, main_sheet, archive_sheet):

    # All of the completed job times from the main Excel sheet are appended into the new completed job times dictionary.
    main_job_times_dict = add_start_to_complete_times_dict(month_dict, main_sheet)

    # All of the completed job times from the archive Excel sheet are appended into the completed job times dictionary.
    start_to_complete_job_times_dict = add_start_to_complete_times_dict(main_job_times_dict, archive_sheet)

    # The completed job times dictionary is returned.
    return start_to_complete_job_times_dict


# Function used for adding 811 call to 811 mark times to the call to mark times dictionary. The function takes the month 
# ID dictionary and the Excel sheet name as inputs.
def add_call811_to_mark811_times_dict(month_dict, sheet):

    # A copy of the month ID dictionary is created for the purpose of storing each month's installation job 811 call to 
    # 811 mark times.
    call811_to_mark811_times_dict = copy.deepcopy(month_dict)

    # The first and last row in the Excel sheet are saved as variables.
    first_row = sheet.first_valid_index()
    last_row = sheet.last_valid_index()

    # The function iterates over every row in the Excel sheet.
    # NOTE: if a row does not pass every check, then no call to mark value for the row is appended to the mark times 
    #       dictionary.
    for row in sheet[first_row:last_row + 1].itertuples():

        # The call to mark time for the specified row is assigned to a variable.
        mark811_time = checks.call811_to_mark811_time_check(row)

        # The call to mark time value is check to see there is an actual call to mark time value present
        if mark811_time is not None:

            # The month ID for the specfied row is generated in the same fashion as the months() function.
            row_date = re.split(r'[-\s]+', str(row[1]))

            # The job start date column is check to see if its value is a date.
            if len(row_date) != 1:
                month_key = f"{row_date[1]}-{row_date[0]}"

                # The value of the month ID is checked to see if it matches a key present in the mark times 
                # dictionary.
                if month_key in call811_to_mark811_times_dict:

                    # If there is a present key in the dictionary, the mark time is appended to its list.
                    call811_to_mark811_times_dict[month_key].append(mark811_time)

    # After the function iterates over every row in the Excel sheet, the mark times dictionary containing the
    # corresponding time values is returned.
    return call811_to_mark811_times_dict


# Function used to gather all 811 call to 811 mark time values from both the main Excel sheet and the archive Excel 
# sheets into one dictionary. It takes the month ID dictionary, main sheet name, and archive sheet name as inputs.
def create_total_call811_to_mark811_dict(month_dict, main_sheet, archive_sheet):

    # All of the call to mark times from the main Excel sheet are appended into the new call to mark job times dictionary.
    main_mark811_times_dict = add_call811_to_mark811_times_dict(month_dict, main_sheet)

    # All of the call to mark times from the archive Excel sheet are appended into the call to mark times dictionary.
    complete_mark811_times_dict = add_call811_to_mark811_times_dict(main_mark811_times_dict, archive_sheet)

    # The 811 call to 811 mark times dictionary is returned.
    return complete_mark811_times_dict


# Function used for adding job completion times (post-mark) to the mark to complete times dictionary. The function takes the 
# area ID dictionary and the Excel sheet name as inputs.
def add_mark811_to_complete_times_dict(month_dict, sheet):

    # A copy of the month ID dictionary is created for the purpose of storing each work area's installation job 811 mark to 
    # completion times.
    mark811_to_complete_times_dict = copy.deepcopy(month_dict)

    # The first and last row in the Excel sheet are saved as variables.
    first_row = sheet.first_valid_index()
    last_row = sheet.last_valid_index()

    # The function iterates over every row in the Excel sheet.
    # NOTE: if a row does not pass every check, then no call to mark value for the row is appended to the mark to complete 
    #       times dictionary.
    for row in sheet[first_row:last_row + 1].itertuples():

        # The job completion time (post-mark) for the specified row is assigned to a variable.
        complete_post_mark_time = checks.mark811_to_complete_time_check(row)

        # The job completion time value is check to see if there is an actual job completion time value present
        if complete_post_mark_time is not None:

            # The month ID for the specfied row is generated in the same fashion as the months() function.
            row_date = re.split(r'[-\s]+', str(row[1]))

            # The job start date column is check to see if its value is a date.
            if len(row_date) != 1:
                month_key = f"{row_date[1]}-{row_date[0]}"

                # The value of the month ID is checked to see if it matches a key present in the mark to complete times 
                # dictionary.
                if month_key in mark811_to_complete_times_dict:

                    # If there is a present key in the dictionary, the mark to complete time is appended to its list.
                    mark811_to_complete_times_dict[month_key].append(complete_post_mark_time)

    # After the function iterates over every row in the Excel sheet, the mark to complete times dictionary containing 
    # the corresponding time values is returned.
    return mark811_to_complete_times_dict


# Function used to gather all job completion time (post-mark) values from both the main Excel sheet and the archive Excel 
# sheets into one dictionary. It takes the month ID dictionary, main sheet name, and archive sheet name as inputs.
def create_total_mark811_to_complete_dict(month_dict, main_sheet, archive_sheet):

    # All of the job completion times (post-mark) from the main Excel sheet are appended into the new mark to job 
    # completion times dictionary.
    main_complete_times_dict = add_mark811_to_complete_times_dict(month_dict, main_sheet)

    # All of the job completion times (post-mark) from the archive Excel sheet are appended into the mark to job completion 
    # times dictionary.
    complete_complete_times_dict = add_mark811_to_complete_times_dict(main_complete_times_dict, archive_sheet)

    # The 811 mark to job completion times dictionary is returned.
    return complete_complete_times_dict

# FUNCTIONS FOR CREATING ANALYSIS METRICS 👇 ---------------------------------------------------------------------#

# Function used for creating the dictionary containing each month's average job completion time (from start to finish). It 
# uses the completed job times dictionary as input.
def create_average_start_to_complete_time_dict(start_to_complete_times_dict):

    # A copy of the month ID dictionary is created for the purpose of storing each month's average installation job
    # completion time (from start to finish).
    average_start_to_complete_time_dict = copy.deepcopy(start_to_complete_times_dict)

    # The function iterates over every month in the dictionary.
    for month_ID in average_start_to_complete_time_dict:

        # The list of job time values for the specified month are assigned to a variable.
        job_times = average_start_to_complete_time_dict[month_ID]

        # The job times list is checked to see if it has any values in it.
        if len(job_times) > 0:

            # The average job time for the specified month is calculated and the value for the key corresponding to the month 
            # is updated to the average job time.
            average_job_time = sum(job_times) / len(job_times)
            average_start_to_complete_time_dict.update({month_ID: round(average_job_time, 2)})

        # If there are no values present in the job times list, then the value for the key corresponding to the month is 
        # updated to "N/A".
        else:
            average_start_to_complete_time_dict.update({month_ID: "N/A"})

    # After the function iterates over every month, the newly updated dictionary containing each month's average job completion
    # time (from start to finish) is returned.
    return average_start_to_complete_time_dict


# Function used for creating the dictionary containing each month's average 811 call to 811 mark time. It 
# uses the 811 call to 811 mark times dictionary as input.
def create_average_call811_to_mark811_time_dict(call811_to_mark811_times_dict):

    # A copy of the month ID dictionary is created for the purpose of storing each month's average time for an
    # 811 call to be marked.
    average_call811_to_mark811_time_dict = copy.deepcopy(call811_to_mark811_times_dict)

    # The function iterates over every month in the dictionary.
    for month_ID in average_call811_to_mark811_time_dict:

        # The list of call to mark time values for the specified month are assigned to a variable.
        call_to_mark_times = average_call811_to_mark811_time_dict[month_ID]

        # The call to mark times list is checked to see if it has any values in it.
        if len(call_to_mark_times) > 0:

            # The average call to mark time for the specified month is calculated and the value for the key
            # corresponding to the month is updated to the average call to mark time.
            call_to_mark_avg = sum(call_to_mark_times) / len(call_to_mark_times)
            average_call811_to_mark811_time_dict.update({month_ID: round(call_to_mark_avg, 2)})

        # If there are no values present in the call to mark times list, then the value for the key
        # corresponding to the month is updated to "N/A".
        else:
            average_call811_to_mark811_time_dict.update({month_ID: "N/A"})
    
    # After the function iterates over every month, the newly updated dictionary containing month's average
    # 811 call to 811 mark time is returned.
    return average_call811_to_mark811_time_dict


# Function used for creating the dictionary containing each month's average 811 mark time to job completion 
# time. It uses the 811 mark to job completion times dictionary as input.
def create_average_mark811_to_complete_time_dict(mark811_to_complete_time_dict):

    # A copy of the month ID dictionary is created for the purpose of storing each month's average time for an
    # installation job to be completed once it had been 811 marked.
    average_mark811_to_complete_time_dict = copy.deepcopy(mark811_to_complete_time_dict)

    # The function iterates over every month in the dictionary.
    for month_ID in average_mark811_to_complete_time_dict:

        # The list of job completion time (post-mark) values for the specified month are assigned to a 
        # variable.
        mark_to_complete_times = average_mark811_to_complete_time_dict[month_ID]

        # The job completion times (post-mark) list is checked to see if it has any values in it.
        if len(mark_to_complete_times) > 0:
            key_average = sum(mark_to_complete_times) / len(mark_to_complete_times)
            average_mark811_to_complete_time_dict.update({month_ID: round(key_average, 2)})

        # If there are no values present in the job completion times (post-mark) list, then the value for the key
        # corresponding to the month is updated to "N/A".
        else:
            average_mark811_to_complete_time_dict.update({month_ID: "N/A"})
    
    # After the function iterates over every month, the newly updated dictionary containing each month's
    # average 811 mark time to job completion is returned.
    return average_mark811_to_complete_time_dict


# Function used to create the three dictionaries containg the amount of completed installation jobs that were
# completed, from start to finish, in specific time ranges (< 10 days, 10 to 16 days, and > 16 days). It takes
# The completed job times dictionary as input.
def create_amount_jobs_per_time_range_dicts(start_to_complete_times_dict):
    
    # Three empty dictionaries are created for each of the three time ranges.
    jobs_under_10 = {}
    jobs_10_to_16 = {}
    jobs_above_16 = {}

    # The function iterates over each month in the dictionary.
    for month_ID in start_to_complete_times_dict:

        # The three time range dictionaries are assigned the month ID as a key with no values in it.
        jobs_under_10[month_ID] = None
        jobs_10_to_16[month_ID] = None
        jobs_above_16[month_ID] = None

        # Three variables are created that will be used to store the amount of jobs that were completed in the
        # three time ranges for the specified month. They are initially assigned a value of 0, but will be 
        # incremented as completed job times are assessed.
        amount_under_10 = 0
        amount_10_to_16 = 0
        amount_above_16 = 0

        # The list of job time values for the specified month are assigned to a variable.
        job_times = start_to_complete_times_dict[month_ID]

        # The function iterates over every value in the list of job time values.
        for job_time in job_times:

            # If the value is less than ten, then the variable for the amount of jobs completed in under 10 days
            # is incremented by 1.
            if job_time < 10:
                amount_under_10 += 1

            # If the value between 10 and 16, then the variable for the amount of jobs completed in between 10 and 
            # 16 days is incremented by 1.
            elif 10 <= job_time <= 16:
                amount_10_to_16 += 1

            # If the value is greater than 16, then the variable for the amount of jobs completed in above 16 days
            # is incremented by 1.
            else:
                amount_above_16 += 1
        
        # After the function iterates over every value in the list of job times, then the value for the key
        # corresponding to the month in each of the three time range dictionaries is updated with the variable
        # for the amount of jobs completed in the corresponding time range.
        jobs_under_10.update({month_ID: amount_under_10})
        jobs_10_to_16.update({month_ID: amount_10_to_16})
        jobs_above_16.update({month_ID: amount_above_16})

    # Once the function iterates over each month, the three dictionaries containg the amount of completed 
    # installation jobs for the three time ranges are returned.
    return jobs_under_10, jobs_10_to_16, jobs_above_16


# Function used for creating the dictionary containing the percentage of installation jobs that were completed on 
# time (from start to finish) during each month. It uses the completed job times dictionary, the amount of jobs
# completed in under 10 days dictionary, and the amount of jobs completed in between 10 and 16 days as input.
def create_percent_jobs_on_time_dict(start_to_complete_times_dict, jobs_under_10, jobs_10_to_16):

    # A copy of the month ID dictionary is created for the purpose of storing each month's percentage of installation
    # jobs that were completed on time (<= 16 days from start to finish.)
    percent_jobs_on_time_dict = copy.deepcopy(start_to_complete_times_dict)

    # The function iterates over each month in the dictionary.
    for month_ID in percent_jobs_on_time_dict:

        # The length of the key corresponding to the month ID in the completed job times dictionary is checked.
        if len(percent_jobs_on_time_dict[month_ID]) != 0:

            # If there are completed jobs in the specified month, then the percentage of jobs that were
            # completed on time (<= 16 days) is calculated by dividing the amount of jobs completed in 16 or less
            # days in the corresponding month by the total amount of jobs completed in the corresponding month.
            # This percentage is assigned to a varaiable
            percent_under_16 = round(((jobs_under_10[month_ID] + jobs_10_to_16[month_ID]) / len(percent_jobs_on_time_dict[month_ID]) * 100), 2)

        # If there area no completed jobs in the corresponding month, the precentage variable is set to "N/A"    
        else:
            percent_under_16 = "N/A"
        
        # The month's key value is updated to the percentage variable
        percent_jobs_on_time_dict.update({month_ID: percent_under_16})
    
    # Once the function iterates over each month, the dictionary containing the percentage of installation 
    # jobs that were completed on time (from start to finish) during each month is returned.
    return percent_jobs_on_time_dict

# MAIN FUNCTIONALITY 👇 -----------------------------------------------------------------------------------------#

def update_month_metrics(main_sheet, archive_sheet):

    # Create a dictionary in which the keys are every month ID present in the main and archive sheets.
    month_dict = merge_and_sort_sheet_dicts(main_sheet, archive_sheet)

    # Create a dictionary that associates all months with the time total time each installation job took to be 
    # completed (from start to finish) during each month.
    start_to_complete_times_dict = create_total_start_to_complete_dict(month_dict, main_sheet, archive_sheet)

    # Create a dictionary that associates all months with the time total time each 811 call took to be marked during each 
    # month.
    call811_to_mark811_times_dict = create_total_call811_to_mark811_dict(month_dict, main_sheet, archive_sheet)

    # Create a dictionary that associates all months with the time total time each installation job that was 811 marked 
    # took to be completed after the mark was placed during each month.
    mark811_to_complete_time_dict = create_total_mark811_to_complete_dict(month_dict, main_sheet, archive_sheet)

    # Create a dictionary that associates all months with the average time an installation job takes to be completed 
    # (from start to finish) during each month.
    average_start_to_complete_time = create_average_start_to_complete_time_dict(start_to_complete_times_dict)

    # Create a dictionary that associates all months with the average time an 811 call on a job took to be marked during 
    # each month.
    average_call_to_mark_time = create_average_call811_to_mark811_time_dict(call811_to_mark811_times_dict)

    # Create a dictionary that associates all months with the average time an installation job took to be completed
    # after being 811 marked during each month.
    average_mark_to_complete_time = create_average_mark811_to_complete_time_dict(mark811_to_complete_time_dict)

    # Create three dictionaries that each associate all months with the amount of completed installation jobs in each
    # for specific time ranges.
    jobs_under_10, jobs_10_to_16, jobs_above_16 = create_amount_jobs_per_time_range_dicts(start_to_complete_times_dict)

    # Create a dictionary that associates all months with the percentage of installation jobs that were completed
    # on time during each month.
    percent_jobs_on_time = create_percent_jobs_on_time_dict(start_to_complete_times_dict, jobs_under_10, jobs_10_to_16)

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
