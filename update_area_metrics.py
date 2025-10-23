# The purpose of this file is to update the area metrics analysis sheet. This sheet is modified based off the 
# installation job metrics that are pulled in from the main and archive sheets. The metrics are created from 
# scratch upon each iteration of this program.
# Since Pandas does not support advanced Excel formatting, Openpyxl will be used in another file to restore the 
# formatting of each sheet.
# NOTE: This file contains every piece of functionality for specifically modifying the data present in the area
#       sheet besides the time check functions as these functions are also used by the area metrics updating file.

import checks
import copy
import pandas as pd

# Analysis Metrics

# Average Drop completion time (Start to finish)
# Average 811 call to 811 mark time
# Average 811 mark to completion time
# Amount of jobs completed in less than 10 days
# Amount of jobs completed in between 10 and 16 days
# Amount of jobs completed in more than 16 days
# Amount of jobs completed on time (<= 16 days)
# Percentage of jobs completed on time (<= 16 days)


# FUNCTIONS FOR GATHERING DATA 👇 -------------------------------------------------------------------------------#

# Function used for creating dictionary that assigns all present work area IDs in the dataset for a specified
# Excel sheet as keys. It takes the sheet name as input.
def areas(sheet):

    # An empty dictionary is assigned as the area ID dictionary.
    area_dict = {}

    # The first and last row in the Excel sheet are saved as variables.
    first_row = sheet.first_valid_index()
    last_row = sheet.last_valid_index()

    # The function iterates over every row in the Excel sheet.
    for row in sheet[first_row:last_row + 1].itertuples():
        
        # The value in the eight column (which contains the area ID abbreviations) is assigned as the area key
        # for the specified row
        area_key = row[8]

        # To ensure that the area key is of the correct format, the value's data type is checked to see if it
        # is a text string. Then the value is stripped of leading / ending spaces and is capitalized.
        # NOTE: if the value's data type is not a text string, then the value for the specified row is not 
        # considered to be added as a key.
        if type(area_key) == str:
            area_key = area_key.strip().upper()

            # The newly modified value is checked to see if:
            #   1. The value is already present in the dictionary as a key
            #   2. The length of the value is 2 characters (All work area IDs are two letters wrong)
            # If the value meets both conditions, then it is added as a key in the dictionary.
            # NOTE: all keys in the dictionary are assigned with an empty list as their value.
            if area_key not in area_dict and len(area_key) == 2:
                area_dict[area_key] = []

    # After the function iterates over every row in the Excel sheet, the area ID dictionary is returned.
    return area_dict


# Function used for merging the area ID dictionaries from both the main and archive Excel sheets into one
# dictionary. It takes the main sheet name and archive sheet name as inputs.
def merge_sheet_dicts(main_sheet, archive_sheet):

    # The area ID dictionaries for both Excel sheets are assigned as variables
    main_areas = areas(main_sheet)
    archive_areas = areas(archive_sheet)
    
    # Both dictionaries are merged together into one dictionary and assigned to a variable.
    # NOTE: There will be no overlapping keys as by using this merge method, overlapping values from the
    #       second dictionary, archive_areas, will overwrite overlapping values from the first dictionary,
    #       main_areas.
    area_dict = main_areas | archive_areas

    # After the merge is complete, the new dictionary is returned.
    return area_dict


# Function used for adding completion times (from start to finish) to the installation job times 
# dictionary. The function takes the area ID dictionary and the Excel sheet name as inputs.
def add_start_to_complete_times_dict(area_dict, sheet):
    
    # A copy of the area dictionary is created for the purpose of storing each work area's completed
    # installation job times (from start to finish).
    start_to_complete_times_dict = copy.deepcopy(area_dict)

    # The first and last row in the Excel sheet are saved as variables.
    first_row = sheet.first_valid_index()
    last_row = sheet.last_valid_index()

    # The function iterates over every row in the Excel sheet.
    # NOTE: if a row does not pass every check, then no job time value for the row is appended to the area
    #       ID dictionary.
    for row in sheet[first_row:last_row + 1].itertuples():

        # The job completion time for the specified row is assigned to a variable.
        job_time = checks.start_to_complete_time_check(row)

        # The job time value is check to see if:
        #   1. There is an actual job time value present
        #   2. The date type of the row's area ID column is a string
        if job_time is not None and type(area_ID := row[8]) == str:

            # The value of the area ID column is checked to see if it matches a area ID present in the area 
            # ID dictionary, then the job time is appended as a value to the corresponding area ID.
            area_ID = area_ID.strip().upper()
            if area_ID in start_to_complete_times_dict:
                start_to_complete_times_dict[area_ID].append(job_time)

    # After the function iterates over each row in the Excel sheet, the area ID dictionary containing the
    # corresponding time values is returned.
    return start_to_complete_times_dict


# Function used to gather all completed job time values from both the main Excel sheet and the archive
# Excel sheets into one dictionary. It takes the area ID dictionary, main sheet name, and archive sheet
# name as inputs.
def create_total_start_to_complete_dict(area_dict, main_sheet, archive_sheet):

    # All of the completed job times from the main Excel sheet are appended into the new completed job
    # times dictionary.
    main_job_times_dict = add_start_to_complete_times_dict(area_dict, main_sheet)

    # All of the completed job times from the archive Excel sheet are appended into the completed job
    # times dictionary.
    start_to_complete_job_times_dict = add_start_to_complete_times_dict(main_job_times_dict, archive_sheet)

    # The completed job times dictionary is returned.
    return start_to_complete_job_times_dict


# Function used for adding 811 call to 811 mark times to the call to mark times dictionary. The function 
# takes the area ID dictionary and the Excel sheet name as inputs.
def add_call811_to_mark811_times_dict(area_dict, sheet):

    # A copy of the area dictionary is created for the purpose of storing each work area's installation job 811 
    # call to 811 mark times.
    call811_to_mark811_times_dict = copy.deepcopy(area_dict)

    # The first and last row in the Excel sheet are saved as variables.
    first_row = sheet.first_valid_index()
    last_row = sheet.last_valid_index()

    # The function iterates over every row in the Excel sheet.
    # NOTE: if a row does not pass every check, then no call to mark value for the row is appended to the 
    #       area ID dictionary.
    for row in sheet[first_row:last_row + 1].itertuples():

        # The call to mark time for the specified row is assigned to a variable.
        mark811_time = checks.call811_to_mark811_time_check(row)

        # The call to mark time value is check to see if:
        #   1. There is an actual call to mark time value present
        #   2. The date type of the row's area ID column is a string
        if mark811_time is not None and type(area_ID := row[8]) == str:

                # The value of the area ID column is checked to see if it matches a area ID present in the area 
                # ID dictionary, then the job time is appended as a value to the corresponding area ID.
                area_ID = area_ID.strip().upper()
                if area_ID in call811_to_mark811_times_dict:
                    call811_to_mark811_times_dict[area_ID].append(mark811_time)

    # After the function iterates over each row in the Excel sheet, the area ID dictionary containing the
    # corresponding time values is returned.
    return call811_to_mark811_times_dict


# Function used to gather all 811 call to 811 mark time values from both the main Excel sheet and the 
# archive Excel sheets into one dictionary. It takes the area ID dictionary, main sheet name, and archive 
# sheet name as inputs.
def create_total_call811_to_mark811_dict(area_dict, main_sheet, archive_sheet):

    # All of the call to mark times from the main Excel sheet are appended into the new call to mark job 
    # times dictionary.
    main_mark811_times_dict = add_call811_to_mark811_times_dict(area_dict, main_sheet)

    # All of the call to mark times from the archive Excel sheet are appended into the  call to mark times 
    # dictionary.
    complete_mark811_times_dict = add_call811_to_mark811_times_dict(main_mark811_times_dict, archive_sheet)

    # The 811 call to 811 mark times dictionary is returned.
    return complete_mark811_times_dict


# Function used for adding job completion times (post-mark) to the mark to complete times dictionary. The 
# function takes the area ID dictionary and the Excel sheet name as inputs.
def add_mark811_to_complete_times_dict(area_dict, sheet):

    # A copy of the area dictionary is created for the purpose of storing each work area's installation job 811 mark to
    # completion times.
    mark811_to_complete_times_dict = copy.deepcopy(area_dict)

    # The first and last row in the Excel sheet are saved as variables.
    first_row = sheet.first_valid_index()
    last_row = sheet.last_valid_index()

    # The function iterates over every row in the Excel sheet.
    # NOTE: if a row does not pass every check, then no call to mark value for the row is appended to the 
    #       area ID dictionary.
    for row in sheet[first_row:last_row + 1].itertuples():

        # The job completion time (post-mark) for the specified row is assigned to a variable.
        complete_post_mark_time = checks.mark811_to_complete_time_check(row)

        # The job completion time value is check to see if:
        #   1. There is an actualjob completion time value present
        #   2. The date type of the row's area ID column is a string
        if complete_post_mark_time is not None and type(area_ID := row[8]) == str:

                # The value of the area ID column is checked to see if it matches a area ID present in the area 
                # ID dictionary, then the job time is appended as a value to the corresponding area ID.
                area_ID = area_ID.strip().upper()
                if area_ID in mark811_to_complete_times_dict:
                    mark811_to_complete_times_dict[area_ID].append(complete_post_mark_time)

    # After the function iterates over each row in the Excel sheet, the area ID dictionary containing the
    # corresponding time values is returned.
    return mark811_to_complete_times_dict


# Function used to gather all job completion time (post-mark) values from both the main Excel sheet and the 
# archive Excel sheets into one dictionary. It takes the area ID dictionary, main sheet name, and archive 
# sheet name as inputs.
def create_total_mark811_to_complete_dict(area_dict, main_sheet, archive_sheet):

    # All of the job completion times (post-mark) from the main Excel sheet are appended into the new mark to
    # job completion times dictionary.
    main_complete_times_dict = add_mark811_to_complete_times_dict(area_dict, main_sheet)

    # All of the job completion times (post-mark) from the archive Excel sheet are appended into the mark to
    # to job completion times dictionary.
    complete_complete_times_dict = add_mark811_to_complete_times_dict(main_complete_times_dict, archive_sheet)

    # The 811 mark to job completion times dictionary is returned.
    return complete_complete_times_dict

# FUNCTIONS FOR CREATING ANALYSIS METRICS 👇 --------------------------------------------------------------------#

# Function used for creating the dictionary containing each work area's average job completion time (from start to
# finish). It uses the completed job times dictionary as input.
def create_average_start_to_complete_time_dict(start_to_complete_times_dict):

    # A copy of the area dictionary is created for the purpose of storing each work area's average installation job
    # completion time (from start to finish).
    average_start_to_complete_time_dict = copy.deepcopy(start_to_complete_times_dict)

    # The function iterates over every work area in the dictionary.
    for area_ID in average_start_to_complete_time_dict:

        # The list of job time values for the specified work area are assigned to a variable.
        job_times = average_start_to_complete_time_dict[area_ID]

        # The job times list is checked to see if it has any values in it.
        if len(job_times) > 0:

            # The average job time for the specified work area is calculated and the value for the key
            # corresponding to the work area is updated to the average job time.
            average_job_time = sum(job_times) / len(job_times)
            average_start_to_complete_time_dict.update({area_ID: round(average_job_time, 2)})

        # If there are no values present in the job times list, then the value for the key
        # corresponding to the work area is updated to "N/A".
        else:
            average_start_to_complete_time_dict.update({area_ID: "N/A"})

    # After the function iterates over each work area, the newly updated dictionary containing each work area's
    # average job completion time (from start to finish) is returned.
    return average_start_to_complete_time_dict


# Function used for creating the dictionary containing each work area's average 811 call to 811 mark time. It 
# uses the 811 call to 811 mark times dictionary as input.
def create_average_call811_to_mark811_time_dict(call811_to_mark811_times_dict):

    # A copy of the area dictionary is created for the purpose of storing each work area's average time for an
    # 811 call to be marked.
    average_call811_to_mark811_time_dict = copy.deepcopy(call811_to_mark811_times_dict)

    # The function iterates over every work area in the dictionary.
    for area_ID in average_call811_to_mark811_time_dict:

        # The list of call to mark time values for the specified work area are assigned to a variable.
        call_to_mark_times = average_call811_to_mark811_time_dict[area_ID]

        # The call to mark times list is checked to see if it has any values in it.
        if len(call_to_mark_times) > 0:
            call_to_mark_avg = sum(call_to_mark_times) / len(call_to_mark_times)
            average_call811_to_mark811_time_dict.update({area_ID: round(call_to_mark_avg, 2)})

        # If there are no values present in the call to mark times list, then the value for the key
        # corresponding to the work area is updated to "N/A".
        else:
            average_call811_to_mark811_time_dict.update({area_ID: "N/A"})
    
    # After the function iterates over each work area, the newly updated dictionary containing each work area's
    # average 811 call to 811 mark time is returned.
    return average_call811_to_mark811_time_dict


# Function used for creating the dictionary containing each work area's average 811 mark time to job completion 
# time. It uses the 811 mark to job completion times dictionary as input.
def create_average_mark811_to_complete_time_dict(mark811_to_complete_time_dict):

    # A copy of the area dictionary is created for the purpose of storing each work area's average time for an
    # installation job to be completed once it had been 811 marked.
    average_mark811_to_complete_time_dict = copy.deepcopy(mark811_to_complete_time_dict)

    # The function iterates over every work area in the dictionary.
    for area_ID in average_mark811_to_complete_time_dict:

        # The list of job completion time (post-mark) values for the specified work area are assigned to a 
        # variable.
        mark_to_complete_times = average_mark811_to_complete_time_dict[area_ID]

        # The job completion times (post-mark) list is checked to see if it has any values in it.
        if len(mark_to_complete_times) > 0:
            key_average = sum(mark_to_complete_times) / len(mark_to_complete_times)
            average_mark811_to_complete_time_dict.update({area_ID: round(key_average, 2)})

        # If there are no values present in the job completion times (post-mark) list, then the value for the key
        # corresponding to the work area is updated to "N/A".
        else:
            average_mark811_to_complete_time_dict.update({area_ID: "N/A"})
    
    # After the function iterates over each work area, the newly updated dictionary containing each work area's
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

    # The function iterates over every work area in the dictionary.
    for area_ID in start_to_complete_times_dict:

        # The three time range dictionaries are assigned the work area ID as a key with no values in it.
        jobs_under_10[area_ID] = None
        jobs_10_to_16[area_ID] = None
        jobs_above_16[area_ID] = None

        # Three variables are created that will be used to store the amount of jobs that were completed in the
        # three time ranges for the specified work area. They are initially assigned a value of 0, but will be 
        # incremented as completed job times are assessed.
        amount_under_10 = 0
        amount_10_to_16 = 0
        amount_above_16 = 0

        # The list of job time values for the specified work area are assigned to a variable.
        job_times = start_to_complete_times_dict[area_ID]

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
        
        # After the function iterates over every value in the list of job times values, then the value for the key
        # corresponding to the work area in each of the three time range dictionaries is updated with the variable
        # for the amount of jobs completed in the corresponding time range.
        jobs_under_10.update({area_ID: amount_under_10})
        jobs_10_to_16.update({area_ID: amount_10_to_16})
        jobs_above_16.update({area_ID: amount_above_16})

    # Once the function iterates over each work area, the three dictionaries containg the amount of completed 
    # installation jobs for the three time ranges are returned.
    return jobs_under_10, jobs_10_to_16, jobs_above_16


# Function used for creating the dictionary containing the percentage of installation jobs that were completed on 
# time (from start to finish) in each work area. It uses the completed job times dictionary, the amount of jobs
# completed in under 10 days dictionary, and the amount of jobs completed in between 10 and 16 days as input.
def create_percent_jobs_on_time_dict(start_to_complete_times_dict, jobs_under_10, jobs_10_to_16):

    # A copy of the area dictionary is created for the purpose of storing each work area's percentage of installation
    # jobs that were completed on time (<= 16 days from start to finish.)
    percent_jobs_on_time_dict = copy.deepcopy(start_to_complete_times_dict)

    # The function iterates over every work area in the dictionary.
    for area_ID in percent_jobs_on_time_dict:

        # The length of the key corresponding to the work area in the completed job times dictionary is checked.
        if len(percent_jobs_on_time_dict[area_ID]) != 0:

            # If there are completed jobs in the specified work area, then the percentage of jobs that were
            # completed on time (<= 16 days) is calculated by dividing the amount of jobs completed in 16 or less
            # days in the corresponding work area by the total amount of jobs completed in the corresponding work
            # area. This percentage is assigned to a varaiable
            percent_under_16 = round(((jobs_under_10[area_ID] + jobs_10_to_16[area_ID]) / len(percent_jobs_on_time_dict[area_ID]) * 100), 2)

        # If there area no completed jobs in the corresponding work area, the precentage variable is set to "N/A"    
        else:
            percent_under_16 = "N/A"
        
        # The work area's key value is updated to the percentage variable
        percent_jobs_on_time_dict.update({area_ID: percent_under_16})
    
    # Once the function iterates over each work area, the dictionary containing the percentage of installation 
    # jobs that were completed on time (from start to finish) in each work area is returned.
    return percent_jobs_on_time_dict

# MAIN FUNCTIONALITY 👇 -----------------------------------------------------------------------------------------#

def update_area_metrics(main_sheet, archive_sheet):

    # Create a dictionary in which the keys are every work area ID present in the main and archive sheets.
    area_dict = merge_sheet_dicts(main_sheet, archive_sheet)

    # Create a dictionary that associates all work areas with the time total time each installation job took to be 
    # completed (from start to finish) in each area.
    start_to_complete_times_dict = create_total_start_to_complete_dict(area_dict, main_sheet, archive_sheet)

    # Create a dictionary that associates all work areas with the time total time each 811 call took to be marked in each 
    # area.
    call811_to_mark811_times_dict = create_total_call811_to_mark811_dict(area_dict, main_sheet, archive_sheet)

    # Create a dictionary that associates all work areas with the time total time each installation job that was 811 marked 
    # took to be completed after the mark was placed in each area.
    mark811_to_complete_time_dict = create_total_mark811_to_complete_dict(area_dict, main_sheet, archive_sheet)

    # Create a dictionary that associates all work areas with the average time an installation job takes to be completed 
    # (from start to finish) in each area.
    average_start_to_complete_time = create_average_start_to_complete_time_dict(start_to_complete_times_dict)

    # Create a dictionary that associates all work areas with the average time an 811 call on a job took to be marked in 
    # each area.
    average_call_to_mark_time = create_average_call811_to_mark811_time_dict(call811_to_mark811_times_dict)

    # Create a dictionary that associates all work areas with the average time an installation job took to be completed
    # after being 811 marked.
    average_mark_to_complete_time = create_average_mark811_to_complete_time_dict(mark811_to_complete_time_dict)

    # Create three dictionaries that each associate all work areas with the amount of completed installation jobs in each
    # for specific time ranges.
    jobs_under_10, jobs_10_to_16, jobs_above_16 = create_amount_jobs_per_time_range_dicts(start_to_complete_times_dict)

    # Create a dictionary that associates all work areas with the percentage of installation jobs that were completed
    # on time in each area.
    percent_jobs_on_time = create_percent_jobs_on_time_dict(start_to_complete_times_dict, jobs_under_10, jobs_10_to_16)
 
    # Create the DataFrame that will be saved into the "Area Metrics" sheet of the Excel file.
    updated_area_metrics = pd.DataFrame({
        'Area ID': [],
        'Average Completion Time (Start to Finish)': [],
        'Average 811 Call to 811 Mark Time': [],
        'Average 811 Mark to Completion Time': [],
        '% of Jobs Completed on Time (<= 16 Days)': [],
        '# of Jobs (< 10 Days)': [],
        '# of Jobs ( 10 <= x <= 16 Days)': [],
        '# of Jobs (> 16 Days)': []
    })

    # Import the values of each metric-containing dictionary into the corrseponding column in the DataFrame.
    updated_area_metrics['Area ID'] = list(area_dict.keys())
    updated_area_metrics['Average Completion Time (Start to Finish)'] = list(average_start_to_complete_time.values())
    updated_area_metrics['Average 811 Call to 811 Mark Time'] = list(average_call_to_mark_time.values())
    updated_area_metrics['Average 811 Mark to Completion Time'] = list(average_mark_to_complete_time.values())
    updated_area_metrics['% of Jobs Completed on Time (<= 16 Days)'] = list(percent_jobs_on_time.values())
    updated_area_metrics['# of Jobs (< 10 Days)'] = list(jobs_under_10.values())
    updated_area_metrics['# of Jobs ( 10 <= x <= 16 Days)'] = list(jobs_10_to_16.values())
    updated_area_metrics['# of Jobs (> 16 Days)'] = list(jobs_above_16.values())

    # Once the DataFrame is complete, it is returned.
    return updated_area_metrics
