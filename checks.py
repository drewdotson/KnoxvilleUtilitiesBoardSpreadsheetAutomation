# The purpose of this file is to contain time check functions that are used by both the area metrics and the
# month-by-month metrics analysis sheets.

import pandas as pd

# Function used to check the amount of time (in days) a fiber installation job took to be completed. This
# function is factoring the using the "Drop Installation Date" column as the starting time. It takes
# a row number as input.
# NOTE: if the specified row does not pass every one of the checks, the function will return None.
def start_to_complete_time_check(row):

    # The value in the job status column is assigned as the row's job status.
    status = row[14]

    # The value is checked to see if if the data type for the value is a text string.
    if type(status) == str: 
        
        # The value is temporarily lowercased and checked to see if it matches the word "completed".
        if status.lower() == "completed":

            # The job start and job completion dates are assigned to variables.
            job_start = row[1]
            job_complete = row[6]

            # The values for both date columns are checked to see if they are empty or not.
            if pd.isnull(job_start) != True and pd.isnull(job_complete) != True:

                # If both variables do contain a non-null value, then the amount of elapsed
                # time for the job is calculated by taking the difference of the two times.
                difference = str(job_complete - job_start).split(" ")

                # To avoid both outliers and date typos, for example a job taking a year to complete,
                # only job completion times of between 0 and 100 days are returned.
                if 0 <= int(difference[0]) <= 100:
                    return int(difference[0])


# Function used to check the amount of time (in days) a fiber installation job took to be marked after an 811 
# call was made. This function is factoring the using the "811 Called in Date" column as the call time and 
# the "811 Marked Date" column as the mark time. It takes a row number as input.
# NOTE: if the specified row does not pass both of the checks, the function will return None.
def call811_to_mark811_time_check(row):

    # The 811 call and 811 mark dates are assigned to variables.
    job_call811 = row[2]
    job_mark811 = row[4]

    # The values for both date columns are checked to see if they are empty or not.
    if pd.isnull(job_call811) != True and pd.isnull(job_mark811) != True:

        # If both variables do contain a non-null value, then the amount of elapsed
        # time for the job is calculated by taking the difference of the two times.
        difference = str(job_mark811 - job_call811).split(" ")

        # To avoid both outliers and date typos, for example a call taking a year to be marked, only call to 
        # mark times of between 0 and 100 days are returned.
        if 0 <= int(difference[0]) <= 100:
            return int(difference[0])
        

# Function used to check the amount of time (in days) a fiber installation job took to be completed after an 
# 811 call was marked. This function is factoring the using the "811 Marked Date" column as the mark time and 
# the "Job Completed Date" column as the completion time. It takes a row number as input.
# NOTE: if the specified row does not pass both of the checks, the function will return None.
def mark811_to_complete_time_check(row):

    # The 811 mark and job completion dates are assigned to variables
    job_mark811 = row[4]
    job_complete = row[6]

    # The values for both date columns are checked to see if they are empty or not.
    if pd.isnull(job_mark811) != True and pd.isnull(job_complete) != True:

        # If both variables do contain a non-null value, then the amount of elapsed
        # time for the job is calculated by taking the difference of the two times.
        difference = str(job_complete - job_mark811).split(" ")

        # To avoid both outliers and date typos, for example a marked job taking a year to be completed, only 
        # mark to completion times of between 0 and 100 days are returned.
        if 0 <= int(difference[0]) <= 100:
            return int(difference[0])