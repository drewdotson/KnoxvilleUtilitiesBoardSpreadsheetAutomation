
# The purpose of this file is to house time check functions that are used by both the area metrics and the month-by-month metrics analysis sheets.

import pandas as pd
from datetime import datetime

# Function used to check the amount of time (in days) a fiber installation job took to be completed. This function is factoring the using the "Drop 
# Installation Date" column as the starting time. It uses a row number as input.
# NOTE: if the specified row does not pass every one of the checks, the function will return None.
def start_to_complete_time_check(row):

    # The value of the "Status" column in the specified row is saved to a variable
    status = row[14]

    # The "Status" column's value is checked to see if the value equals the string "completed" (case-insensitive).
    if type(status) == str and status.lower() == "completed": 

            # The values of the "Drop Installation Date" and "Job Completed Date" columns in the specified row are saved to variables.
            job_start = row[1]
            job_complete = row[6]

            # The data types of both date values are check to see if they are either pandas timestamp or datetime objects.
            if type(job_start) == pd._libs.tslibs.timestamps.Timestamp or type(job_start) == datetime:
                if type(job_complete) == pd._libs.tslibs.timestamps.Timestamp or type(job_complete) == datetime:

                    # If both variables are one of the two data types, then the job's completion time is calculated and saved to a variable.
                    # the two times.
                    difference = str(job_complete - job_start).split(" ")

                    # To avoid both outliers and date typos, for example a job taking a year to complete, only job completion times of between 0 
                    # and 100 days are returned.
                    if 0 <= int(difference[0]) <= 100:
                        return int(difference[0])


# Function used to check the amount of time (in days) a fiber installation job took to be marked after an 811 call was made. This function is 
# factoring the using the "811 Called in Date" column as the call time and  the "811 Marked Date" column as the mark time. It uses a row number 
# as input.
# NOTE: if the specified row does not pass both of the checks, the function will return None.
def call811_to_mark811_time_check(row):

    # The values of the "811 Called in Date" and "811 Mark Date" columns in the specified row are saved to variables.
    job_call811 = row[2]
    job_mark811 = row[4]

    # The data types of both date values are check to see if they are either pandas timestamp or datetime objects.
    if type(job_call811) == pd._libs.tslibs.timestamps.Timestamp or type(job_call811) == datetime:
        if type(job_mark811) == pd._libs.tslibs.timestamps.Timestamp or type(job_mark811) == datetime:
 
            # If both variables are one of the two data types, then the job's call to mark time is calculated and saved to a variable.
            difference = str(job_mark811 - job_call811).split(" ")

            # To avoid both outliers and date typos, for example a call taking a year to be marked, only call to mark times of between 0 and 
            # 100 days are returned.
            if 0 <= int(difference[0]) <= 100:
                return int(difference[0])
        

# Function used to check the amount of time (in days) a fiber installation job took to be completed after an  811 call was marked. This function 
# is factoring the using the "811 Marked Date" column as the mark time and the "Job Completed Date" column as the completion time. It uses a row 
# number as input.
# NOTE: if the specified row does not pass both of the checks, the function will return None.
def mark811_to_complete_time_check(row):

    # The values of the "811 Mark Date" and "Job Completed Date" columns in the specified row are saved to variables.
    job_mark811 = row[4]
    job_complete = row[6]

    # The data types of both date values are check to see if they are either pandas timestamp or datetime objects.
    if type(job_mark811) == pd._libs.tslibs.timestamps.Timestamp or type(job_mark811) == datetime:
        if type(job_complete) == pd._libs.tslibs.timestamps.Timestamp or type(job_complete) == datetime:
 
            # If both variables are one of the two data types, then the job's mark to completion time is calculated and saved to a variable.
            difference = str(job_complete - job_mark811).split(" ")

            # To avoid both outliers and date typos, for example a marked job taking a year to be completed, only 
            # mark to completion times of between 0 and 100 days are returned.
            if 0 <= int(difference[0]) <= 100:
                return int(difference[0])
