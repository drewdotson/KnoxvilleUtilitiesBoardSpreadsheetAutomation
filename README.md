# Knoxville Utilities Board Spreadsheet Automation
<img width="450" height="802" alt="Screenshot 2025-10-18 004430" src="https://github.com/user-attachments/assets/2cc56973-2279-4dd5-92f9-1d97b69af81f" />

### Link to Excel File:
**Pre-Update**: https://1drv.ms/x/c/4815cd31ccfbf255/EUDynAK9XgVMsIeZp3PYVR8BuzmbZUceBGYD-fdtmKhoOQ?e=8d6Fnz

**Post-Update**: https://1drv.ms/x/c/4815cd31ccfbf255/EYEuWrIxzSpKmD1ho3LHKtcBz6fvTyiXVSTDSqtXDD_tEw?e=cC8LWz
__________________________________________________________________________________________________________________________________________________________________
### The task
During my time as an intern in Knoxville Utilities Board's fiber department, I was tasked with creating a script to conduct daily automated updates of their Excel
file that contains data related to all of their fiber installation jobs. I was given three main tasks to accomplish with this script:
  1. The script must move installation jobs that have been completed for 90 days from the main installation job data sheet to the archive sheet.
  2. Create an analysis sheet that captures specific installation job time metrics for every work area that KUB operates in and have it updated daily.
  3. Create an analysis sheet that caputres specific installation job time metrics for every month that is present in the dataset and have it updated daily.
    
**DISCLAIMER: None of the data in the version of the Excel file or Python script present contains real customer data as I have replaced it with sample data.**
__________________________________________________________________________________________________________________________________________________________________
### Script Structure
I chose to take a sort of modular programming approach when creating the Script's structure, the easiest way to build and maintain this script was to split the
script up to where each main component of the script was isolated in its own file. This allowed me to easily keep up with my errors as I coded the Script and it
allows future programmers who will inevitably modify this script at a later date to discern where a potential issue could arise as the Excel file changes.

#### Files:

_main.py_ - Main file of the script that is responsible for running every component of the automation process.

_check_backup_directory_and_run_time_log.py_ - File that contains functionality for verifying that the current month's backup directory and script run time logs
                                               exist. If they do not exist, then the file creates a directory and log for the current month.

_update_main_and_archive.py_ - File that contains all functionality responsible for importing rows from the main sheet to the archive sheet and subsequently
                               removing imported rows from the main sheet. This file operates through the Pandas library.

_update_area_metrics.py_ - File that contains all functionality responsible for updating the metrics located in the work area analysis sheet. This file operates
                           through the Pandas library.

_update_month_metrics.py_ - File that contains all functionality responsible for updating the metrics located in the month-by-month analysis sheet. This file
                            operates under the Pandas library.

_checks.py_ - File that contains time check functions used in both the work area and month-by-month metrics files.

_restore_main_and_archive.py_ - File that contains all functionality responsible for restoring the advanced formatting of the main and archive sheets. This file
                                operates under the OpenpyXL library.

_restore_analysis_sheets.py_ - File that contains all functionality responsible for restoring the advanced formatting of both analysis sheets. This file operates
                               under the OpenpyXL library.
__________________________________________________________________________________________________________________________________________________________________
