# Knoxville Utilities Board Spreadsheet Automation

During my internship in Knoxville Utilities Board’s Fiber Department, I developed a Python script to automate daily updates of their fiber installation job data in Excel. The script archived jobs that were completed for over 90 days and generated multiple analysis sheets containing many time-based analysis metrics. (All data used is sample data, not real customer information.)
    
**DISCLAIMER: None of the data in the version of the Excel file or Python script present contains real customer data as I have replaced it with sample data.**

<p align="center">
<img width="922" height="614" alt="KUB Python Automation" src="https://github.com/user-attachments/assets/a44d78d6-3f58-454d-abca-0496a1d1ee83" />
</p>

### Link to Excel File:
**Pre-Update**: https://1drv.ms/x/c/4815cd31ccfbf255/EUDynAK9XgVMsIeZp3PYVR8BuzmbZUceBGYD-fdtmKhoOQ?e=8d6Fnz

**Post-Update**: https://1drv.ms/x/c/4815cd31ccfbf255/EYEuWrIxzSpKmD1ho3LHKtcBz6fvTyiXVSTDSqtXDD_tEw?e=cC8LWz
__________________________________________________________________________________________________________________________________________________________________
### Script Structure
I used a modular programming approach, organizing each main component of the script into separate files. This made debugging easier and ensures future developers can quickly identify and update specific parts as the Excel file evolves.

#### Main Process Files:

_main.py_ - Main file of the script that is responsible for running every component of the automation process.

_check_backup_directory_and_run_time_log.py_ - File that contains functionality for verifying that the current month's backup directory and script run time logs
                                               exist. If they do not exist, then the file creates a directory and log for the current month.

_update_main_and_archive.py_ - File that imports rows from the main sheet to the archive sheet and subsequently removes imported rows from the main sheet. This 
                               process operates through the Pandas library.

_update_area_metrics.py_ - File that updates the metrics located in the work area analysis sheet. This process operates through the Pandas library.

_update_month_metrics.py_ - File that updates the metrics located in the month-by-month analysis sheet. This process operates under the Pandas library.

_restore_main_and_archive.py_ - File that restores the advanced formatting of the main and archive sheets. This process operates under the OpenpyXL library.

_restore_analysis_sheets.py_ - File that restores the advanced formatting of both analysis sheets. This process operates under the OpenpyXL library.

#### Function Files:

_FUNCTIONS_update_main_and_archive.py_ = File that contains all functions used in _update_main_and_archive.py_ file.

_FUNCTIONS_update_area_metrics.py_ = File that contains all functions used in _update_area_metrics.py_ file.

_FUNCTIONS_month_metrics.py_ = File that contains all functions used in _update_month_metrics.py_ file.

_FUNCTIONS_time_checks.py_ = File that contains multiple time check functions that are used in the _update_area_metrics.py_ and _update_month_metrics.py_ files.

_FUNCTIONS_restore_main_and_archive.py_ = File that contains all functions used in _restore_main_and_archive.py_ file.

_FUNCTIONS_restore_analysis_sheets.py_ = File that contains all functions used in _restore_analysis_sheets.py_ file.

_FUNCTIONS_cell_color.py_ = File that contains all cell fill color assignments that are used in the file restoration process.
__________________________________________________________________________________________________________________________________________________________________
