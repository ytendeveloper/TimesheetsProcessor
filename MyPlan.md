# TimesheetsProcessor

> **Version**: 1.3


## 1. Problem Statement
Employees email timesheet attachments to a designated email address. We need to:

0. **Init** by reading config/config.yaml file
    - START_DATE param is of MM-DD-YYYY
    - END_DATE param is of MM-DD-YYY (default to current date)
    - EMPLOYEES_PARENT_DIR_PATH param is path to employees folder parent path (default to working dir)
    - EMPLOYEES_LIST_FILE is the file that contains list of employees with Employee Name, Employee Email ID and Project Name (if param value is not set look for file in config/EmployeesList.md)
    - EMAIL_BATCH_SIZE is number of emails to retrieve at a time (default to 5)
    - PROCESSING_FOLDER_NAME is temporary folder where attachments are downloaded to (default to 'downloaded')

1. **Download (Job-1-Download)** their attachments from gmail into organized local folders
    - Read EmployeesList file
    - Loop through one employee at a time
        - Search gmail for emails from the employees email id, has attachments, from START_DATE, to END_DATE (only download EMAIL_BATCH_SIZE number of emails)
        - In folder path given by [ config param value of EMPLOYEES_PARENT_DIR_PATH] , Create a folder with with [employee name + "_" + Project Name (remove spaces and dots and special chars) ]
        - Create a download-{batch-run#}.csv file
        - from the emails, Download attachments to PROCESSING_FOLDER_NAME param value named subfolder in employee folder
        - enter following information to the earlier created download csv file. :  gmail's unique ID for the email, number of attachments, date & time of the email.
    - name this job's executable as Job-1-Download

2. **Parse (Job-2-Parse)** downloaded timesheets files to extract structured data (name, hours, dates, client/project)
    - Loop through subfolders in parent directory path = (value of EMPLOYEES_PARENT_DIR_PATH config param)
    - In each employee directory parse timesheet files that are in "downloaded" folder and extract structured data (name, hours, dates, client/project)
    - if its weekly timesheet then Hours should be Sun-Sat weekly  OR if its monthly timesheet should be from start to end of the month.
    - If the hours cannot be parsed, move the file to "unableToParse" folder within the employee folder (create the folder if it does not exist)
    - In the employee folder if the file does not exist create an excel file with employee name + project name
    - In the above created excel file create/update one tab per year and prepopulate with weeks or months date ranges
    - with the dates extracted from the timesheet
        - update the excel ( in the appropriate year's tab and in the appropriate week/month row  ) with the hours worked and the project name
    - if client/project value is empty find something unique that can be used as the value. employee could be working on multiple projects at same time
    - rename the timesheet file with start and end dates and move from "downloaded" to "processed" folder
3. **MoveToProcessedTag (Job-3-MoveToProcessedTag)** move processed files in gmail to Processed tag
    - Loop through subfolders in parent directory path = (value of EMPLOYEES_PARENT_DIR_PATH config param)
    - In each employee directory look for download csv , loop through the "gmail_message_id" values and move the emails with the respective IDs from 'inbox' to "claudeprocessed" tag/folder
    - Never delete the emails
    - once all the emails in the csv are  moved then move the csv file to 'processed' folder

## Notes
    - DON'T Ever delete emails
    - Create separate executable for each of the major modules ( #0 module is required by the other modules)
    - Document the generated code appropriately, within the code
    - log all major events to a log file and add mechanism to rotate logs and store logs in logs/ folder
