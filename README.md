## VBA_code
FILE 1. IE SCRIPT EXAMPLE. Daily report automation to get the pending active workload. Internet Explorer script that goes into a website, inputs information, downloads a file and formats the file. It uses conversion tables and pivot tables.
   
FILE 2. SAP SCRIPT. VIM REPORT. Report that connects to 5 diff SAP systems to get the current overview of VIM (Vendor Invoice Management). It downloads several files, puts them together and formats all data. It updates a Tableau source file and incorporates a macro to autosend an email. 


FILE 3. SAP SCRIPT. REPORT RUN WEEKLY TO GET NUMBER OF POSTINGS. Report that connects to 5 diff SAP systems to get daily number of postings done by each analyst in the last week. It downloads several files, puts them together and formats all data. It shows final data in pivot tables.

FILE 4. EXCEL TASK. THRESHOLD. EMAIL DRAFT VERSION. Monthly task to run an analysis from a sheet where data comes from an external database (it was input manually). The idea was selecting x% of invoices to fully verify again and make sure charges were accurate. Several macros embedded inside to allow human interaction between tasks. It only selects vendors whose % weight is greater than x. Both x% mentioned before can be adjusted in "Selection Summary" sheet. The macro finds out the number of invoices to be selected per vendor and region for a relevant selection. after some input requested from the person in charge of the tasks, it selects individual invoices to be analysed from the "Raw" sheet. It incorporates a macro to create an email draft and revise content before sending it. 

FILE 5. SAP &IE SCRIPTS + EXCEL TASK. MONTHLY DATA RECONCILIATION. Monthly task to reconcile & analyse data. The purpose is to compile data monthly and maintain a format to create a database.  Several macros embedded inside to allow human interaction between tasks. It starts from "Raw data" sheet where data comes from an external database (it was input manually). It goes from copying data to applying formulas, connecting to SAP, connecting to a website through Internet Explorer and formatting it all. It allows human interaction as requested by the person in charge of the task to achieve supervision.

FILE 6. INVOICE PROCESSING AUTOMATION TOOL. SAP scripts to automate daily tasks related to invoice processing. From launching some duplicate checks in SAP to finding the necessary coding information (cost objects) or inputing information in SAP directly. It has an interesting function to input data to SAP. It also contains 2 of the methods I used to use to extract data from SAP, downloading a file and going through the results array

FILE 7. SAP SCRIPT. TASK TO REBILL COST. Monthly task to rebills costs between internal affiliates. Several macros embedded inside to allow human interaction between tasks. It starts downloading data from SAP several times to create the base data. This task is done 2 times, with buttons 1 and 3. Buttons 2 and 4 are used to copy-paste values and output the format needed.

FILE 8. SAP SCRIPT. REPORT TOOL. Daily task to be carried. Macro built in order to gather data and build a report in order to assess and allocate workload.

FILE 9. SAP &IE SCRIPTS + EXCEL TASK. MONTHLY DATA RECONCILIATION 2. Monthly task to reconcile & analyse data. The purpose is to compile data monthly and maintain a format to compare and reconcile 2 systems. It downloads data coming from an external database using a date range imput by the user. It goes from copying data to applying formulas, connecting to SAP, connecting to a website through Internet Explorer and formatting and comparing it all. It provides a usmmary of the unreconciled findings.

FILE 10. SAP SCRIPT. PROCESS TO REDUCE MANUAL TASKS. SAP scripts to download data & format it to reduce manual tasks from users. The purpose is running the tool every week to automate daily tasks related to invoice processing. the intention is finding the necessary coding information (cost objects).

FILE 11 & FILE 12. Used together to automate a task. FILE 11. SAP Script. Yellow note part 1. downloads data from SAP and formats it to upload it using file 12. FILE 12. RFC. Yellow note part 2. The only Access file of the group. This file was pure gld at the time. We used Open Text to view invoices in SAP. This file input the information provided as a note (yellow note) inside the Open text files of the invoices. It uses RFC (Remote Function Calls), some SQL and several vba functions. 

FILE 13. SAP Script. Monthly report. Monthly task to check the current volumes on the refineries, find imbalances and perform a reconciliaition. The result is formatted on to a table found on the "RESULT" sheet.

FILE 14. EXCEL TASK. MONTHLY COMPILATION AND RECONCILIAITION OF DATA. Monthly task that saves hours of data collection manual task from excel spreadsheets. That is the most interesting part of this file using arrays to compile all data quickly. It stores in arrays all individual volumes per material and dates. It then formats data, apply formulas and calculations, creates pivot tables to get totals per type of material and checks for differences and potentiall issues on the calculations. 
