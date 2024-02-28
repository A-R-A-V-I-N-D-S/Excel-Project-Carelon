1) Create an excel sheet with 2 workbooks - "CTM Dettails" & "Checklist".
2) In the first workbook CTM Details, copy and paste all the job details from Workload Manager.
3) a. For CTM200 server paste the following in the Application field in the workload manager to get data of all the jobs - **ASCS, ASCS_PROD, BLCS_AGENT_TYPE_INFO**
   b. For GA_DIST_PROD server paste the following in the Application field in the workload manager to get data of all the jobs - **BLCS_PRODUCER_SBO_JOBS_PROD, ASCS, ASCS_PROD, BLCS_PROD, BLCS_AUDIT_OF_CONTRACTS_REPORT_PROD**
4) In the second workbook Checklist, copy and paste all the rows for the particular **order date** from our daily checklist monitoring sheet till the column named 'Shift'.
5) Give the location and name of the sheet in the code where we're reading and writing the file.
6) And then, once you run the code, the program will ask for the order date of the jobs which you need to enter. 
7) Finally code will write the timings and status for the job's that are entered in the CTM Details workbook.
