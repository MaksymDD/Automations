## PROJECT: REMOVING DUPLICATES 

Hello, welcome to this automation project rooted in real-world work experiences—tasks routinely carried out by middle offices worldwide. 
These are the types of activities that can be easily automated, resulting in significant time savings.

An important note: you may find that this project is somewhat intricate, and certain aspects could be optimized for greater efficiency. 
However, as mentioned earlier, this project is a reflection of my work experience. I've intentionally added a touch of complexity to my tasks to make the process more enjoyable.

# Project Description

The purpose of this project is to automate a process inspired by one of my past tasks. The process involves handling a CSV file with the following columns:

- **Column A:** Account numbers
- **Column B:** The location of accounts
- **Column D:** Number of shares held by the accounts

The task at hand is to identify and consolidate duplicated accounts, summing up their total positions. Once this sorting is complete, the processed file needs to be saved in a new folder and subsequently sent to our colleagues on the team.

# Technologies Used

To automate this process, I employ a combination of Python and VBA. Specifically:

- **Python:** Primarily used to execute VBA code within Excel and generate new email messages via Outlook.
- **VBA:** Responsible for tasks such as copying data from CSV files, iterating through duplicated accounts, and subsequently saving the new files in specified directories.

# Workflow

First, let me provide part of python that is responsible for opening Excel file and running macro. 

```python
import win32com.client
import os

import xlwings as xw

# Open the Excel workbook
wb = xw.Book(r"C:\Offile Desktop\Study Material\Python\Projects\VBA+Python\Data\VBA_dup.xlsm")

# Use the 'run' method to execute the macro
wb.macro('FullComb.ImportLatestCSVData').run()
wb.macro('FullComb.ConsolidateAndDeleteDuplicates').run()
wb.macro('FullComb.SaveAsCSVWithDynamicName').run()
wb.close()


