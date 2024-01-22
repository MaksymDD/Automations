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

First, let me provide a snippet of Python code that opens an Excel file and runs a macro.

```python
import win32com.client
import os
import xlwings as xw #

#Path to the Excel file with VBA.
wb = xw.Book(r"C:\Offile Desktop\Study Material\Python\Projects\VBA+Python\Data\VBA_dup.xlsm")

# Execution of macro with comman 'run'
wb.macro('FullComb.ImportCSVData').run() 
wb.macro('FullComb.Duplicates').run()
wb.macro('FullComb.SaveCSV').run()
wb.close()
```

Now let's us move to VBA script. As you can see above, my VBA is divided in to 3 different Sub:

1. Sub that import data from folder with CSV file and duplicated accounts 

```vba
'Check if a CSV file was found
    If latestCSVFile <> "" Then
        ' Clear existing data in the worksheet
        ws.Cells.Clear

        ' Import data from the latest CSV file column by column
        For columnCounter = 1 To 1
            ImportColumnFromCSV latestCSVFile, ws, columnCounter
        Next columnCounter
```




