## PROJECT: REMOVING DUPLICATES 

Hello, welcome to this automation project rooted in real-world work experiences—tasks routinely carried out by middle offices worldwide. 
These are the types of activities that can be easily automated, resulting in significant time savings.

An important note: you may find that this project is somewhat intricate, and certain aspects could be optimized for greater efficiency, however, I've intentionally added a touch of complexity to my tasks to make the process more enjoyable.

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
  
*Please note, for this project, both VBA and Python script are using the latest avilible file in our folders*

Below snippet from VBA 

```vba
    ' Loop through files in the folder
    Do While fileName <> ""
        currentFile = folderPath & fileName

        ' Check if the file is a CSV file and if its date is later than the current latest date
        If UCase(Right(fileName, 4)) = ".CSV" And FileDateTime(currentFile) > latestDate Then
            latestDate = FileDateTime(currentFile)
            GetLatestCSVFile = currentFile
        End If

        ' Get the next file
        fileName = Dir
    Loop
End Function
```

& Python:

```python
latest_csv_file = max([f for f in os.listdir(folder_path) if f.lower().endswith('.csv')],
                      key=lambda x: os.path.getctime(os.path.join(folder_path, x)))
```


# Workflow

**PART 1**

First, let me provide a snippet of Python code that opens an Excel file and runs a macro.

```vba
import win32com.client
import os
import xlwings as xw #

#Path to the Excel file with VBA.
wb = xw.Book(r"C:\Offile Desktop\Study Material\Python\Projects\VBA+Python\Data\VBA_dup.xlsm")

# Execution of macro with command 'run'
wb.macro('FullComb.ImportCSVData').run() 
wb.macro('FullComb.Duplicates').run()
wb.macro('FullComb.SaveCSV').run()
wb.close()
```


**PART 2**

Now let's us move to VBA script. As you can see above, my VBA is divided in to 3 different Sub:

1. Sub that import data from folder with CSV file and duplicated accounts 

```vba
'Looping for csv file
    If latestCSVFile <> "" Then
        ' Clear existing data in the worksheet
        ws.Cells.Clear

        ' Import data from the csv
        For columnCounter = 1 To 1
            ImportColumnFromCSV latestCSVFile, ws, columnCounter
        Next columnCounter
```

2. Sub that capturing duplicated accounts and total its position .

```vba
  ' Loop through each row of data
    For i = 2 To lastRow ' Assuming your headers are in row 1
        
        ' Check if the account number is in the dictionary
        If acctDict.Exists(ws.Cells(i, "A").Value) Then
            ' If it exists, add the shares to the existing total
            acctDict(ws.Cells(i, "A").Value) = acctDict(ws.Cells(i, "A").Value) + ws.Cells(i, "C").Value
        Else
            ' If it doesn't exist, add the account number and shares to the dictionary
            acctDict(ws.Cells(i, "A").Value) = ws.Cells(i, "C").Value
        End If
        
        ' Clear the shares in the current row
        ws.Cells(i, "C").Value = 0
    Next i
    
    ' Loop through the dictionary and update the shares in the worksheet
    For i = 2 To lastRow
        If acctDict.Exists(ws.Cells(i, "A").Value) Then
            ws.Cells(i, "C").Value = acctDict(ws.Cells(i, "A").Value)
            ' Remove the account from the dictionary to avoid duplicattes
            acctDict.Remove ws.Cells(i, "A").Value
        End If
    Next i
    
    ' Clean up and release the dictionary.
    Set acctDict = Nothing
    
    ' Delete duplicated rows.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = lastRow To 2 Step -1
        If Application.WorksheetFunction.CountIf(ws.Range("A2:A" & i), ws.Cells(i, "A").Value) > 1 Then
            ws.Rows(i).Delete
        End If
    Next i
End Sub
```

*P.S. In this example, I've chosen to capture duplicated accounts based on a single criterion — ID number. However, if needed, we can update the VBA code to capture additional data, such as ISIN.*

3. And last Sub that saves our new file without duplicates in new folder, adding also todays date.

```vba
       ' File name + today's date
    saveFileName = "Accounts " & Format(Date, "yyyymmdd") & ".csv"

    ' Save new CSV
    ActiveWorkbook.SaveAs fileName:=saveFolderPath & saveFileName, _
        FileFormat:=xlCSV, CreateBackup:=False
```

**PART 3**

Once we have executed our VBAs, let's proceed to the final step — preparing a new email message in Outlook with our new CSV file, free of duplicated accounts. Thos part will be done using Python script, and win32com module. 

```python
# Get the latest CSV file in the folder
latest_csv_file = max([f for f in os.listdir(folder_path) if f.lower().endswith('.csv')],
                      key=lambda x: os.path.getctime(os.path.join(folder_path, x)))

# Full path to the latest CSV file
attachment_path = os.path.join(folder_path, latest_csv_file)

# Create Outlook application object
outlook_app = win32com.client.Dispatch("Outlook.Application")

# Create new mail item
mail_item = outlook_app.CreateItem(0)  # 0 represents olMailItem

# Set the recipient email address
mail_item.To = "aaa.bbb@gmail.com"

# Set the subject
mail_item.Subject = "VBA Accounts"

# Attach the CSV file
mail_item.Attachments.Add(attachment_path)

# Display email
mail_item.Display()
```

