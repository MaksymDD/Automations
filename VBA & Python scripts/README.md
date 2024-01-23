In this folder, you can find the file VBA_Dup.xlsm and the macro automate_gt.py. Feel free to review or test them.

## Execution Instructions

To test this project on your own, you will need the following:

1. Save automate_gt.py
2. VBA_Dup.xlsm
3. CSV file with duplicated accounts.
4. Download xlwings if required. 

Make sure to update the paths accordingly. Inside the Python file (`automate_gt.py`), provide the path to your locally saved `VBA_Dup.xlsm` on your computer.

```python
import win32com.client
import xlwings as xw
import os

# Open the Excel workbook - PLEASE SET UP PATH WHERE XLSM FILE WILL BE SAVED 
wb = xw.Book(r"PLEASE PROVIDE PATH TO xlsm. FILE")
```

Accordingly in VBA:


```vba
'PART 1
Sub ImportCSVData()

    Dim ws As Worksheet
    Dim csvFolderPath As String
    Dim latestCSVFile As String
    Dim columnCounter As Integer

    ' Worksheet with accounts
    Set ws = ThisWorkbook.Sheets("Accounts")

    ' Path to CSV file - TO BE EDITED
    csvFolderPath = "PLEASE PROVIDE PATH TO CSV FILE WITH DUPLICATED ACCOUNTS"


PART 2
'PART 3, SAVE AS CSV

Sub SaveCSV()
    Dim saveFolderPath As String
    Dim saveFileName As String

    ' Folder path for results - TO BE EDITED
    saveFolderPath = "PLEASE PROVIDE PATH WHERE YOU WOULD LIKE TO SAVE YOUR NEW CSV FILE WITHOUT DUPLICATES"

    ' File name with "Accounts" and today's date
    saveFileName = "Accounts " & Format(Date, "yyyymmdd") & ".csv"

    ' Save the workbook as CSV with the dynamic name
    ActiveWorkbook.SaveAs fileName:=saveFolderPath & saveFileName, _
        FileFormat:=xlCSV, CreateBackup:=False
End Sub
```

## Make sure you have your CSV file with diplicated accounts saved on your computer. This file could be found in folder named *'CSV with duplicated accounts'*

