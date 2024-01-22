# PART 1 , RUNS MACRO

import win32com.client
import xlwings as xw
import os

# Open the Excel workbook - PLEASE SET UP PATH WHERE XLSM FILE WILL BE SAVED 
wb = xw.Book(r"......")

# Use the 'run' method to execute the macro
wb.macro('FullComb.ImportCSVData').run() 
wb.macro('FullComb.Duplicates').run()
wb.macro('FullComb.SaveCSV').run()
wb.close()

#PART 2, PREPARE OUTLOOK EMAIL

# Path to the folder containing CSV files - PLEASE SET UP PATH WHERE XLSM FILE WILL BE SAVED 
folder_path = r"....."  

# Get the latest CSV file in the folder
latest_csv_file = max([f for f in os.listdir(folder_path) if f.lower().endswith('.csv')],
                      key=lambda x: os.path.getctime(os.path.join(folder_path, x)))

# Full path to the latest CSV file
attachment_path = os.path.join(folder_path, latest_csv_file)

# Create Outlook application object
outlook_app = win32com.client.Dispatch("Outlook.Application")

# Create a new mail item
mail_item = outlook_app.CreateItem(0)  # 0 represents olMailItem

# Set the recipient email address
mail_item.To = "aaa.bbb@gmail.com"

# Set the subject
mail_item.Subject = "VBA Accounts"

# Attach the latest CSV file
mail_item.Attachments.Add(attachment_path)

# Display the new email
mail_item.Display()
