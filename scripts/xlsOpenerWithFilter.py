import os
import win32com.client

import criteria as crit;

filePath= "C:/Users/avtom/OneDrive/Desktop/chevrolet center/data/exmp.xlsm"

if os.path.exists(filePath):
    crit.setCriteria(9899)
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Workbooks.Open(Filename=filePath, ReadOnly=1)
    xl.Application.Run("OpenAndFilterExcelFile")
    xl.Visible = True  # Make Excel visible to the user
else:
    print("File not found")





