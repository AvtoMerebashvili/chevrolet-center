import os
import win32com.client

import criteria as crit;

filePath= "C:/Users/avtom/OneDrive/Desktop/chevrolet center/data/clients-table.xlsm"

def openExcelAndFilter(card_id):   
    if os.path.exists(filePath):
        crit.setCriteria(card_id) # set id of the card
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Workbooks.Open(Filename=filePath, ReadOnly=1) #open excel
        xl.Application.Run("OpenAndFilterExcelFile") # run vba macro of excel (it should be excisted in excel)
        xl.Visible = True  # Make Excel visible to the user
    else:
        print("File not found")





# openExcelAndFilter("1312")