# native_win32com.py  - Version 1.0 20th September 2023
#
# Example of using the natively shipped win32com library to interact with Excel for reading and writing.   
# This solution does not require additional libraries and allows access to all excel functionallity. 
#
# Usage: Copy script and .xlsl file into your Abaqus Workfolder and run script there. 
# 
# By Ebbe Smith 2023
###################################################################################################################

from win32com.client import constants
import win32com, os

def convertCellDef(cellDef):
    column = ord(cellDef[0].lower())-96
    row = int(cellDef[1:])
    return (column, row)

def readFromWB(wb, sheet, startPos, endPos): 

    extractedData = []
    for thisRow in range(startPos[1], endPos[1]+1):
        for thisColumn in range(startPos[0], endPos[0]+1):
            value = wb.Sheets(sheet).Cells(thisRow, thisColumn).Value
            if value != None:
                extractedData.append(float(value))
    return extractedData

def writeToWB(wb, sheet, startPos, endPos, data): 
    
    numElements = (endPos[0]-startPos[0]+1)*(endPos[1]-startPos[1]+1)
    if numElements != len(data):
        print("CellArea and len(data) does not correspond, exiting") 
        return False 
      
    index = 0
    for thisRow in range(startPos[1], endPos[1]+1):
        for thisColumn in range(startPos[0], endPos[0]+1):
            wb.Sheets(sheet).Cells(thisRow, thisColumn).Value = data[index]   
            print(index, data[index])
            print(thisRow, thisColumn)
            index += 1
    return True
    
def main():
    # Define output filename and Values
    excelFile = os.getcwd()+"\\"+"Bok15.xlsx"
    
    # Access the MS Excel object 
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = xl.DisplayAlerts =  False
    
    try: 
        # Access Workbook
        wb = xl.Workbooks.Open(excelFile)
        
        # Read Data - 2D Array 
        dataFromExcel = readFromWB(wb, "Ark1", convertCellDef("D7"), convertCellDef("G10")) 
        print(dataFromExcel)
        
        # Write Data - 2D Array
        dataFromExcel = [x*3 for x in dataFromExcel]            
        writeToWB(wb, "Ark1", convertCellDef("D13"), convertCellDef("G16"), dataFromExcel)
  
    except: 
        print "Error during Reading or Writing :: Shutting Down!" 
    
    # Save changes and close object
    wb.Close(SaveChanges=1)
    xl.Quit()
    
    return True

if __name__ == '__main__':
    main()