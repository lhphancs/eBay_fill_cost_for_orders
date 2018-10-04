import openpyxl
import csv
import os
from pathlib import Path

import getUserInput

def getReadExcelPath(readDirPath):
    readExcelPath = None
    for file in os.listdir(readDirPath):
        filePath  = os.path.join(readDirPath, file)
        if file.endswith('xlsx'):
            readExcelPath = filePath
            break
    return readExcelPath

def getNameToSheetDict(readWb):
    retDict = {}
    for sheet in readWb:
        retDict[sheet.title] = sheet
    return retDict

def hasAllPhrases(title, phrases):
    for phrase in phrases:
        if phrase not in title:
            return False
    return True

def writeCostToCell(sheet, rowIndex, cost):
    pass
    
def handleFindCommand(nameToSheetDict, nameToTitleAndCostLocationDict):
    phrases = getUserInput.getPhrases()
    cost = getUserInput.getCost()
    for key in nameToSheetDict:
        sheet = nameToSheetDict[key]
        nameToTitleAndCostLocation = nameToTitleAndCostLocationDict[sheet.title]
        titleCol = nameToTitleAndCostLocation[0][1]
        costCol = nameToTitleAndCostLocation[1][1]
        for rowIndex in range(nameToTitleAndCostLocation[0][0] + 1, sheet.max_row + 1):
            title = sheet.cell(rowIndex, titleCol).value
            if hasAllPhrases(title, phrases):
                writeCostToCell(sheet, rowIndex, cost)

    
def handleWriteCommand(nameToSheetDict, nameToTitleAndCostLocationDict):
    pass

def handleCommand(excelWb, outputPath, nameToSheetDict, nameToTitleAndCostLocationDict, uppercasedCommand):
    if uppercasedCommand == 'F':
        handleFindCommand(nameToSheetDict, nameToTitleAndCostLocationDict)
    elif uppercasedCommand == 'W':
        handleWriteCommand(nameToSheetDict, nameToTitleAndCostLocationDict)
        excelWb.save(outputPath)

    elif uppercasedCommand == 'E':
        pass
    else:
        print('Invalid command. Try again.')
    print()

def getLocationOfWord(sheet, word):
    MAX_INDEX_TO_ITER = 20
    for rowIndex in range(1, MAX_INDEX_TO_ITER + 1):
        for colIndex in range(1, MAX_INDEX_TO_ITER + 1):
            cellVal = sheet.cell(rowIndex, colIndex).value
            if cellVal != None and type(cellVal) is str:
                cellVal = cellVal.strip(" '")
                if cellVal.upper() == word.upper():
                    return (rowIndex, colIndex)
    return None

def getNameToTitleAndCostLocationDict(excelWb):
    nameToTitleLocationDict = {}
    for sheet in excelWb:
        titleLocation = getLocationOfWord(sheet, 'TITLE')
        costLocation = getLocationOfWord(sheet, 'COST')
        if titleLocation == None:
            print( str.format("Error...Could not find 'Title' header in sheet: {}", sheet.title) )
            return None
        elif costLocation == None:
            print( str.format("Error...Could not find 'Cost' header in sheet: {}", sheet.title) )
            return None
        else:
            nameToTitleLocationDict[sheet.title] = (titleLocation, costLocation)
        
    return nameToTitleLocationDict

def handlePromptAndResponses(excelWb, outputPath):
    nameToSheetDict = getNameToSheetDict(excelWb)
    nameToTitleAndCostLocationDict = getNameToTitleAndCostLocationDict(excelWb)
    if nameToTitleAndCostLocationDict != None:
        command = None
        while command != 'E':
            uppercasedCommand = getUserInput.getMenuCmd()
            handleCommand(excelWb, outputPath, nameToSheetDict, nameToTitleAndCostLocationDict, uppercasedCommand)

if __name__ == '__main__':
    curDir = Path( os.getcwd() )
    readDirPath = os.path.join(curDir, 'placeExcelFileHere')
    readExcelPath = getReadExcelPath(readDirPath)
    outputPath = os.path.join(curDir, 'result', 'result.xlsx')

    if readExcelPath != None:
        excelWb = openpyxl.load_workbook(readExcelPath)
        handlePromptAndResponses(excelWb, outputPath)
    else:
        print( str.format("No file with extension 'xlsx' in: {}", readDirPath) )
    print("Program has finished...")