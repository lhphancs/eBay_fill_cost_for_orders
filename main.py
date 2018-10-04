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
    uppercasedTitle = title.upper()
    for phrase in phrases:
        if phrase.upper() not in uppercasedTitle:
            return False
    return True

def getDictOfMatches(nameToSheetDict, nameToTitleAndCostLocationDict, phrases):
    retDict = {}
    for key in nameToSheetDict:
        matchingRows = []
        sheet = nameToSheetDict[key]
        nameToTitleAndCostLocation = nameToTitleAndCostLocationDict[sheet.title]
        titleCol = nameToTitleAndCostLocation[0][1]
        for rowIndex in range(nameToTitleAndCostLocation[0][0] + 1, sheet.max_row + 1):
            title = sheet.cell(rowIndex, titleCol).value
            if hasAllPhrases(title, phrases):
                matchingRows.append(rowIndex)
        retDict[key] = matchingRows
    return retDict

def printResultOfFind(nameToSheetDict, nameToTitleAndCostLocationDict, dictOfMatches):
    for key in nameToSheetDict:
        titleCol = nameToTitleAndCostLocationDict[key][0][1]
        print( str.format('{}: ', key) )
        for row in dictOfMatches[key]:
            print( '\t' + nameToSheetDict[key].cell(row, titleCol).value)

def getTitlePhrasesAndPrintFinds(nameToSheetDict, nameToTitleAndCostLocationDict):
    phrases = getUserInput.getPhrases()
    dictOfMatches = getDictOfMatches(nameToSheetDict, nameToTitleAndCostLocationDict, phrases)
    printResultOfFind(nameToSheetDict, nameToTitleAndCostLocationDict, dictOfMatches)
    return dictOfMatches


def writeCostToCell(nameToSheetDict, nameToTitleAndCostLocationDict, dictOfMatches, cost):
    for key in nameToSheetDict:
        sheet = nameToSheetDict[key]
        costCol = nameToTitleAndCostLocationDict[key][1][1]
        for row in dictOfMatches[key]:
            sheet.cell(row, costCol).value = cost

def handleWriteCommand(nameToSheetDict, nameToTitleAndCostLocationDict):
    dictOfMatches = getTitlePhrasesAndPrintFinds(nameToSheetDict, nameToTitleAndCostLocationDict)
    cost = getUserInput.getCost()
    confirmResp = getUserInput.getConfirmation("Are you sure you want to write? (Y)es or N(o): ")
    if confirmResp == 'Y':
        writeCostToCell(nameToSheetDict, nameToTitleAndCostLocationDict, dictOfMatches, cost)
    else:
        print("Write was ignored. Returning to menu...")

def handleCommand(excelWb, outputPath, nameToSheetDict, nameToTitleAndCostLocationDict, uppercasedCommand):
    if uppercasedCommand == 'F':
        getTitlePhrasesAndPrintFinds(nameToSheetDict, nameToTitleAndCostLocationDict)
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