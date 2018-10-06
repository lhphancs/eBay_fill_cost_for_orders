import openpyxl
import csv
import os
from pathlib import Path

import getUserInput
import inputMsgs

def getReadExcelPath(readDirPath):
    readExcelPath = None
    for file in os.listdir(readDirPath):
        if file.endswith('xlsx'):
            readExcelPath = os.path.join(readDirPath, file)
            break
    return readExcelPath

def getNameToSheetDict(readWb):
    retDict = {}
    for sheet in readWb:
        retDict[sheet.title] = sheet
    return retDict

"""This also allows 'must not contain'
by entering phrase of '! Apple cider', this means title must not contain Apple cider"""
def meetsAllPhraseConditions(title, phrases):
    uppercasedTitle = title.upper()
    for phrase in phrases:
        uppercasedPhrase = phrase.upper()
        splittedPhrase = uppercasedPhrase.split(' ', 1)
        if splittedPhrase[0] == '!' and len(splittedPhrase) > 1:
            if splittedPhrase[1] in uppercasedTitle:
                return False
        elif uppercasedPhrase not in uppercasedTitle:
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
            if meetsAllPhraseConditions(title, phrases):
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
    writeAmt = 0
    for key in nameToSheetDict:
        writeAmt = writeAmt + len( dictOfMatches[key] )
        sheet = nameToSheetDict[key]
        costCol = nameToTitleAndCostLocationDict[key][1][1]
        for row in dictOfMatches[key]:
            sheet.cell(row, costCol).value = cost
    return writeAmt

def handleWriteCommand(nameToSheetDict, nameToTitleAndCostLocationDict):
    dictOfMatches = getTitlePhrasesAndPrintFinds(nameToSheetDict, nameToTitleAndCostLocationDict)
    cost = getUserInput.getCost()
    confirmResp = getUserInput.getConfirmation("Are you sure you want to write? (Y)es or (N)o: ")
    if confirmResp == 'Y':
        return writeCostToCell(nameToSheetDict, nameToTitleAndCostLocationDict, dictOfMatches, cost)

    else:
        return 0

def handleCommand(excelWb, outputPath, nameToSheetDict, nameToTitleAndCostLocationDict, uppercasedCommand):
    if uppercasedCommand == 'F':
        getTitlePhrasesAndPrintFinds(nameToSheetDict, nameToTitleAndCostLocationDict)
    elif uppercasedCommand == 'W':
        amtWrite = handleWriteCommand(nameToSheetDict, nameToTitleAndCostLocationDict)
        if amtWrite > 0:
            while True:
                try:
                    excelWb.save(outputPath)
                    print( str.format("Write successful. {} cells were written.", amtWrite) )
                    break
                except PermissionError as e:
                    print(e)
                    print("Is the write file, 'result.xlsx', open?")
                    print("Press enter to try again...")
                    input('(Note: Changes made to the excel file outside of this program will not be saved)\n')
        else:
            print("Write was ignored. Returning to menu...")
    elif uppercasedCommand == 'H':
        print(inputMsgs.menuHelp) 

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
        uppercasedCommand = None
        while uppercasedCommand != 'E':
            uppercasedCommand = getUserInput.getMenuCmd()
            handleCommand(excelWb, outputPath, nameToSheetDict, nameToTitleAndCostLocationDict, uppercasedCommand)

if __name__ == '__main__':
    curDir = Path( os.getcwd() )
    readDirPath = os.path.join(curDir, 'placeSingleExcelFileHere')
    readExcelPath = getReadExcelPath(readDirPath)
    outputPath = os.path.join(curDir, 'result', 'result.xlsx')

    if readExcelPath != None:
        excelWb = openpyxl.load_workbook(readExcelPath)
        handlePromptAndResponses(excelWb, outputPath)
    else:
        print( str.format("No file with extension 'xlsx' in: {}", readDirPath) )
    print("Program has finished...")