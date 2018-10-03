import openpyxl
import csv
import os
from pathlib import Path

import inputMsgs

def getReadExcelPath(readDirPath):
    readExcelPath = None
    for file in os.listdir(readDirPath):
        filePath  = os.path.join(readDirPath, file)
        if file.endswith('xlsx'):
            readExcelPath = filePath
            break
    return readExcelPath

def getSheetDict(readWb):
    retDict = {}
    for sheet in readWb:
        sheetName = sheet.title
        print(sheetName)

def copyExcel(readExcelPath, outputPath):
    return ''

def handlePromptAndResponses(writeExcelWb):
    pass

if __name__ == '__main__':
    curDir = Path( os.getcwd() )
    readDirPath = os.path.join(curDir, 'placeExcelFileHere')
    readExcelPath = getReadExcelPath(readDirPath)
    outputPath = os.path.join(curDir, 'result') + 'result.xlsx'

    if readExcelPath != None:
        writeExcelWb = copyExcel(readExcelPath, outputPath)
        handlePromptAndResponses(writeExcelWb)
    else:
        print( str.format("No file with extension 'xlsx' in: {}", readDirPath) )
