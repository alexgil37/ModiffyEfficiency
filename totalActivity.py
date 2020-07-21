import openpyxl
import json
import os
import sys
import xlsxwriter


instrumentModel = '2360/43-93'

filesWithNoMatchingSN = list()
sheetsOfFilesWithNoMatchingSN = list()
counts = list()
backgroundCounts = list()
invalidSheets = list()


def main(path, savePath):
    if os.path.isfile(savePath):
        os.mkdir(savePath)

    # create excel QC file
    QCworkbook = xlsxwriter.Workbook(savePath + '\\' + 'Total_Activity.xlsx')
    QCworksheet = QCworkbook.add_worksheet()

    # create columns with headers
    QCworksheet.write(0, 0, 'File Name')
    QCworksheet.write(0, 1, 'Sheet Name')
    QCworksheet.write(0, 2, 'Survey Number')
    QCworksheet.write(0, 3, 'Gross Counts')
    QCworksheet.write(0, 4, 'Background Counts')
    QCworksheet.write(0, 5, 'Efficiency')


    def resource_path(relative_path):
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.dirname(__file__)))
        return os.path.join(base_path, relative_path)


    def getListOfFiles(dirName):
        listOfFile = os.listdir(dirName)
        allFiles = list()

        for file in listOfFile:
            fullPath = os.path.join(dirName, file)
            if os.path.isdir(fullPath):
                allFiles = allFiles + getListOfFiles(fullPath)
            else:
                allFiles.append(fullPath)

        return allFiles


    def find_instrument_model_cell(currentSheet):
        for row in range(1, 30):
            for column in "GHIJKLMNOPQRSTUVWXYZ":  # Here you can add or reduce the columns
                modelCell = "{}{}".format(column, row)
                if currentSheet[modelCell].value == instrumentModel:
                    print("the row is {0} and the column {1}" .format(row, column))
                    print(currentSheet[modelCell].value)
                    print(modelCell)

                    return [row, column, modelCell]

        return [0, 0, None]


    def check_for_BettaGamma(num):
        found = 0
        for row in range(1, 30):
            for column in "GHIJKLMNOPQRSTUVWXYZ":

                modelVal = currentSheet[column +str(row)].value
                if modelVal != "Beta-Gamma":
                    continue

                found += 1
                if found == num:
                    return [row, column]

        return [None, None]


    def find_height(currentSheet, topCol, topRow):
        height = 0
        for row in range(0, 20):
            value = currentSheet[topCol + str(topRow+row)].value
            if value != None:
                height += 1

            else :
                return height


    def find_instrument_sn_cell(instModelRow, instModelColumn):
        snRow = str(int(instModelRow) + 1)
        snCol = instModelColumn
        snCell = currentSheet[snCol + snRow]

        return snCell


    def find_cal_due_date(instModelRow, instModelColumn):
        calRow = str(int(instModelRow) + 2)
        calCol = chr(ord(instModelColumn))
        calCell = currentSheet[calCol + calRow]

        return calCell


    def find_instrument_efficiency(instModelRow, instModelColumn):
        effRow = str(int(instModelRow) + 3)
        effCol = chr(ord(instModelColumn) + 2)
        effCell = currentSheet[effCol + effRow]

        if type(effCell).__name__ == 'MergedCell':
            effCol = chr(ord(instModelColumn) + 3)
            effCell = currentSheet[effCol + effRow]

        return effCell

    """
    Take the current sheet and returns the value of the cell to the right of the cell containing Survey No.
    """
    def find_survey_number(currentSheet):
        for column in "ABCDEFGHI":
            for row in range(1, 20):
                modelCell = "{}{}".format(column, row)
                newCol = column

                while (currentSheet[modelCell].value == "Survey No") or (currentSheet[modelCell].value == "Survey Number"):
                    if newCol >= "V":
                        continue

                    newCol = chr(ord(newCol) + 1)
                    cell = currentSheet[newCol + str(row)]

                    if type(cell).__name__ != 'MergedCell':
                        newVal = currentSheet[newCol + str(row)].value
                        return newVal

        return "None"


    def find_efficiency(instSNcell):
        for inst in instrumentsData:
            if inst['sn'] == instSNcell.value:
                return [inst['sn'], inst['betaEfficiency']]

        return [None, None]


    files = getListOfFiles(path)

    """This is used for the exe"""
    # with open(file=resource_path("package.json")) as instruments_file:
    #     instrumentsData = json.load(instruments_file)

    """This is used when running the program"""
    with open('package.json') as instruments_file:
        instrumentsData = json.load(instruments_file)

    QCfileRow = 1
    dateFormat = QCworkbook.add_format({'num_format': 'mm/dd/yyyy'})

    for file in files:
        theFile = openpyxl.load_workbook(file)
        allSheetNames = theFile.sheetnames

        print("All sheet names {} ".format(theFile.sheetnames))

        for x in allSheetNames:
            print("Current sheet name is {}" .format(x))
            currentSheet = theFile[x]
            instModelRow, instModelColumn, instModelCell = find_instrument_model_cell(currentSheet)
            surveyNumber = find_survey_number(currentSheet)

            print("The cell is {}, the row is {} and the column is {} ".format(instModelCell, instModelRow, instModelColumn))

            if instModelCell is None:
                continue
            instSNcell = find_instrument_sn_cell(instModelRow, instModelColumn)
            instEfficiencyCell = find_instrument_efficiency(instModelRow, instModelColumn)

            oldinstEfficiencyCell = instEfficiencyCell.value
            instCalDueDate = find_cal_due_date(instModelRow, instModelColumn)
            serialNumber = find_efficiency(instSNcell)

            if serialNumber[0] is None:
                filesWithNoMatchingSN.append(file)
                sheetsOfFilesWithNoMatchingSN.append(currentSheet)

            # Find the file name
            head, tail = os.path.split(file)

            # Find the Name of the worksheet
            currentSheetString = str(currentSheet)
            currentSheetString = currentSheetString[12:]
            currentSheetString = currentSheetString[:-2]

            betaRow, betaCol = check_for_BettaGamma(3)
            backgroundCol = chr(ord(instModelColumn) + 1)
            index = 0
            if betaRow is None or betaCol is None:
                invalidSheets.append(currentSheet)

            else:
                # There will always be at most 20 counts per survey
                for cell in range(2, 22):
                    cellValue = currentSheet[betaCol + str(betaRow+cell)].value
                    if cellValue is None:
                        continue
                    else:
                        counts.append(cellValue)
                        backgroundCounts.append(currentSheet[backgroundCol + str(betaRow+cell)].value)

                    index += 1

            print("Counts")
            print(counts)
            print(backgroundCounts)

            for x in range(0, len(counts)):

                # Write the current Worksheet
                QCworksheet.write(QCfileRow, 0, tail)
                QCworksheet.write(QCfileRow, 1, currentSheetString)
                QCworksheet.write(QCfileRow, 2, surveyNumber)
                QCworksheet.write(QCfileRow, 3, counts[x])
                QCworksheet.write(QCfileRow, 4, backgroundCounts[x])
                QCworksheet.write(QCfileRow, 5, oldinstEfficiencyCell)

                QCfileRow += 1

        theFile.close()
        theFile.save(file)

    QCworkbook.close()

    print("The files with no s/n are {}, the sheet is {}".format(filesWithNoMatchingSN, sheetsOfFilesWithNoMatchingSN))
    os.startfile(savePath + '\\' + 'Total_Activity.xlsx')

    return
