import openpyxl
import os
import sys
import xlsxwriter
from pycel import ExcelCompiler


def main(path, savePath):
    dpmCounts = list()
    MDCCounts = list()
    removableCounts = list()
    invalidSheets = list()
    badfile = list()

    # Create the output folder
    if os.path.isfile(savePath):
        os.mkdir(savePath)

    print(path)

    # create excel QC file
    QCworkbook = xlsxwriter.Workbook(savePath + '\\' + 'SurveyTrending.xlsx')
    QCworksheet = QCworkbook.add_worksheet()

    # create columns with headers
    QCworksheet.write(0, 0, 'File Name')
    QCworksheet.write(0, 1, 'Sample ID')
    QCworksheet.write(0, 2, 'Date')
    QCworksheet.write(0, 3, 'Location')
    QCworksheet.write(0, 4, 'Sample Type')
    QCworksheet.write(0, 5, 'Alpha Activity')
    QCworksheet.write(0, 6, 'Alpha MDC')
    QCworksheet.write(0, 7, 'Beta Activity')
    QCworksheet.write(0, 8, 'Beta MDC')

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

        print(allFiles)

        return allFiles

    def find_cell(currentSheet, parameterToFind):
        for row in range(1, 40):
            for column in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":  # Here you can add or reduce the columns

                cell = "{}{}".format(column, row)
                cellValue = currentSheet[cell].value

                if currentSheet[cell].value == parameterToFind:
                    print("the row is {0} and the column {1}".format(row, column))

                    print(currentSheet[cell].value)
                    print(cell)

                    return [row, column, currentSheet[cell]]

        return [0, 0, None]

    # This is used when we need to find a value and the searched term occurs more than once
    def find_date_cell(currentSheet, parameterToFind, occurenceNeeded):
        found = 0
        for row in range(1, 30):
            for column in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":

                modelVal = currentSheet[column + str(row)].value
                if modelVal != parameterToFind:
                    continue

                found += 1
                if found == occurenceNeeded:
                    return [row, column]

        return [None, None]

    def find_title_vals(currentSheet):
        # find Survey number
        sampleIdTitleCell = find_cell(currentSheet, "Survey No")
        if sampleIdTitleCell[2] is None:
            print("Success test")
            sampleIdTitleCell = find_cell(currentSheet, "Survey Number")
        sampleIdCell = find_title_data(currentSheet, sampleIdTitleCell)
        print("Survey Number")

        # find survey techs
        surveyTechTitleCell = find_cell(currentSheet, "Survey Tech")
        surveyTechCell = find_title_data(currentSheet, surveyTechTitleCell)
        print("Survey tech")

        # find Count room tech
        countRoomTechTitleCell = find_cell(currentSheet, "Date Counted")
        countRoomTechCell = find_title_data(currentSheet, countRoomTechTitleCell)
        print("Survey count tech")

        # find type
        typeTitleCell = find_cell(currentSheet, "Survey Type")
        typeCell = find_title_data(currentSheet, typeTitleCell)
        print("Survey type")

        # find Level of Posting
        postTitleCell = find_cell(currentSheet, "Level Of Posting")
        postingCell = find_title_data(currentSheet, postTitleCell)
        print("Survey posting")

        # find Item Surveyed
        locationTitleCell = find_cell(currentSheet, "Survey No")
        if locationTitleCell[2] is None:
            locationTitleCell = find_cell(currentSheet, "Survey Number")
        locationCell = find_title_data(currentSheet, locationTitleCell)
        print("Survey Item")

        return [sampleIdCell, surveyTechCell, countRoomTechCell, typeCell, postingCell, locationCell]

    def find_title_data(currentSheet, titleCell):
        row = int(titleCell[0])
        print(row)
        print(titleCell[1])
        col = chr(ord(titleCell[1]) + 1)
        print(col)

        while type(currentSheet[col + str(row)]).__name__ == 'MergedCell':
            col = chr(ord(col) + 1)

        valueCell = currentSheet[col + str(row)]

        return valueCell

    def check_for_BettaGamma(num):
        found = 0
        for row in range(1, 30):
            for column in "GHIJKLMNOPQRSTUVWXYZ":

                modelVal = currentSheet[column + str(row)].value
                if modelVal != "Beta-Gamma":
                    continue

                found += 1
                if found == num:
                    return [row, column]

        return [None, None]

    def check_for_BettaGamma2(num):
        found = 0
        for row in range(1, 30):
            for column in "GHIJKLMNOPQRSTUVWXYZ":

                modelVal = currentSheet[column + str(row)].value
                if modelVal != "Beta-Gamma":
                    continue

                found += 1
                if found == num:
                    return [row, column]

        return [None, None]

    files = getListOfFiles(path)

    QCfileRow = 1
    # Create Date format to write to xlsx files
    dateFormat = QCworkbook.add_format({'num_format': 'mm/dd/yyyy'})

    for file in files:
        # For PyCel
        excel = ExcelCompiler(filename=file)
        # For Openpyxl
        theFile = openpyxl.load_workbook(file)
        allSheetNames = theFile.sheetnames

        print("All sheet names {} ".format(theFile.sheetnames))



        for x in allSheetNames:
            print("Current sheet name is {}".format(x))
            currentSheet = theFile[x]

            # If it is a map sheet skip
            currentSheetString = str(currentSheet)
            currentSheetString = currentSheetString[12:]
            currentSheetString = currentSheetString[0:3]
            if currentSheetString == "Map":
                continue

            # Check which format it is
            betaRow, betaCol = check_for_BettaGamma(3)
            index = 0
            needToContinue = False

            # Prevent repeats be added to invalid Sheets list
            if betaRow is None or betaCol is None:
                if len(invalidSheets) > 0:
                    if invalidSheets.count(file) > 0:
                        needToContinue = False
                else:
                    invalidSheets.append(file)
                    needToContinue = True

            # Find MDC Value and DPM for total activity as well as Removable DPM
            else:
                removableBetaRow, removableBetaCol = check_for_BettaGamma(4)
                MDCcol = chr(ord(betaCol) + 2)
                DPMcol = chr(ord(betaCol) + 4)
                removableDPMcol = chr(ord(removableBetaRow) + 2)

                n = 1
                # Go until it is not None
                while currentSheet[betaCol + str(betaRow + n)].value is None:
                    n += 1

                # There will always be at most 20 counts per survey
                for cell in range(n + 1, n + 21):
                    cellValue = currentSheet[betaCol + str(betaRow + cell)].value
                    removableValue = currentSheet[removableDPMcol + str(removableBetaRow + cell)].value

                    # If there is nothing in both removable and total activity
                    if cellValue is None and removableValue is None:
                        continue

                    # If there is removable but no total activity
                    elif cellValue is None and removableValue is not None:
                        MDCCounts.append(None)
                        dpmCounts.append(None)
                        removableCounts.append(removableValue)

                    # If there is total activity but no removable
                    elif cellValue is not None and removableValue is None:
                        MDCValue = currentSheet[MDCcol + str(betaRow + cell)].value
                        DPMValue = currentSheet[DPMcol + str(betaRow + cell)].value
                        MDCCounts.append(MDCValue)
                        dpmCounts.append(DPMValue)
                        removableCounts.append(None)

                    # If there is both removable and total activity
                    else:
                        MDCValue = currentSheet[MDCcol + str(betaRow + cell)].value
                        DPMValue = currentSheet[DPMcol + str(betaRow + cell)].value
                        MDCCounts.append(MDCValue)
                        dpmCounts.append(DPMValue)
                        removableCounts.append(removableValue)

                    print(file)
                    print("DPM value: " + str(DPMValue))
                    print("betaCol: " + betaCol)
                    print("backgroundCol: " + MDCcol)
                    print("betaRow+Cell" + str(betaRow + cell))
                    print("MDCValue: " + str(MDCValue))
                    print("removableValue: " + str(DPMValue))

                    index += 1

            if needToContinue is True:
                continue

            # ***********Find Title Data**********
            titleVals = find_title_vals(currentSheet)
            print("After titlevals")

            # find date
            dateTitleCell = find_cell(currentSheet, "Date")
            if dateTitleCell[2] is None or dateTitleCell[1] == 0:
                dateTitleCell = find_cell(currentSheet, "Date Counted")
            dateCell = find_title_data(currentSheet, dateTitleCell)
            print("After date")

            # Find Count Room Date Counted
            dateTitleCell = find_date_cell(currentSheet, "Date Counted", 2)
            if dateTitleCell[1] is None or dateTitleCell[1] == 0:
                dateTitleCell = find_date_cell(currentSheet, "Date Counted", 1)
            secondDateCell = find_title_data(currentSheet, dateTitleCell)
            print("After second date")





            # Find the Name of the worksheet
            currentSheetString = str(currentSheet)
            currentSheetString = currentSheetString[12:]
            currentSheetString = currentSheetString[:-2]

            # PyCel
            # Find the values of mdc and dpm






            # Write the results to the QC file
            # Write the current Worksheet
            head, tail = os.path.split(file)

            for x in range(0, len(removableCounts)):
                QCworksheet.write(QCfileRow, 0, tail)                           # File Name
                QCworksheet.write(QCfileRow, 1, titleVals[0].value)             # Survey Number
                QCworksheet.write(QCfileRow, 2, dateCell.value, dateFormat)     # Date
                QCworksheet.write(QCfileRow, 3, titleVals[1].value)             # Survey Tech
                QCworksheet.write(QCfileRow, 4, titleVals[2].value)             # Count room Tech
                QCworksheet.write(QCfileRow, 5, secondDateCell.value)           # Date of Count Room Tech
                QCworksheet.write(QCfileRow, 6, titleVals[3].value)             # Survey Type
                QCworksheet.write(QCfileRow, 7, titleVals[4].value)             # Level of Posting
                QCworksheet.write(QCfileRow, 8, titleVals[5].value)             # Item Surveyed

                QCfileRow += 1

        theFile.close()
        theFile.save(file)
        MDCCounts.clear()
        dpmCounts.clear()
        removableCounts.clear()

    QCworkbook.close()
    os.startfile(savePath + '\\' + 'SurveyTrending.xlsx')
