import openpyxl
import re
import os
import sys
import xlsxwriter
import numpy
import statistics
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
    StatisticSheet = QCworkbook.add_worksheet()

    # create columns with headers
    QCworksheet.write(0, 0, 'File Name')
    QCworksheet.write(0, 1, 'Survey Number')
    QCworksheet.write(0, 2, 'Date')
    QCworksheet.write(0, 3, 'Survey Tech')
    QCworksheet.write(0, 4, 'Count room Tech')
    QCworksheet.write(0, 5, 'Date of Count Room')
    QCworksheet.write(0, 6, 'Survey Type')
    QCworksheet.write(0, 7, 'Level of Posting')
    QCworksheet.write(0, 8, 'Item Surveyed')
    QCworksheet.write(0, 9, 'MDC of total Activity')
    QCworksheet.write(0, 10, 'DPM total activity')
    QCworksheet.write(0, 11, 'MDC of removable')
    QCworksheet.write(0, 12, 'Removable DPM')

    # Creaate statistics sheet Headers
    StatisticSheet.write(0, 0, 'File Name')
    StatisticSheet.write(0, 1, 'Survey Number')
    StatisticSheet.write(0, 2, 'Sheet Name')
    StatisticSheet.write(0, 3, 'Total Activity Min')
    StatisticSheet.write(0, 4, 'Total Activity Max')
    StatisticSheet.write(0, 5, 'Total Activity Average')
    StatisticSheet.write(0, 6, 'Total Activity Standard Deviation')
    StatisticSheet.write(0, 7, 'Removable Min')
    StatisticSheet.write(0, 8, 'Removable Max')
    StatisticSheet.write(0, 9, 'Removable Average')
    StatisticSheet.write(0, 10, 'Removable Standard Deviation')

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
            sampleIdTitleCell = find_cell(currentSheet, "Survey Number")
        sampleIdCell = find_title_data(currentSheet, sampleIdTitleCell)
        print("Survey Number")

        # find survey techs
        surveyTechTitleCell = find_cell(currentSheet, "Survey Tech")
        surveyTechCell = find_title_data(currentSheet, surveyTechTitleCell)
        print("Survey tech")

        # find Count room tech
        countRoomTechTitleCell = find_cell(currentSheet, "Count Room Tech")
        countRoomTechCell = find_title_data(currentSheet, countRoomTechTitleCell)
        print("Survey count tech")

        # find type
        typeTitleCell = find_cell(currentSheet, "Survey Type")
        typeCell = find_title_data(currentSheet, typeTitleCell)
        print("Survey type")

        # find Level of Posting
        postTitleCell = find_cell(currentSheet, "Level Of Posting")
        if postTitleCell[2] is None:
            postTitleCell = find_cell(currentSheet, "Level of Posting")
        postingCell = find_title_data(currentSheet, postTitleCell)
        print("Survey posting")

        # find Item Surveyed
        locationTitleCell = find_cell(currentSheet, "Item Surveyed")
        if locationTitleCell[2] is None:
            locationTitleCell = find_cell(currentSheet, "Survey Number")
        locationCell = find_title_data(currentSheet, locationTitleCell)
        print("Survey Item")

        return [sampleIdCell, surveyTechCell, countRoomTechCell, typeCell, postingCell, locationCell]

    def find_title_data(currentSheet, titleCell):
        row = int(titleCell[0])
        col = chr(ord(titleCell[1]) + 1)

        while type(currentSheet[col + str(row)]).__name__ == 'MergedCell':
            col = chr(ord(col) + 1)

        valueCell = currentSheet[col + str(row)]

        return valueCell

    def remove_isblank(sheet, cellCord):
        print(cellCord)
        temp = sheet[cellCord].value
        temp = re.sub("ISBLANK\([^)]+\)", "FALSE", temp)
        sheet[cellCord].value = temp



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
    SecondSheetRow = 1
    # Create Date format to write to xlsx files
    dateFormat = QCworkbook.add_format({'num_format': 'mm/dd/yyyy'})

    for file in files:
        # For Openpyxl
        theFile = openpyxl.load_workbook(file)
        allSheetNames = theFile.sheetnames
        print(file)
        print("All sheet names {} ".format(theFile.sheetnames))

        for x in allSheetNames:
            print("Current sheet name is {}".format(x))
            currentSheet = theFile[x]

            # If it is a map sheet skip
            currentSheetString = str(currentSheet)
            currentSheetString = currentSheetString[12:]
            checkBlankString = currentSheetString[0:5]
            currentSheetString = currentSheetString[0:3]
            if currentSheetString == "Map" or checkBlankString == "Blank":
                continue

            # Check which format it is
            betaRow, betaCol = check_for_BettaGamma(3)
            index = 0
            needToContinue = False

            # Prevent repeats be added to invalid Sheets list
            if betaRow is None or betaCol is None:
                needToContinue = True
                if len(invalidSheets) > 0:
                    if invalidSheets.count(file) == 0:
                        invalidSheets.append(file)

                else:
                    invalidSheets.append(file)

            # Find MDC Value and DPM for total activity as well as Removable DPM
            else:
                removableBetaRow, removableBetaCol = check_for_BettaGamma(4)
                MDCcol = chr(ord(betaCol) + 2)
                DPMcol = chr(ord(betaCol) + 4)
                removableDPMcol = chr(ord(removableBetaCol) + 2)

                n = 1
                # Go until it is not None
                while currentSheet[betaCol + str(betaRow + n)].value is None:
                    n += 1

                # There will always be at most 20 counts per survey
                for cell in range(n + 1, n + 21):
                    cellValue = currentSheet[betaCol + str(betaRow + cell)].value
                    removableValue = currentSheet[removableBetaCol + str(removableBetaRow + cell)].value

                    # If there is nothing in both removable and total activity
                    if cellValue is None and removableValue is None:
                        continue

                    # If there is removable but no total activity
                    elif cellValue is None and removableValue is not None:
                        removableDPMcell = removableDPMcol + str(betaRow + cell)
                        MDCCounts.append(None)
                        dpmCounts.append(None)
                        removableCounts.append(removableDPMcell)

                    # If there is total activity but no removable
                    elif cellValue is not None and removableValue is None:
                        MDCcell = MDCcol + str(betaRow + cell)
                        DPMcell = DPMcol + str(betaRow + cell)
                        MDCCounts.append(MDCcell)
                        dpmCounts.append(DPMcell)
                        removableCounts.append(None)

                    # If there is both removable and total activity
                    else:
                        MDCcell = MDCcol + str(betaRow + cell)
                        DPMcell = DPMcol + str(betaRow + cell)
                        removableDPMcell = removableDPMcol + str(betaRow + cell)
                        MDCCounts.append(MDCcell)
                        dpmCounts.append(DPMcell)
                        removableCounts.append(removableDPMcell)

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
            for x in range(0, len(removableCounts)):
                print("The file " + str(file))
                print("The Sheet " + currentSheetString)
                print("MDCCounts " + str(MDCCounts[x]))
                ttest = 0

                if dpmCounts[x] is not None:
                    netCPM = dpmCounts[x]
                    netCol = chr(ord(netCPM[0]) - 1)
                    netRow = netCPM[1:]
                    netCPM = netCol + netRow

                    remove_isblank(currentSheet, str(MDCCounts[x]))
                    remove_isblank(currentSheet, netCPM)
                    remove_isblank(currentSheet, str(dpmCounts[x]))
                    theFile.save(file)
                    excel = ExcelCompiler(filename=file)
                    tempCell = currentSheetString + "!" + str(MDCCounts[x])
                    print(tempCell)
                    MDCCounts[x] = excel.evaluate(tempCell)
                    tempCell = currentSheetString + "!" + str(dpmCounts[x])
                    print(tempCell)
                    dpmCounts[x] = excel.evaluate(tempCell)
                    ttest += 1


                if removableCounts[x] is not None:
                    if ttest > 0:
                        ttest = ttest
                    netCPM = removableCounts[x]
                    netCol = chr(ord(netCPM[0]) - 1)
                    netRow = netCPM[1:]
                    netCPM = netCol + netRow

                    remove_isblank(currentSheet, netCPM)
                    remove_isblank(currentSheet, str(removableCounts[x]))
                    theFile.save(file)
                    excel = ExcelCompiler(filename=file)
                    tempCell = currentSheetString + "!" + str(removableCounts[x])
                    removableCounts[x] = excel.evaluate(tempCell)

            # Write the results to the QC file
            # Write the current Worksheet
            head, tail = os.path.split(file)
            length = len(removableCounts)
            for x in range(0, len(removableCounts)):
                QCworksheet.write(stat, 0, tail)                                       # File Name
                QCworksheet.write(QCfileRow, 1, titleVals[0].value)                         # Survey Number
                QCworksheet.write(QCfileRow, 2, dateCell.value, dateFormat)                 # Date
                QCworksheet.write(QCfileRow, 3, titleVals[1].value)                         # Survey Tech
                QCworksheet.write(QCfileRow, 4, titleVals[2].value)                         # Count room Tech
                QCworksheet.write(QCfileRow, 5, secondDateCell.value, dateFormat)           # Date of Count Room
                QCworksheet.write(QCfileRow, 6, titleVals[3].value)                         # Survey Type
                QCworksheet.write(QCfileRow, 7, titleVals[4].value)                         # Level of Posting
                QCworksheet.write(QCfileRow, 8, titleVals[5].value)                         # Item Surveyed
                QCworksheet.write(QCfileRow, 9, MDCCounts[x])                               # MDC of total Activity
                QCworksheet.write(QCfileRow, 10, dpmCounts[x])                              # DPM total activity
                QCworksheet.write(QCfileRow, 11, "MDCremovableCounts[x]")                   # MDC of removable
                QCworksheet.write(QCfileRow, 12, removableCounts[x])                        # Removable DPM

                QCfileRow += 1

            # Find the statistics
            totalAverage = sum(dpmCounts) / len(dpmCounts)
            totalMax = max(dpmCounts)
            totalMin = min(dpmCounts)
            totalStdDev = statistics.pstdev(dpmCounts)

            removableAvg = sum(removableCounts) / len(removableCounts)
            removableMax = max(removableCounts)
            removableMin = min(removableCounts)
            removableStdDev = statistics.pstdev(removableCounts)

            StatisticSheet.write(SecondSheetRow, 0, tail)                               # File Name
            StatisticSheet.write(SecondSheetRow, 1, titleVals[0].value)                 # Survey Number
            StatisticSheet.write(SecondSheetRow, 2, currentSheetString)                 # Current Sheet
            StatisticSheet.write(SecondSheetRow, 3, totalMin)
            StatisticSheet.write(SecondSheetRow, 4, totalMax)
            StatisticSheet.write(SecondSheetRow, 5, totalAverage)
            StatisticSheet.write(SecondSheetRow, 6, totalStdDev)
            StatisticSheet.write(SecondSheetRow, 7, removableMin)
            StatisticSheet.write(SecondSheetRow, 8, removableMax)
            StatisticSheet.write(SecondSheetRow, 9, removableAvg)
            StatisticSheet.write(SecondSheetRow, 10, removableStdDev)

            SecondSheetRow += 1

        theFile.close()
        theFile.save(file)
        MDCCounts.clear()
        dpmCounts.clear()
        removableCounts.clear()

    """redo but for the other list"""
    print("IN THE INVALID BETA ")
    del dpmCounts
    del MDCCounts
    del removableCounts
    dpmCounts = list()
    removableCounts = list()

    for file in invalidSheets:
        theFile = openpyxl.load_workbook(file)
        allSheetNames = theFile.sheetnames
        print(file)
        print("All sheet names {} ".format(theFile.sheetnames))

        for x in allSheetNames:
            print("Current sheet name is {}".format(x))
            currentSheet = theFile[x]

            # If it is a map sheet skip
            currentSheetString = str(currentSheet)
            currentSheetString = currentSheetString[12:]
            checkBlankString = currentSheetString[0:5]
            currentSheetString = currentSheetString[0:3]
            if currentSheetString == "Map" or checkBlankString == "Blank":
                continue

            betaRow, betaCol = check_for_BettaGamma2(3)
            if betaRow is not None or betaCol is not None:
                continue

            else:
                betaRow, betaCol = check_for_BettaGamma(1)
                removableBetaRow, removableBetaCol = check_for_BettaGamma(2)
                DPMcol = chr(ord(betaCol) + 1)
                removableDPMcol = chr(ord(removableBetaCol) + 1)

                n = 1
                # Go until it is not None
                while currentSheet[betaCol + str(betaRow + n)].value != "gross counts":
                    print(currentSheet[betaCol + str(betaRow + n)].value)
                    n += 1

                # There will always be at most 20 counts per survey
                for cell in range(n + 1, n + 21):
                    cellValue = currentSheet[betaCol + str(betaRow + cell)].value
                    removableValue = currentSheet[removableBetaCol + str(removableBetaRow + cell)].value

                    # If there is nothing in both removable and total activity
                    if cellValue is None and removableValue is None:
                        continue

                    # If there is removable but no total activity
                    elif cellValue is None and removableValue is not None:
                        removableDPMcell = removableDPMcol + str(betaRow + cell)
                        dpmCounts.append(None)
                        removableCounts.append(removableDPMcell)

                    # If there is total activity but no removable
                    elif cellValue is not None and removableValue is None:
                        DPMcell = DPMcol + str(betaRow + cell)
                        dpmCounts.append(DPMcell)
                        removableCounts.append(None)

                    # If there is both removable and total activity
                    else:
                        DPMcell = DPMcol + str(betaRow + cell)
                        removableDPMcell = removableDPMcol + str(betaRow + cell)
                        dpmCounts.append(DPMcell)
                        removableCounts.append(removableDPMcell)

                    index += 1

                test = 1

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

            currentSheetString = str(currentSheet)
            currentSheetString = currentSheetString[12:]
            currentSheetString = currentSheetString[:-2]

            if currentSheet == "2360-190602 (2)":
                ("on second sprintheet")
            # Pycel
            # Find the values of mdc and dpm
            print(dpmCounts)
            for x in range(0, len(removableCounts)):

                if dpmCounts[x] is not None:
                    remove_isblank(currentSheet, str(dpmCounts[x]))
                    theFile.save(file)
                    excel = ExcelCompiler(filename=file)
                    tempCell = currentSheetString + "!" + str(dpmCounts[x])
                    dpmCounts[x] = excel.evaluate(tempCell)

                if removableCounts[x] is not None:
                    remove_isblank(currentSheet, str(removableCounts[x]))
                    theFile.save(file)
                    excel = ExcelCompiler(filename=file)
                    tempCell = currentSheetString + "!" + str(removableCounts[x])
                    removableCounts[x] = excel.evaluate(tempCell)

            print("Adding data to file.")
            head, tail = os.path.split(file)
            for x in range(0, len(removableCounts)):
                QCworksheet.write(QCfileRow, 0, tail)  # File Name
                QCworksheet.write(QCfileRow, 1, titleVals[0].value)  # Survey Number
                QCworksheet.write(QCfileRow, 2, dateCell.value, dateFormat)  # Date
                QCworksheet.write(QCfileRow, 3, titleVals[1].value)  # Survey Tech
                QCworksheet.write(QCfileRow, 4, titleVals[2].value)  # Count room Tech
                QCworksheet.write(QCfileRow, 5, secondDateCell.value, dateFormat)  # Date of Count Room
                QCworksheet.write(QCfileRow, 6, titleVals[3].value)  # Survey Type
                QCworksheet.write(QCfileRow, 7, titleVals[4].value)  # Level of Posting
                QCworksheet.write(QCfileRow, 8, titleVals[5].value)  # Item Surveyed
                QCworksheet.write(QCfileRow, 9, "MDCCounts[x]")  # MDC of total Activity
                QCworksheet.write(QCfileRow, 10, dpmCounts[x])  # DPM total activity
                QCworksheet.write(QCfileRow, 11, "MDCremovableCounts[x]")  # MDC of removable
                QCworksheet.write(QCfileRow, 12, removableCounts[x])  # Removable DPM

                QCfileRow += 1
                
            # Find the statistics
            totalAverage = sum(dpmCounts) / len(dpmCounts)
            totalMax = max(dpmCounts)
            totalMin = min(dpmCounts)
            totalStdDev = statistics.pstdev(dpmCounts)

            removableAvg = sum(removableCounts) / len(removableCounts)
            removableMax = max(removableCounts)
            removableMin = min(removableCounts)
            removableStdDev = statistics.pstdev(removableCounts)

            StatisticSheet.write(SecondSheetRow, 0, tail)  # File Name
            StatisticSheet.write(SecondSheetRow, 1, titleVals[0].value)  # Survey Number
            StatisticSheet.write(SecondSheetRow, 2, currentSheetString)  # Current Sheet
            StatisticSheet.write(SecondSheetRow, 3, totalMin)
            StatisticSheet.write(SecondSheetRow, 4, totalMax)
            StatisticSheet.write(SecondSheetRow, 5, totalAverage)
            StatisticSheet.write(SecondSheetRow, 6, totalStdDev)
            StatisticSheet.write(SecondSheetRow, 7, removableMin)
            StatisticSheet.write(SecondSheetRow, 8, removableMax)
            StatisticSheet.write(SecondSheetRow, 9, removableAvg)
            StatisticSheet.write(SecondSheetRow, 10, removableStdDev)

            SecondSheetRow += 1



            dpmCounts.clear()
            removableCounts.clear()

        theFile.close()
        theFile.save(file)
        dpmCounts.clear()
        removableCounts.clear()

    QCworkbook.close()
    os.startfile(savePath + '\\' + 'SurveyTrending.xlsx')
