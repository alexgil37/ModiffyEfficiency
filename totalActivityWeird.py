import openpyxl
import totalActivity
import os
import sys
import xlsxwriter


filesWithNoMatchingSN = list()
sheetsOfFilesWithNoMatchingSN = list()
counts = list()
backgroundCounts = list()
invalidSheets = list()

def check_for_BettaGamma(num, currentSheet):
    found = 0
    for row in range(1, 30):
        for column in "GHIJKLMNOPQRSTUVWXYZ":

            modelVal = currentSheet[column + str(row)].value
            if modelVal == "Background Counts":
                Col = chr(ord(column) + 5)
                modelVal = currentSheet[Col + str(row)].value
                if modelVal is not None:
                    return [0, 0, modelVal]

            if modelVal == "Bldg Material Bkg":
                return [row, column, modelVal]

    return [None, None]


def main(weirdFiles, QCfileRow):
    for file in weirdFiles:
        theFile = openpyxl.load_workbook(file)
        allSheetNames = theFile.sheetnames

        for x in allSheetNames:
            print("Current sheet name is {}" .format(x))
            currentSheet = theFile[x]
            instModelRow, instModelColumn, instModelCell = totalActivity.find_instrument_model_cell(currentSheet)
            surveyNumber = totalActivity.find_instrument_model_cell(currentSheet)

            if instModelCell is None:
                continue
            instSNcell = totalActivity.find_instrument_model_cell(instModelRow, instModelColumn)
            instEfficiencyCell = totalActivity.find_instrument_model_cell(instModelRow, instModelColumn)

            oldinstEfficiencyCell = instEfficiencyCell.value
            serialNumber = totalActivity.find_instrument_model_cell(instSNcell)

            if serialNumber[0] is None:
                totalActivity.filesWithNoMatchingSN.append(file)
                totalActivity.sheetsOfFilesWithNoMatchingSN.append(currentSheet)

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
                invalidSheets.append(file)
                continue

            else:
                # There will always be at most 20 counts per survey
                print(betaRow)
                n = 1
                # Go until you get to Gross Counts
                while currentSheet[betaCol + str(betaRow+n)].value is None:
                    n += 1

                for cell in range(n+1, n+21):
                    cellValue = currentSheet[betaCol + str(betaRow+cell)].value
                    if cellValue is None:
                        continue
                    else:
                        counts.append(cellValue)
                        backgroundCounts.append(currentSheet[backgroundCol + str(betaRow+cell)].value)

                    index += 1

            for x in range(0, len(counts)):

                # Write the current Worksheet
                QCworksheet.write(QCfileRow, 0, tail)
                QCworksheet.write(QCfileRow, 1, currentSheetString)
                QCworksheet.write(QCfileRow, 2, surveyNumber)
                QCworksheet.write(QCfileRow, 3, counts[x])
                QCworksheet.write(QCfileRow, 4, backgroundCounts[x])
                QCworksheet.write(QCfileRow, 5, oldinstEfficiencyCell)

                QCfileRow += 1

            counts.clear()
            backgroundCounts.clear()

        theFile.close()
        theFile.save(file)

    QCworkbook.close()

    print(invalidSheets)