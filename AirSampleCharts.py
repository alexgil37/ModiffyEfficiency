import shutil
import openpyxl
import json
import os
import sys
import xlsxwriter
from pycel import ExcelCompiler

def main(path, savePath):

    # # Create the output folder
    # if os.path.isfile(savePath):
    #     os.mkdir(savePath)

    def getFile(dirName):
        listOfFile = os.listdir(dirName)
        allFiles = list()

        for file in listOfFile:
            filePath = os.path.join(dirName, file)
            theFile = openpyxl.load_workbook(filePath)

            break

        return theFile

    def find_location(dataFileWS):

        # sheet current row
        sbCurrentRow = 3
        # cwCurrentRow = 3
        # bCurrentRow = 3
        # firstSeCurrentRow = 3
        # secondSeCurrentRow = 3
        # secondHwCurrentRow = 3
        # secondOutCurrentRow = 3
        # thirdSeCurrentRow = 3
        # fourthSeCurrentRow = 3
        # fifthSeCurrentRow = 3
        # sixthSeCurrentRow = 3
        # seventhSeCurrentRow = 3
        # seventh726CurrentRow = 3
        # eigthSeCurrentRow = 3
        # roofEastCurrentRow = 3
        # roofWestCurrentRow = 3

        for column in "D":
            for row in range(1, 4000):

                cell1 = "{}{}".format(column, row)

                if "Hallway" in dataFileWS[cell1].value:
                    sbCurrentRow = sbCurrentRow + 1

                    sbSheet.cell(sbCurrentRow, 1).value = dataFileWS["B", str(row)]
                    # sbSheet.cell(int(sbCurrentRow), 2).value = dataFile["C", str(row)]
                    # sbSheet.cell(int(sbCurrentRow), 3).value = dataFile["F", str(row)]
                    # sbSheet.cell(int(sbCurrentRow), 4).value = dataFile["G", str(row)]
                    # sbSheet.cell(int(sbCurrentRow), 5).value = dataFile["H", str(row)]
                    # sbSheet.cell(int(sbCurrentRow), 6).value = dataFile["I", str(row)]

    dataFileWB = getFile(path)
    dataFileWS = dataFileWB["Sheet1"]

    file = shutil.copy("Air Sample Results Trending Template.xlsx", savePath + "\\" + "AirSamplesCharts.xlsx")
    wb = openpyxl.load_workbook(file)

    #sheets
    sbSheet = wb["SB Svc Ele"]
    # cwSheet = wb["CW Svc Ele"]
    # bSheet = wb["B Svc Ele"]
    # firstSeSheet = wb["1st Svc Ele"]
    # secondSeSheet = wb["2nd Svc Ele"]
    # secondHwSheet = wb["2nd Hall"]
    # secondOutSheet = wb["2nd Out"]
    # thirdSeSheet = wb["3rd Svc Ele"]
    # fourthSeSheet = wb["4th Svc Ele"]
    # fifthSeSheet = wb["5th Svc Ele"]
    # sixthSeSheet = wb["6th Svc Ele"]
    # seventhSeSheet = wb["7th Svc Ele"]
    # seventh726Sheet = wb["7th Rm 726"]
    # eigthSeSheet = wb["8th Svc Ele"]
    # roofEastSheet = wb["9th East"]
    # roofWestSheet = wb["9th West"]

    find_location(dataFileWS)

    wb.close()
    wb.save(file)
