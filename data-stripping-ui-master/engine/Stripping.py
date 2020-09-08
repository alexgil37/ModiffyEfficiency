import csv
import json
import math
from tkinter import *
import xlsxwriter
import sys


# Establishing file names
#fileName = 'Geo_176_spec_merge.csv'
#workBookName = fileName + '.xlsx'

fileName = sys.argv[1]
workBookName = sys.argv[2] + '.xlsx'

# using pandas for graphing
#writer = pd.ExcelWriter(workBookName, engine='xlsxwriter')

# Creating Output Excel WorkBook
outWB = xlsxwriter.Workbook('engine/Results/' + workBookName)
outSheetValues = outWB.add_worksheet("Values")
outSheetGraph = outWB.add_worksheet("Graph")

print("Output created")

outSheetValues.write(0, 0, "Measurement_Number")
outSheetValues.write(0, 1, "Date/Time")
outSheetValues.write(0, 2, "Position_X")
outSheetValues.write(0, 3, "Position_Y")
outSheetValues.write(0, 4, "Gross_Gamma")
outSheetValues.write(0, 5, "Cs-137")
outSheetValues.write(0, 6, "Co-60_1173_keV")
outSheetValues.write(0, 7, "Co-60_1332_keV")
outSheetValues.write(0, 8, "K-40")
outSheetValues.write(0, 9, "Uranium")
outSheetValues.write(0, 10, "Thorium")

with open(fileName, 'r') as csv_file:
    csv_reader = csv.reader(csv_file)

    with open('engine/ROIs') as config_file:
        ROIdata = json.load(config_file)

    # empty arrays for the spectra
    totalSpectra = [0] * 1024
    currentSpectrum = [0] * 1024
    background = [0] * 1024

    # system properties
    # channelToEnergy = 2.922
    # resolution = 0.07
    # FWHMcoefficient = 1.5

    # arbitrary stripping coefficients
    alpha = 0.5
    beta = 0.5
    gamma = 0.5

    def backgroundSubstraction():
        for loc in range(1024):
            currentSpectrum[loc] = int(currentSpectrum[loc]) - background[loc]
            if currentSpectrum[loc] < 0:
                currentSpectrum[loc] = 0

    def createROI(ROIstart, ROIend):
        ROI = []
        for loc in range(1024):
            if ROIstart <= loc <= ROIend:
                ROI.append(currentSpectrum[loc])

        return ROI

    def addCountsROI(ROI):
        totalCounts = 0
        for k in range(len(ROI)):
            totalCounts = totalCounts + int(ROI[k])

        return totalCounts

    def startChannelROI(isotopeName):
        for ist in ROIdata:
            if ist['isotope'] == isotopeName:
                ROIstart = ist["start"]

        return ROIstart

    def endChannelROI(isotopeName):
        for ist in ROIdata:
            if ist['isotope'] == isotopeName:
                ROIend = ist["end"]

        return ROIend


    # reads the CSV file and iterates through each line, generates array with spectrum data
    row = 0
    for i, line in enumerate(csv_reader):
        column = 0
        channel = 0

        if i > 0:
            if not line[0] or not line[1] or not line[2]:
                continue
            else:
                for j in line:
                    if column == 0:
                        dateTime = j
                        outSheetValues.write(row, 1, dateTime)
                    elif column == 1:
                        xPosition = j
                        outSheetValues.write(row, 2, float(xPosition))
                    elif column == 2:
                        yPosition = j
                        outSheetValues.write(row, 3, float(yPosition))
                    elif 11 < column < 1036:
                        currentSpectrum[channel] = j
                        totalSpectra[channel] = totalSpectra[channel] + int(j)
                        channel += 1

                    column += 1

        if i > 0:
            if not line[0] or not line[1] or not line[2]:
                continue
            else:
                outSheetValues.write(row, 0, row)

                # start and end of ROIs for each isotope
                CsStart = startChannelROI("Cs-137")
                CsEnd = endChannelROI("Cs-137")

                CoLowStart = startChannelROI("Co-60-1173")
                CoLowEnd = endChannelROI("Co-60-1173")

                CoHighStart = startChannelROI("Co-60-1332")
                CoHighEnd = endChannelROI("Co-60-1332")

                Kstart = startChannelROI("K")
                Kend = endChannelROI("K")

                Ustart = startChannelROI("U")
                Uend = endChannelROI("U")

                ThStart = startChannelROI("Th")
                ThEnd = endChannelROI("Th")

                # Create ROI and get total counts
                ROIofCs = createROI(CsStart, CsEnd)
                CsTotalCounts = addCountsROI(ROIofCs)

                ROIofCoLow = createROI(CoLowStart, CoLowEnd)
                CoLowTotalCounts = addCountsROI(ROIofCoLow)

                ROIofCoHigh = createROI(CoHighStart, CoHighEnd)
                CoHighTotalCounts = addCountsROI(ROIofCoHigh)

                ROIofK = createROI(Kstart, Kend)
                KtotalCounts = addCountsROI(ROIofK)

                ROIofU = createROI(Ustart, Uend)
                UtotalCounts = addCountsROI(ROIofU)

                ROIofTH = createROI(ThStart, ThEnd)
                ThTotalCounts = addCountsROI(ROIofTH)

                grossGamma = addCountsROI(currentSpectrum)

                #stripp values and add to Excel

                realCsCounts = int(CsTotalCounts - (beta * ThTotalCounts) - (gamma * UtotalCounts))
                outSheetValues.write(row, 5, realCsCounts)

                outSheetValues.write(row, 6, CoLowTotalCounts)

                outSheetValues.write(row, 7, CoHighTotalCounts)

                realKCounts = int(KtotalCounts - (beta * ThTotalCounts) - (gamma * UtotalCounts))
                outSheetValues.write(row, 8, realKCounts)

                realUCounts = int(UtotalCounts - (beta * ThTotalCounts))
                outSheetValues.write(row, 9, realUCounts)

                outSheetValues.write(row, 10, ThTotalCounts)

                outSheetValues.write(row, 4, grossGamma)

        row += 1

    # df = pd.DataFrame({'counts': totalSpectra})
    # writer = pd.ExcelWriter('Results/' + workBookName, engine='xlsxwriter')
    # df.to_excel(writer, sheet_name=outSheetGraph)
    # chart = outWB.add_chart({'type': 'column'})



            # Stripp to get real counts
            # realCsCounts = int(CsTotalCounts - (beta * ThTotalCounts) - (gamma * UtotalCounts))
            # realKCounts = int(KtotalCounts - (beta * ThTotalCounts) - (gamma * UtotalCounts))

            # print(KtotalCounts)
            # print(realKCounts)

outWB.close()

    #print(int(Kstart))
    #print(int(Kend))
