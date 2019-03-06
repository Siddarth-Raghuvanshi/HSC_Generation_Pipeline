#A program to take in the JMP output and transform it into an input for Epimotion machine

#!/usr/bin/python
import xlrd
import sys
import csv
import numpy as np
import string
from math import ceil
from datetime import datetime

#Global Variables
Rack_Layout = ["A1","A2","A3","A4","A5","A6","B1","B2","B3","B4","B5","B6","C1","C2","C3","C4","C5","C6","D1","D2","D3","D4","D5","D6","D7"]
Volume = "10"
Tool = "TS_50"
Source = []

#Opens the Workbook and first sheet using the xlrd
def JMP_Input(Input_File):

    Input_workbook = xlrd.open_workbook(Input_File)
    Input_Sheet = Input_workbook.sheet_by_index(0)

    return Input_Sheet

#Takes the information from the JMP file and places everything into a format readable by Epimotion
def Rearrangment(JMP_Sheet, Layout):

    Plate_wells = 60

    Levels = list(set(JMP_Sheet.col_values(1,1)))
    num_Levels = len(Levels)
    num_Factors = JMP_Sheet.row_len(0) - 2

    for i in range(num_Factors):
        if (i == 0):
            Source.append(["Edge Liquid",Layout[i]])
            Source.append([JMP_Sheet.cell_value(0,i+1),Layout[i+1]])
            j = i + 1
        else:
            Source.append([JMP_Sheet.cell_value(0,i+1),Layout[j]])
        for k in range(num_Levels):
            j = j + 1
            Source.append([JMP_Sheet.cell_value(0,i+1) + "_" + str(Levels[k]),Layout[j]])
        j = j + 1

    num_plates = ceil((JMP_Sheet.nrows - 1)/Plate_wells)

    Plates = []

    for i in range(num_plates):
        Plates.append(Plate())
        Run = 1
        for j in range(Plate_wells):
            Row = JMP_Sheet.row_values(Run+j)
            for k in range (1,len(Row)-1):
                Feed_Location =  Layout[(num_Levels*(k-1)+k+Levels.index(Row[k])+1)] # a bit much should simplfy.
                Well_Location = Plates[i].Wells[j]
                Plates[i].Commands.append([1,Feed_Location,i+1,Well_Location,Volume, Tool])
        Run = Run + j

    return Plates

#Outputs a CSV that is usable by Epimotion
def Epimotion_Output(Plates_Info):

    #Header
    Header_Data = [["Rack","Source","Rack","Destination","Volume","Tool"]]

    #Creates a CSV file and feeds in the new data for Epimotion
    name = "Epimotion_" + str(datetime.now())+"_SR.csv"
    with open(name, "w") as csvFile:
        writer = csv.writer(csvFile)
        writer.writerows(Header_Data)

        for i in range(len(Plates_Info)):
            writer.writerows(Plates_Info[i].EdgeData)
            writer.writerows(Plates_Info[i].Commands)
    csvFile.close()

class Plate:

    #Adds Liquid to account for Edge Effects
    EdgeData = []
    num_Edgewells = 36
    for i in range(num_Edgewells):
        if (i<12) or (i>24):
            Letter = "A"
            if (i>24):
                Letter = "H"
            Value = Letter+str(i%12+1)
            EdgeData.append(["1","A1","1",Value,Volume,Tool])
        elif (i<18):
            Value = list(string.ascii_uppercase)[i%6+1]+"1"
            EdgeData.append(["1","A1","1",Value,Volume,Tool])
        else:
            Value = list(string.ascii_uppercase)[i%6+1]+"12"
            EdgeData.append(["1","A1","1",Value,Volume,Tool])

    #Creates values for useable wells
    Wells = []
    num_Cols = 11
    num_rows = 6
    for i in range(num_Cols):
        for j in range(num_rows):
            Well_Value = list(string.ascii_uppercase)[j%6+1]+str(((i+2)%12))
            Wells.append(Well_Value)

    def __init__(self):
      self.Commands = []

if __name__ == '__main__':

    JMP_Sheet = JMP_Input(sys.argv[1])
    Output_Plates = Rearrangment(JMP_Sheet, Rack_Layout)
    Epimotion_Output(Output_Plates)
    #Protcol_Output(OutputInfo)
