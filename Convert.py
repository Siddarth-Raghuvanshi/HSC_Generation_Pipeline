#A program to take in the JMP output and convert it into an input for Epmotion machine

#!/usr/bin/python
import xlrd
import sys
import csv
import numpy as np
import string
from math import ceil
from datetime import datetime
from tkinter import messagebox

#Global Variables
Rack_Layout = ["A1","A2","A3","A4","A5","A6","B1","B2","B3","B4","B5","B6","C1","C2","C3","C4","C5","C6","D1","D2","D3","D4","D5","D6"]
Tool = "TS_50"
Source = []

#Opens the Workbook and first sheet using the xlrd
def JMP_Input(Input_File):

    Input_workbook = xlrd.open_workbook(Input_File)
    Input_Sheet = Input_workbook.sheet_by_index(0)

    return Input_Sheet

#Takes the information from the JMP file and places everything into a format readable by Epmotion
def Rearrangment(JMP_Sheet, Layout, Dil_Volume, PlateType, Level_Vol):

    Plate_wells = 60
    num_Factors = JMP_Sheet.row_len(0) - 2
    Levels = set()
    #Used to take into account fractional factorial, i.e. uses all of the levels created by all the factors
    for i in range(num_Factors):
        Levels_for_Factor = set(JMP_Sheet.col_values(i+1,1))
        Levels.update(Levels_for_Factor)
    Levels = list(Levels)
    num_Levels = len(Levels)
    Factors = []

    for i in range(num_Factors):
        Source.append([JMP_Sheet.cell_value(0,i+1),Layout[i]])
        Factors.append(JMP_Sheet.cell_value(0,i+1))

    Dilution_Locations, name = Dilute(Levels, Factors, Dil_Volume)
    if (Dilution_Locations == False):
        while (Dilution_Locations == False):
            messagebox.showerror("Error", "A factor source value is lower than it's level dilution value. Please fill it again")
            Dilution_Locations = Dilute(Levels, Factors, Dil_Volume,True,name)

    num_plates = ceil((JMP_Sheet.nrows - 1)/Plate_wells)
    Plates = []
    for i in range(num_plates):
        Plates.append(Plate(PlateType))
        Run = 1
        for j in range(Plate_wells):
            Row = JMP_Sheet.row_values(Run+j)
            for k in range (1,len(Row)-1):
                Feed_Location =  Dilution_Locations[num_Levels*(k-1)+Levels.index(Row[k])] # a bit much should simplfy.
                Well_Location = Plates[i].Wells[j]
                Plates[i].Commands.append([3,Feed_Location,2,Well_Location,Level_Vol, Tool])
        Run = Run + j

    return Plates

#Create a CSV that dilutes the stock concentrations as per the user input.
def Dilute(Levels, Factors, User_Vol,Screwup = False, name = ""):

    if not Screwup:
        Total_Volume = float(User_Vol)
        Header = [["Factors", "Source"] + Levels]
        name = "Dilution_Concentrations_" + str(datetime.now())+"_SR.csv"
        with open(name, "w") as csvFile:
            writer = csv.writer(csvFile)
            writer.writerows(Header)
            for i in range(len(Factors)):
                writer.writerow([Factors[i]])
        csvFile.close()
        messagebox.showinfo("Dilutions", "A CSV for Dilutions has been created for you to populate with the concentrations of your Factors and Levels, please fill it now")

    input("\nPress Enter to continue once completed...")

    Max_Dilution_Factor = 1/1000 # Max Dilution before it resorts to cereal dilutions
    Dilutions = []
    Commands = []

    with open(name) as Dilution_Concentrations:
        # I know I should use Pandas, but I'm on a time crunch and I don't want to learn it right now *Future programmers should add it for additional functionaity if they choose*
        csv_reader = csv.reader(Dilution_Concentrations, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                line_count += 1
            else:
                for i in range(len(Levels)):
                    if (int(row[1]) < int(row[i+2])):
                        return (False,name)
                    Well_Location = Rack_Layout[(line_count-1)*len(Levels)+i]
                    Volume_to_add = float(row[i+2])/float(row[1])*Total_Volume
                    Top_up_Volume = Total_Volume - Volume_to_add #How much liquid needs to be added to top up to the correct concentration
                    if (Volume_to_add%10 > 1):
                        Commands.append([1,Source[line_count-1][1],1,Well_Location,Volume_to_add%10, "TS_10"])
                        Volume_to_add =  Volume_to_add - Volume_to_add%10
                    if (Top_up_Volume%10 > 1):
                        Commands.append([2,1,1,Well_Location,Top_up_Volume%10, "TS_10"])
                        Top_up_Volume = Top_up_Volume - Top_up_Volume%10
                    if (Volume_to_add%50 > 1):
                        Commands.append([1,Source[line_count-1][1],1,Well_Location,Volume_to_add%50, "TS_50"])
                        Volume_to_add =  Volume_to_add - Volume_to_add%50
                    if (Top_up_Volume%50 > 1):
                        Commands.append([2,1,1,Well_Location,Top_up_Volume%50, "TS_50"])
                        Top_up_Volume = Top_up_Volume - Top_up_Volume%50
                    while (Volume_to_add >= 50):
                        Commands.append([1,Source[line_count-1][1],1,Well_Location,50, "TS_50"])
                        Volume_to_add =  Volume_to_add - 50
                    while(Top_up_Volume >= 50):
                        Commands.append([2,1,1,Well_Location,50, "TS_50"])
                        Top_up_Volume = Top_up_Volume - 50
                    Dilutions.append(Well_Location)
                line_count += 1

    Epmotion_Output(Commands,"DILUTIONS")

    return Dilutions,"No Need"

#Outputs a CSV that is usable by Epmotion
def Epmotion_Output(Info,Purpose):
    #Header
    Header_Data = [["Rack","Source","Rack","Destination","Volume","Tool"]]

    if (Purpose == "PLATE"):
        for i in range(len(Info)):
            #Creates a CSV file and feeds in the new data for Epmotion
            name = "Epmotion_Plates_"+ str(i+1)  + "_" + str(datetime.now())+"_SR.csv"
            with open(name, "w") as csvFile:
                writer = csv.writer(csvFile)
                writer.writerows(Header_Data)
                writer.writerows(Info[i].EdgeData)
                writer.writerows(Info[i].Commands)
            csvFile.close()
    else:
        name = "Dilution_Commands_"+ str(datetime.now())+"_SR.csv"
        with open(name, "w") as csvFile:
            writer = csv.writer(csvFile)
            writer.writerows(Header_Data)
            writer.writerows(Info)
        csvFile.close()

#Outputs a written protocol for original rack placement
def Protcol_Output():
    name = "Protocol_" +  str(datetime.now())+"_SR.txt"
    File =  open(name,"w")

    File.write("Epmotion Protocol\n")
    File.write("Please place Edge Liquid into reservior 1: Well 1")

    File.close()

class Plate():

    def __init__(self, Plate_Type):
        if(Plate_Type == "96 Well"):
            self.num_Edgewells = 1
            self.Rows = 8
            self.Cols = 12
            self.Vol = 200
            self.Tool = 'TS_300'

        elif(Plate_Type == "384 Well"):
            self.num_Edgewells = 2
            self.Rows = 16
            self.Cols = 24
            self.Vol = 50
            self.Tool = 'TS_50'

        self.Commands = []

        #Adds Liquid to account for Edge Effects
        self.EdgeData = []
        for i in range(self.Rows):
            if(i < self.num_Edgewells):
                for j in range(self.Cols):
                    Value = list(string.ascii_uppercase)[i]+str(j+1)
                    self.EdgeData.append(["2","2","2",Value,self.Vol,self.Tool])
            elif(((self.Rows - 1 - i) < self.num_Edgewells)):
                for j in range(self.Cols):
                    Value = list(string.ascii_uppercase)[i]+str(j+1)
                    self.EdgeData.append(["2","2","2",Value,self.Vol,self.Tool])
            else:
                for j in range(self.num_Edgewells):
                    Value = list(string.ascii_uppercase)[i]+str(j+1)
                    self.EdgeData.append(["2","2","2",Value,self.Vol,self.Tool])
                    Value = list(string.ascii_uppercase)[i]+str(self.Cols-j)
                    self.EdgeData.append(["2","2","2",Value,self.Vol,self.Tool])

        #Creates values for useable wells
        self.Wells = []
        for i in range(self.Cols - (self.num_Edgewells * 2)):
            for j in range(self.Rows - (self.num_Edgewells * 2)):
                Well_Value = list(string.ascii_uppercase)[j]+str(i)
                self.Wells.append(Well_Value)
