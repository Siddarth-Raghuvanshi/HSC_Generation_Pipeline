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
import os

#Global Variables
Rack_Layout = ["A1","A2","A3","A4","A5","A6","B1","B2","B3","B4","B5","B6","C1","C2","C3","C4","C5","C6","D1","D2","D3","D4","D5","D6"]
Tool = "TS_50"
Source = []
Folder_Name = os.getcwd()+"/Experiment_Files_" + str(datetime.now()) + "_SR"
os.mkdir(Folder_Name)

#Takes the information from the JMP file and places everything into a format readable by Epmotion
def Rearrangment(JMP_Sheet, Layout, Dil_Volume, PlateType, Level_Vol):

    num_Factors = JMP_Sheet.row_len(0) - 2
    Levels = set()
    num_Runs = JMP_Sheet.nrows - 1
    #Used to take into account fractional factorial, i.e. uses all of the levels created by all the factors
    for i in range(num_Factors):
        Levels_for_Factor = set(JMP_Sheet.col_values(i+1,1))
        Levels.update(Levels_for_Factor)
    Levels = list(Levels)
    num_Levels = len(Levels)
    Factors = []
    Total_Tests = num_Levels * num_Factors

    for i in range(num_Factors):
        Source.append([JMP_Sheet.cell_value(0,i+1),Layout[i]])
        Factors.append(JMP_Sheet.cell_value(0,i+1))

    Dilution_Locations, name = Dilute(Levels, Factors, Dil_Volume)
    if (Dilution_Locations == False):
        while (Dilution_Locations == False):
            messagebox.showerror("Error", "A factor source value is lower than it's level dilution value. Please fill it again")
            Dilution_Locations = Dilute(Levels, Factors, Dil_Volume,True,name)

    Plate_wells = len(Plate(PlateType).Wells)
    num_plates = ceil((JMP_Sheet.nrows - 1)/Plate_wells)
    Plates = []
    Trials = 1
    for i in range(num_plates):
        Plates.append(Plate(PlateType))
        for j in range(Plate_wells):
            if(Trials+j > num_Runs):
                break
            Row = JMP_Sheet.row_values(Trials+j)
            for k in range (1,len(Row)-1):
                Feed_Location =  Dilution_Locations[num_Levels*(k-1)+Levels.index(Row[k])] # a bit much should simplfy.
                Well_Location = Plates[i].Wells[j]
                Plates[i].Commands.append([3,Feed_Location,2,Well_Location,Level_Vol, Tool])
        Trials += j

    os.rename("Dilution_Concentrations_SR.csv", Folder_Name+"/Dilution_Concentrations_SR.csv")

    return Plates,Total_Tests

#Create a CSV that dilutes the stock concentrations as per the user input.
def Dilute(Levels, Factors, User_Vol,Screwup = False, name = ""):

    Total_Volume = float(User_Vol)
    if not Screwup:

        Header = [["Factors", "Source"] + Levels]
        name = "Dilution_Concentrations_SR.csv"
        with open(name, "w") as csvFile:
            writer = csv.writer(csvFile)
            writer.writerows(Header)
            for i in range(len(Factors)):
                writer.writerow([Factors[i]])
        csvFile.close()
        messagebox.showinfo("Dilutions", "A CSV for Dilutions has been created for you to populate with the concentrations of your Factors and Levels, please fill it now")

    input("\nPress Enter to continue once completed...")

    Min_Dilution = 5 # Minimum Voluem of dilution before it resorts to cereal dilutions
    Dilutions = []
    Commands = []
    cereal_Dilutions = 0

    with open("/Users/siddarthraghuvanshi/Documents/Code/HSC_Generation_Pipeline/test_concentrations.csv") as Dilution_Concentrations:
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
                    Well_Location = Rack_Layout[(line_count-1)*len(Levels)+i+cereal_Dilutions]
                    Volume_to_add = float(row[i+2])/float(row[1])*Total_Volume
                    if (Volume_to_add < Min_Dilution)
                        Volume_to_add = 10
                        cereal_Run = True
                        cereal_Dilutions +=1
                    Top_up_Volume = Total_Volume - Volume_to_add #How much liquid needs to be added to top up to the correct concentration
                    if (Volume_to_add%50 > 1):
                        Commands.append([1,Source[line_count-1][1],1,Well_Location,Volume_to_add%50, "TS_50"])
                        Volume_to_add =  Volume_to_add - Volume_to_add%50
                    if (Top_up_Volume%50 > 1):
                        Commands.append([2,1,1,Well_Location,Top_up_Volume%50, "TS_50"])
                        Top_up_Volume = Top_up_Volume - Top_up_Volume%50
                    if (Volume_to_add%300 > 1):
                        Commands.append([1,Source[line_count-1][1],1,Well_Location,Volume_to_add%300, "TS_300"])
                        Volume_to_add =  Volume_to_add - Volume_to_add%300
                    if (Top_up_Volume%300 > 1):
                        Commands.append([2,1,1,Well_Location,Top_up_Volume%300, "TS_300"])
                        Top_up_Volume = Top_up_Volume - Top_up_Volume%300
                    while (Volume_to_add >= 300):
                        Commands.append([1,Source[line_count-1][1],1,Well_Location,300, "TS_300"])
                        Volume_to_add =  Volume_to_add - 300
                    while(Top_up_Volume >= 300):
                        Commands.append([2,1,1,Well_Location,300, "TS_300"])
                        Top_up_Volume = Top_up_Volume - 300
                    Dilutions.append(Well_Location)
                    while(cereal_Dilutions):
                        cereal_Dilutions = False
                        Cereal_Location = Well_Location
                        Well_Location = Rack_Layout[(line_count-1)*len(Levels)+i+cereal_Dilutions]
                        Volume_to_add = float(row[i+2])/(float(row[1])/10)*Total_Volume
                        if (Volume_to_add < Min_Dilution)
                            Volume_to_add = 10
                            cereal_Dilutions = True
                            cereal_Dilutions +=1
                        Top_up_Volume = Total_Volume - Volume_to_add #How much liquid needs to be added to top up to the correct concentration
                        if (Volume_to_add%50 > 1):
                            Commands.append([3,Cereal_Location,1,Well_Location,Volume_to_add%50, "TS_50"])
                            Volume_to_add =  Volume_to_add - Volume_to_add%50
                        if (Top_up_Volume%50 > 1):
                            Commands.append([2,1,1,Well_Location,Top_up_Volume%50, "TS_50"])
                            Top_up_Volume = Top_up_Volume - Top_up_Volume%50
                        if (Volume_to_add%300 > 1):
                            Commands.append([3,Cereal_Location,1,Well_Location,Volume_to_add%300, "TS_300"])
                            Volume_to_add =  Volume_to_add - Volume_to_add%300
                        if (Top_up_Volume%300 > 1):
                            Commands.append([2,1,1,Well_Location,Top_up_Volume%300, "TS_300"])
                            Top_up_Volume = Top_up_Volume - Top_up_Volume%300
                        while (Volume_to_add >= 300):
                            Commands.append([3,Cereal_Location,1,Well_Location,300, "TS_300"])
                            Volume_to_add =  Volume_to_add - 300
                        while(Top_up_Volume >= 300):
                            Commands.append([2,1,1,Well_Location,300, "TS_300"])
                            Top_up_Volume = Top_up_Volume - 300
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
            name = Folder_Name + "/Epmotion_Plate_"+ str(i+1) + "_SR.csv"
            with open(name, "w") as csvFile:
                writer = csv.writer(csvFile)
                writer.writerows(Header_Data)
                writer.writerows(Info[i].EdgeData)
                writer.writerows(Info[i].Commands)
            csvFile.close()
    else:
        name = Folder_Name + "/Dilution_Commands_SR.csv"
        with open(name, "w") as csvFile:
            writer = csv.writer(csvFile)
            writer.writerows(Header_Data)
            writer.writerows(Info)
        csvFile.close()

#Outputs a written protocol for original rack placement
def Protcol_Output(Tests):
    name = Folder_Name + "/Protocol_SR.txt"
    File =  open(name,"w")

    File.write("Epmotion Protocol\n\n")
    File.write("RACK PLACEMENT \n\n")
    File.write("1. Place a reservoir into the EpMotion\n")
    File.write("2. Place 2 24-well racks into the EpMotion\n")
    File.write("3. Place a box of 50 and 300 ul tips into the EpMotion\n")
    File.write("4. Ensure that there is a space for a plate in the EpMotion\n")
    File.write("5. Place a TS_50 and a TS_300 into the EpMotion\n\n")

    File.write("LIQUID LAYOUT \n\n")
    File.write("1. Place a boat containing dilution liquid into the 1st slot in the reservoir\n")
    File.write("2. Place a boat containing edge liquid into the 2nd slot in the reservoir\n")
    for i in range(len(Source)):
        File.write("%d. Place the Stock %s into the %s well in the first rack\n" % ( i+3, Source[i][0], Source[i][1]))
    File.write("%d. Place %d sterile Epitubes into the second rack from %s to %s\n\n" %(len(Source) + 3, Tests, Rack_Layout[0], Rack_Layout[Tests-1]))

    File.write("EPBLUE PROTOCOL \n\n")
    File.write("1. Create a new application.\n")
    File.write("2. Insert the Sample transfer command.\n")
    File.write("3. Under the Parameter Option, place the Stock 24-well as Source 1\n")
    File.write("4. Under the Parameter Option, place the reservoir as Source 2\n")
    File.write("5. Under the Parameter Option, place the Dilution 24-well as Source 3\n")
    File.write("6. Under the Parameter Option, place the Dilution 24-well as Destination 1\n")
    File.write("7. Under the Parameter Option, place the Plate as Destination 2\n")
    File.write("8. Ensure all other setting in the transfer command are satisfactory\n")
    File.write("9. Click the Sample transfer command.\n")
    File.write("10. Click on the CSV Command in top bar.\n")
    File.write("11. Select the Dilution_Command CSV and hit open.\n")
    File.write("12. Select the Pipette option and hit OK\n")
    File.write("13. Once finished repeat step 11 with the CSV required for each plate.\n")

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
                Well_Value = list(string.ascii_uppercase)[j]+str(i+1)
                self.Wells.append(Well_Value)
