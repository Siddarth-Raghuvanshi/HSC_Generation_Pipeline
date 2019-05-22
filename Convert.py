#A program to take in the JMP output and convert it into an input for Epmotion machine

#!/usr/bin/python
import csv
import numpy as np
import string
from math import ceil
from tkinter import messagebox

#Takes the information from the JMP file and places everything into a format readable by Epmotion
def Rearrangment(JMP_Sheet, Layout, Dil_Volume, PlateType, Level_Vol,Edge_Well):

    Source = []
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

    Dilution_Locations, name, Dilution_Commands = Dilute(Layout, Source, Levels, Factors, Dil_Volume)
    if (Dilution_Locations == False):
        while (Dilution_Locations == False):
            messagebox.showerror("Error", "A factor source value is lower than it's level dilution value. Please fill it again")
            Dilution_Locations = Dilute(Layout,Source,Levels, Factors, Dil_Volume,True,name)

    Plate_wells = len(Plate(PlateType,Edge_Well).Wells)
    num_plates = ceil((JMP_Sheet.nrows - 1)/Plate_wells)
    Plates = []
    Trials = 1
    for i in range(num_plates):
        Plates.append(Plate(PlateType,Edge_Well))
        for j in range(Plate_wells):
            if(Trials+j > num_Runs):
                break
            Row = JMP_Sheet.row_values(Trials+j)
            for k in range (1,len(Row)-1):
                Feed_Location =  Dilution_Locations[num_Levels*(k-1)+Levels.index(Row[k])] # a bit much should simplfy.
                Well_Location = Plates[i].Wells[j]
                Plates[i].Commands.append([3,Feed_Location,2,Well_Location,Level_Vol, "TS_50"])
        Trials += j

    return Plates,len(Dilution_Locations), Dilution_Commands, Source

#Create a CSV that dilutes the stock concentrations as per the user input.
def Dilute(Layout,Source, Levels, Factors, User_Vol,Screwup = False, name = ""):

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
                    Well_Location = Layout[(line_count-1)*len(Levels)+i+cereal_Dilutions]
                    Volume_to_add = float(row[i+2])/float(row[1])*Total_Volume
                    if (Volume_to_add < Min_Dilution):
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
                        Well_Location = Layout[(line_count-1)*len(Levels)+i+cereal_Dilutions]
                        Volume_to_add = float(row[i+2])/(float(row[1])/10)*Total_Volume
                        if (Volume_to_add < Min_Dilution):
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

    return Dilutions,"No Need", Commands

class Plate():

    def __init__(self, Plate_Type,Edge):
        if(Plate_Type == "96 Well"):
            self.Rows = 8
            self.Cols = 12
            self.Vol = 250
            self.Tool = 'TS_300'

        elif(Plate_Type == "384 Well"):
            self.Rows = 16
            self.Cols = 24
            self.Vol = 100
            self.Tool = 'TS_300'

        self.num_Edgewells = Edge
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
