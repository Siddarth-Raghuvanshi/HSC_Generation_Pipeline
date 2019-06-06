#A program to take in the JMP output and convert it into an input for Epmotion machine

#!/usr/bin/python
import csv
import numpy as np
import string
from math import ceil
from tkinter import messagebox
import pandas as pd

#Takes the information from the JMP file and places everything into a format readable by Epmotion
def Rearrangment(JMP_Sheet, Layout, PlateType, Well_Vol,Edge_Well, Dead_Vol):

    Rack = len(Layout)
    Source = []
    num_Y_Vars = 0
    for Col in range(JMP_Sheet.row_len(0)):
        if JMP_Sheet.cell_value(1,Col) == "":
            num_Y_Vars += 1
    num_Factors = JMP_Sheet.row_len(0) - 1 - num_Y_Vars
    Factor_Vol = Well_Vol/num_Factors
    Levels = set()
    num_Runs = JMP_Sheet.nrows - 1
    Dil_Volumes = []
    #Used to take into account fractional factorial, i.e. uses all of the levels created by all the factors
    for i in range(num_Factors):
        Levels_for_Factor = set(JMP_Sheet.col_values(i+1,1))
        Levels.update(Levels_for_Factor)
    #Find the volume for each Level which will be tested
    for i in range(num_Factors):
        for j,Coded_Level in enumerate(Levels):
            Dil_Volumes.append(JMP_Sheet.col_values(i+1,1).count(Coded_Level)*Factor_Vol*1.1+Dead_Vol)
    Levels = list(Levels)
    num_Levels = len(Levels)
    Factors = []
    Total_Tests = num_Levels * num_Factors

    for i in range(num_Factors):
        Source.append([JMP_Sheet.cell_value(0,i+1),Layout[i]])
        Factors.append(JMP_Sheet.cell_value(0,i+1))


    #Change the 24 welll Rack to a 96 well plate if there are numerous factors and not that many runs
    if num_Levels*len(Factors) > 24:
        if max(Dil_Volumes) < 250:
            messagebox.showinfo("EpMotion Layout", "The dilutions will be made in a 96 Well Plate as there are too many factors")
            Layout = []
            Rack = 96
            for i in range(12):
                for j in range(8):
                    Layout.append(list(string.ascii_uppercase)[j]+str(i+1))
        else:
            messagebox.showinfo("Error", "Your experiment has too many Factors/Levels/Runs to run on the EpMotion. Please restructure it in JMP")
            quit()

    Dilution_Locations, name, Dilution_Commands, Needed_Vol = Dilute(Layout, Source, Levels, Factors, Dil_Volumes, Dead_Vol)
    if (Dilution_Locations == False):
        while (Dilution_Locations == False):
            messagebox.showerror("Error", "A factor source value is lower than it's level dilution value. Please fill it again")
            Dilution_Locations, name, Dilution_Commands, Needed_Vol = Dilute(Layout,Source,Levels, Factors, Dil_Volumes, Dead_Vol, True,name, )

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
            for k in range (1,len(Row)- 1 - num_Y_Vars):
                Feed_Location =  Dilution_Locations[num_Levels*(k-1)+Levels.index(Row[k])] # a bit much should simplfy.
                Well_Location = Plates[i].Wells[j]
                Plates[i].Commands.append([3,Feed_Location,2,Well_Location,Factor_Vol, "TS_50"])
        Trials += j

    return Plates,len(Dilution_Locations), Dilution_Commands, Source, Needed_Vol, Rack

#Create a CSV that dilutes the stock concentrations as per the user input.
def Dilute(Layout,Source, Levels, Factors, User_Vol, Dead_Vol,Screwup = False, name = ""):

    if not Screwup:

        Header = [["Factors", "Source"] + Levels ]
        name = "Dilution_Concentrations_SR.csv"
        with open("../" + name, "w") as csvFile:
            writer = csv.writer(csvFile)
            writer.writerows(Header)
            for Factor in Factors:
                writer.writerow([Factor])
        csvFile.close()
        # Don't break up the text, it doesn't seem to display properly.
        messagebox.showinfo("Dilutions", "A File called " + name + " has been created for you to populate with the concentrations of your Factors and Levels, please fill it now")

    input("\nPress Enter to continue once completed...")

    Min_Dilution = 0.5 # Minimum Volume of dilution before it resorts to cereal dilutions
    Dilutions = []
    Commands = []
    cereal_Dilutions = 0
    Manual_Concentrations = []
    Needed_Vol = []

    #Find the dilution which should be done manually so as to not waste any factor
    Desired_Volume = (1500 - Dead_Vol)*0.9# 1500 - 20 (Dead Volume for EpMotion) * 0.9 (10% Safety Barrier)
    Dilution_Conc = pd.read_csv("../" + name)
    Dilution_Conc.set_index("Factors", inplace = True)
    for i,Factor in enumerate(Factors):
        Vol_times_Conc = sum(User_Vol[i*len(Levels):(i+1)*len(Levels)]*Dilution_Conc.loc[Factor,:][1:])
        Manual_Concentrations.append(ceil(Vol_times_Conc/Desired_Volume))
        Needed_Vol.append((ceil(Vol_times_Conc/Desired_Volume)*1500)/Dilution_Conc.loc[Factor,:][0])

    Dilution_Conc["Manually Diluted Concentration"] = Manual_Concentrations
    Dilution_Conc.to_csv(name)

    with open(name) as Dilution_Concentrations:
        # I know I should use Pandas, but I'm on a time crunch and I don't want to learn it right now *Future programmers should add it for additional functionaity if they choose*
        csv_reader = csv.reader(Dilution_Concentrations, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                line_count += 1
            else:
                for i in range(len(Levels)):
                    if (float(row[1]) < float(row[i+2])):
                        return (False,name)
                    Well_Location = Layout[(line_count-1)*len(Levels)+i+cereal_Dilutions]
                    Volume_to_add = (float(row[i+2])*len(Factors))/Manual_Concentrations[line_count-1]*User_Vol[(line_count-1)*len(Levels)+i]
                    if (Volume_to_add < Min_Dilution):
                        Volume_to_add = 10
                        cereal_Run = True
                        cereal_Dilutions +=1
                    #How much liquid needs to be added to top up to the correct concentration
                    Top_up_Volume = User_Vol[(line_count-1)*len(Levels)+i] - Volume_to_add
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
                    while(cereal_Dilutions):
                        cereal_Dilutions = False
                        Cereal_Location = Well_Location
                        Well_Location = Layout[(line_count-1)*len(Levels)+i+cereal_Dilutions]
                        Volume_to_add = (float(row[i+2])*len(Factors))/(Manual_Concentrations[line_count-1]/10)*User_Vol[(line_count-1)*len(Levels)+i]
                        if (Volume_to_add < Min_Dilution):
                            Volume_to_add = 1
                            cereal_Dilutions = True
                            cereal_Dilutions +=1
                        #How much liquid needs to be added to top up to the correct concentration
                        Top_up_Volume = User_Vol[(line_count-1)*len(Levels)+i] - Volume_to_add
                        if (Volume_to_add%10 > 1):
                            Commands.append([3,Cereal_Location,1,Well_Location,Volume_to_add%10, "TS_10"])
                            Volume_to_add =  Volume_to_add - Volume_to_add%10
                        if (Top_up_Volume%10 > 1):
                            Commands.append([2,1,1,Well_Location,Top_up_Volume%10, "TS_10"])
                            Top_up_Volume = Top_up_Volume - Top_up_Volume%10
                        if (Volume_to_add%50 > 1):
                            Commands.append([3,Cereal_Location,1,Well_Location,Volume_to_add%50, "TS_50"])
                            Volume_to_add =  Volume_to_add - Volume_to_add%50
                        if (Top_up_Volume%50 > 1):
                            Commands.append([2,1,1,Well_Location,Top_up_Volume%50, "TS_50"])
                            Top_up_Volume = Top_up_Volume - Top_up_Volume%50
                        while (Volume_to_add >= 50):
                            Commands.append([3,Cereal_Location,1,Well_Location,50, "TS_50"])
                            Volume_to_add =  Volume_to_add - 50
                        while(Top_up_Volume >= 50):
                            Commands.append([2,1,1,Well_Location,50, "TS_50"])
                            Top_up_Volume = Top_up_Volume - 50
                        Dilutions.append(Well_Location)
                line_count += 1

    return Dilutions,"No Need", Commands, Needed_Vol

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
