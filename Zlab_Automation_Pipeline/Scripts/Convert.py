#A program to take in the JMP output and convert it into an input for Epmotion machine

#!/usr/bin/python
import csv
import numpy as np
import string
from math import ceil
from tkinter import messagebox
import pandas as pd
from pathlib import Path

Epitube_Vol = 1600 # I know that global variables are considered to be terrible, but frankly this is a much more elegant solution than anything else I've had.

#Takes the information from the JMP file and places everything into a format readable by Epmotion
def Rearrangment(JMP_Sheet, Layout, PlateType, Well_Vol,Edge_Well, Dead_Vol):

    Rack = len(Layout)
    Source = []
    num_Y_Vars = 0
    for Col in range(JMP_Sheet.row_len(0)):
        if JMP_Sheet.cell_value(1,Col) == "":
            num_Y_Vars += 1
    num_Factors = JMP_Sheet.row_len(1) - 1 - num_Y_Vars
    Factor_Vol = Well_Vol/num_Factors
    Levels = set()
    num_Runs = JMP_Sheet.nrows - 1
    Dil_Volumes = []
    Cell_Volume = []
    #Used to take into account fractional factorial, i.e. uses all of the levels created by all the factors
    for i in range(num_Factors):
        Levels_for_Factor = set(JMP_Sheet.col_values(i+1,1))
        Levels.update(Levels_for_Factor)
    Levels = sorted(Levels)
    #Find the volume for each Level which will be tested
    for i in range(num_Factors):
        for j,Coded_Level in enumerate(Levels):
            Dil_Volumes.append(ceil(JMP_Sheet.col_values(i+1,1).count(Coded_Level)*Factor_Vol*1.1+Dead_Vol))
            Cell_Volume.append(JMP_Sheet.col_values(i+1,1).count(Coded_Level)*10)
    num_Levels = len(Levels)
    Factors = []
    Total_Tests = num_Levels * num_Factors

    for i in range(num_Factors):
        Source.append([JMP_Sheet.cell_value(0,i+1),Layout[i]])
        Factors.append(JMP_Sheet.cell_value(0,i+1))

    #Change the 24 welll Rack to a 96 well plate if there are numerous factors and not that many runs
    if num_Levels*num_Factors > 48-num_Factors:
        if max(Dil_Volumes) < 250:
            messagebox.showinfo("EpMotion Layout", "The dilutions will be made in a 96 Well Plate as there are too many factors")
            Layout = []
            Rack = 96
            for i in range(12):
                for j in range(8):
                    Layout.append(list(string.ascii_uppercase)[j]+str(i+1))
        else:
            messagebox.showinfo("Error", "Your Experiment has a large amount of Factors and Runs, which is beyond the space avalible in the EpMotion. I recommended blocking your experiments to avoid nuisance factors")
            quit()

    name, Screwup = Get_Concentrations(Levels, Factors)
    if (Screwup):
        while (Screwup):
            messagebox.showerror("Error", "A factor source value is lower than it's level dilution value. Please fill it again")
            name, Screwup = Get_Concentrations(Levels, Factors,True,name)

    Dilution_Locations, Dilution_Commands, Needed_Vol, Media_Vol_Needed = Dilute(Layout, Source, Levels, Factors, Dil_Volumes, Dead_Vol, name, Cell_Volume)

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
                Plates[i].Commands.append([3,Feed_Location,2,Well_Location,10, "TS_50"])#Factor_Vol, "TS_50"])
        Trials += j

    return Plates,len(Dilution_Locations), Dilution_Commands, Source, Needed_Vol, Rack, Media_Vol_Needed

#Create a CSV which contains the user specified dilutions, and asks the user to input these concentrations
def Get_Concentrations(Levels, Factors,Screwup = False, name = ""):

    if not Screwup:

        Header = [["Factors", "Source"] + Levels ]
        name = "Dilution_Concentrations_SR.csv"
        with open(Path.cwd() / name, "w") as csvFile:
            writer = csv.writer(csvFile, lineterminator = "\n")
            writer.writerows(Header)
            for Factor in Factors:
                writer.writerow([Factor])
        csvFile.close()
        # Don't break up the text, it doesn't seem to display properly.
        messagebox.showinfo("Dilutions", "A File called " + name + " has been created for you to populate with the concentrations of your Factors and Levels, please fill it now")

    input("\nPress Enter to continue once completed...")

    Concentrations = pd.read_csv(Path.cwd() / name, header = 0)
    if not (Concentrations.loc[:, "Source":].idxmax(1) == "Source").all():
        Screwup = True

    return name, Screwup

#Creates a CSV for the epmotion which allows it to dilute the manual concentration to the JMP dilutions.
# It creates these manual concentrations in such a way to ensure that factor is not wasted.
def Dilute(Layout,Source, Levels, Factors, User_Vol, Dead_Vol, name, Cell_Volume):

    Min_Dilution = 0.5 # Minimum Volume of dilution before it resorts to cereal dilutions
    Dilutions = []
    Commands = []
    cereal_Dilutions = 0
    Manual_Concentrations = []
    Needed_Vol = []

    #Find the dilution which should be done manually so as to not waste any factor
    Dilution_Conc = pd.read_csv(Path.cwd() / name)
    Dilution_Conc.set_index("Factors", inplace = True)
    for i,Factor in enumerate(Factors):
        Desired_Volume = (max(User_Vol[i*len(Levels):(i+1)*len(Levels)]) - Dead_Vol) * 0.9 #Divide by 0.9 giving a 10 % safety buffer
        if Desired_Volume > Epitube_Vol: #This way if the total volume is too large, then limit it to the epitube, but 90% of the time, it will generally be less the epitube, so it shouldn't be limited by that
            Desired_Volume = Epitube_Vol - Dead_Vol * 0.9
        Source_HighConc_Ratio = Dilution_Conc.loc[Factor]["Source"] / (Dilution_Conc.loc[Factor][len(Levels)]*len(Factors))
        Vol_times_Conc = sum(User_Vol[i*len(Levels):(i+1)*len(Levels)]*Dilution_Conc.loc[Factor,:][1:])*len(Factors)
        if Source_HighConc_Ratio < 1:
            messagebox.showinfo("Error", Factor + "'s Source concentration is too low for the concentration needed for its HIGH value to fill the plate")
            quit()
        else:
            Manual_Concentrations.append(ceil(Vol_times_Conc/Desired_Volume))
            Needed_Vol.append([(Desired_Volume*Manual_Concentrations[i])/Dilution_Conc.loc[Factor]["Source"], Desired_Volume])

        if Manual_Concentrations[i] > Dilution_Conc.loc[Factor]["Source"]:
            messagebox.showinfo("Error", "The program's recommended manual dilution is higher than the source concentration. This is generally because you have too many runs and a low source concentration.")
            quit()

    Dilution_Conc["Manually Diluted Concentration"] = Manual_Concentrations
    Dilution_Conc.to_csv(Path.cwd() / name)

    Dilution_Liquid_Needed = 0

    with open(Path.cwd() / name) as Dilution_Concentrations:
        csv_reader = csv.reader(Dilution_Concentrations, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                line_count += 1
            else:
                for i in range(len(Levels)):
                    if ((line_count-1)*len(Levels)+i+cereal_Dilutions < 24):
                        Well_Location = Layout[(line_count-1)*len(Levels)+i+cereal_Dilutions]
                        Destination = 1
                    else:
                        Well_Location = Layout[((line_count-1)*len(Levels)+i+cereal_Dilutions) - 24 + len(Factors)]
                        Destination = 2
                    Volume_to_add = (float(row[i+2])*len(Factors))/Manual_Concentrations[line_count-1]*User_Vol[(line_count-1)*len(Levels)+i]
                    cereal_Run = False
                    if (Volume_to_add < Min_Dilution):
                        Volume_to_add = Volume_to_add * Epitube_Vol/User_Vol[(line_count-1)*len(Levels)+i]
                        User_Vol[(line_count-1)*len(Levels)+i] = Epitube_Vol
                    if (Volume_to_add < Min_Dilution):
                        Volume_to_add = 0.5
                        cereal_Run = True
                        cereal_Dilutions +=1
                    #How much liquid needs to be added to top up to the correct concentration
                    Top_up_Volume = User_Vol[(line_count-1)*len(Levels)+i] - Volume_to_add - Cell_Volume[(line_count-1)*len(Levels)+i]
                    Dilution_Liquid_Needed += Top_up_Volume
                    Commands.extend(Fill_Up(Source[line_count-1][1], Top_up_Volume, Volume_to_add, Well_Location, Destination))
                    Dilutions.append(Well_Location)

                    while(cereal_Run):
                        cereal_Run = False
                        Cereal_Location = Well_Location
                        if ((line_count-1)*len(Levels)+i+cereal_Dilutions < 24):
                            Well_Location = Layout[(line_count-1)*len(Levels)+i+cereal_Dilutions]
                            Destination = 1
                        else:
                            Well_Location = Layout[((line_count-1)*len(Levels)+i+cereal_Dilutions) - 24 + len(Factors)]
                            Destination = 2
                        Volume_to_add = (float(row[i+2])*len(Factors))/(Manual_Concentrations[line_count-1]*Volume_to_add/Epitube_Vol)*User_Vol[(line_count-1)*len(Levels)+i]
                        if (Volume_to_add < Min_Dilution):
                            Volume_to_add = Volume_to_add * Epitube_Vol/User_Vol[(line_count-1)*len(Levels)+i]
                            User_Vol[(line_count-1)*len(Levels)+i] = Epitube_Vol
                        if (Volume_to_add < Min_Dilution):
                            Volume_to_add = 0.5
                            cereal_Run = True
                            cereal_Dilutions +=1
                        Top_up_Volume = User_Vol[(line_count-1)*len(Levels)+i] - Volume_to_add
                        Commands.extend(Fill_Up(Source[line_count-1][1], Top_up_Volume, Volume_to_add, Well_Location, Destination))
                        Dilutions.append(Well_Location)
                line_count += 1

    return Dilutions, Commands, Needed_Vol, ceil(Dilution_Liquid_Needed/1000)*1100

def Fill_Up(Source, Top_up_Volume, Volume_to_add, Well_Location, Destination):
    Commands = []
    if (Top_up_Volume%10 > 0.5):
        Commands.append([2,1,Destination,Well_Location,Top_up_Volume%10, "TS_10"])
        Top_up_Volume = Top_up_Volume - Top_up_Volume%10
    while (Top_up_Volume % 50 > 0.5):
        Commands.append([2,1,Destination,Well_Location,Top_up_Volume%50, "TS_50"])
        Top_up_Volume = Top_up_Volume - Top_up_Volume%50
    if (Volume_to_add%10 > 0.5):
        Commands.append([1,Source,Destination,Well_Location,Volume_to_add%10, "TS_10"])
        Volume_to_add =  Volume_to_add - Volume_to_add%10
    while (Volume_to_add % 50 > 0.5):
        Commands.append([1,Source,Destination,Well_Location,Volume_to_add%50, "TS_50"])
        Volume_to_add =  Volume_to_add - Volume_to_add%50
    if (Volume_to_add%1000 > 0):
        Commands.append([1,Source,Destination,Well_Location,Volume_to_add%1000, "TS_1000"])
        Volume_to_add =  Volume_to_add - Volume_to_add%1000
    if (Top_up_Volume%1000 > 0):
        Commands.append([2,1,Destination,Well_Location,Top_up_Volume%1000, "TS_1000"])
        Top_up_Volume = Top_up_Volume - Top_up_Volume%1000
    while (Volume_to_add >= 1000):
        Commands.append([1,Source,Destination,Well_Location,1000, "TS_1000"])
        Volume_to_add =  Volume_to_add - 1000
    while(Top_up_Volume >= 1000):
        Commands.append([2,1,Destination,Well_Location,1000, "TS_1000"])
        Top_up_Volume = Top_up_Volume - 1000

    return Commands

class Plate():

    def __init__(self, Plate_Type,Edge):
        if(Plate_Type == "96 Well"):
            self.Rows = 8
            self.Cols = 12
            self.Vol = 250
            self.Tool = 'TS_1000'

        elif(Plate_Type == "384 Well"):
            self.Rows = 16
            self.Cols = 24
            self.Vol = 100
            self.Tool = 'TS_1000'

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
