#A program to take in the JMP output and convert it into an input for Epmotion machine

#!/usr/bin/python
import csv
import numpy as np
import string
from math import ceil, floor
from tkinter import messagebox
import pandas as pd
from pathlib import Path

Epitube_Vol = 1600 # I know that global variables are considered to be terrible, but frankly this is a much more elegant solution than anything else I've had.

#Takes the information from the JMP file and places everything into a format readable by Epmotion
def Rearrangment(Commands, Layout, PlateType, Well_Vol,Edge_Well, Dead_Vol, Added_Cell_Vol):

    Rack = len(Layout)
    Source = []
    Y_Vars = Commands.columns[Commands.isna().all()].tolist()
    X_Vars = Commands.drop(Y_Vars, axis =1)
    Factor_Vol = Well_Vol/len(X_Vars.columns)  #Volume of Factors the EpiMotion should add
    Cell_Vol = Added_Cell_Vol/len(X_Vars.columns)
    Levels = set()
    num_Runs = len(X_Vars)
    Dil_Volumes = []
    Cell_Volumes = []
    #Used to take into account fractional factorial, i.e. uses all of the levels created by all the factors
    [Levels.update(X_Vars[Factor].unique().tolist()) for Factor in X_Vars.columns]
    Levels = sorted(Levels)

    #Find the volume for each Level which will be tested
    Summary = pd.concat([X_Vars[Factor].value_counts() for Factor in X_Vars.columns], axis =1).T

    for i in range(num_Factors):
        for j,Coded_Level in enumerate(Levels):
            Dil_Volumes.append(ceil(JMP_Sheet.col_values(i+1,1).count(Coded_Level)*Factor_Vol*1.1+Dead_Vol))
            Cell_Volumes.append(JMP_Sheet.col_values(i+1,1).count(Coded_Level)*Cell_Vol)

    if (((Summary*Factor_Vol) > Epitube_Vol).any(axis=None)):
        messagebox.showinfo("Error", "The Volume of a single tube at a specific level is too high for the Epitube, please block your experiment. (Modifications can be made to the program to adjust this by splitting the Levels into two tubes if needed)")
        quit()

    #Let the user know the experimient has too much runs and not enough space on the epmotion.
    if len(Levels)*len(Factors) > 48-Len(Factors):
        messagebox.showinfo("Error", "Your Experiment has a large amount of Factors and Runs, which is beyond the space avalible in the EpMotion. I recommended blocking your experiments to avoid nuisance factors")
        quit()

    num_Levels = len(Levels)
    Factors = X_Vars.columns

    Source = pd.DataFrame(Layout[:len(Factors)], index = Factors)

    FileName, Screwup = Get_Concentrations(Levels, Factors)
    if (Screwup):
        while (Screwup):
            messagebox.showerror("Error", "A factor source value is lower than it's level dilution value. Please fill it again")
            FileName, Screwup = Get_Concentrations(Levels, Factors,True,name)

    Dilution_Locations, Dilution_Commands, Needed_Vol, Media_Vol_Needed, Cereal_Commands, Source_Location = Dilute(Layout, Source, Levels, Factors, Summary, Dead_Vol, FileName)

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
            for k in range (1,len(Row) - num_Y_Vars):
                Feed_Location =  Dilution_Locations[num_Levels*(k-1)+Levels.index(Row[k])] # a bit much should simplfy.
                Source_Rack = Source_Location[num_Levels*(k-1)+Levels.index(Row[k])]
                Well_Location = Plates[i].Wells[j]
                Plates[i].Commands.append([Source_Rack,Feed_Location,2,Well_Location,Factor_Vol-Cell_Vol, "TS_50"])
        Trials += j

    return Plates,len(Dilution_Locations), Dilution_Commands, Source, Needed_Vol, Rack, Media_Vol_Needed, Cereal_Commands

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
def Dilute(Layout, Source, Levels, Factors, Summary, Dead_Vol, FileName):

    Min_Dilution_Vol = 0.5 # Minimum Volume of dilution before it resorts to cereal dilutions
    Dilutions = []
    Commands = []
    Source_Location = []
    Cereal_Commands = []
    cereal_Dilutions = 0
    Needed_Vol = []

    #Find the dilution which should be done manually so as to not waste any factor
    Dilution_Conc = pd.read_csv(Path.cwd() / FileName, index_col = 0)

    #Make sure the Source Concentration is high enough since the concentrations will be diluted by the other factors added, it makes sure that they are 4 times as concentrated
    if ((1/Dilution_Conc.iloc[:,1:].div(Dilution_Conc.Source, axis=0))/len(Factors) < 1 ).any(axis=None):
        messagebox.showinfo("Error", Factor + "'s Source concentration is too low for the concentration need to enter into the plate")
        quit()

    #mass/moles of Factors used  Sum(Vi*Ci = mass/mols)
    Volume_Times_Conc = pd.DataFrame(Dilution_Conc.drop(["Source"], axis =1)
                                    .values*(Summary*Factor_Vol+Dead_Vol),
                                    columns = Summary.columns,
                                    index = Summary.index)

    #Volumes of Manual Concentration which should be made
    Manual_Volumes = Volume_Times_Conc.sum(axis = 1)/(Dilution_Conc.iloc[:,len(Levels)]*1.1)
    Manual_Concentrations = Volume_Times_Conc.sum(axis = 1)*len(Factors)/Manual_Volumes
    Manual_Volumes = Manual_Volumes*1.1**2
    Dilution_Amount = Manual_Volumes*Manual_Concentrations/Dilution_Conc.Source
    Needed_Vol = [list(a) for a in zip(Dilution_Amount.tolist(), Manual_Volumes.tolist())]

    if (Manual_Concentrations > Dilution_Conc.Source).any():
        messagebox.showinfo("Error", "The program's recommended manual dilution is higher than the source concentration. This is generally because you have too many runs and a low source concentration.")
        quit()

    Dilution_Conc["Manually Diluted Concentration"] = Manual_Concentrations.tolist()
    Dilution_Conc.to_csv(Path.cwd() / FileName)

    Dilution_Liquid_Needed = 0

    with open(Path.cwd() / name) as Dilution_Concentrations:
        csv_reader = csv.reader(Dilution_Concentrations, delimiter=',')
        line_count = 0
        for row in csv_reader:
            if line_count == 0:
                line_count += 1
            else:
                for i in range(len(Levels)):
                    Index = (line_count-1)*len(Levels)+i+cereal_Dilutions
                    Rack = 1
                    if (Index < 24):
                        Well_Location = Layout[(line_count-1)*len(Levels)+i+cereal_Dilutions]
                        Destination = 2
                        Source_Location.append(3)
                    else:
                        Well_Location = Layout[((line_count-1)*len(Levels)+i+cereal_Dilutions) - 24 + len(Factors)]
                        Destination = 1
                        Source_Location.append(1)
                    Diluted_Factor_Needed = User_Vol[(line_count-1)*len(Levels)+i]
                    Volume_to_add = (float(row[i+2])*len(Factors))/Manual_Concentrations[line_count-1]*Diluted_Factor_Needed
                    cereal_Run = False
                    if (Volume_to_add < Min_Dilution):
                        Diluted_Factor_Needed = 0.5 * Diluted_Factor_Needed/Volume_to_add
                        Volume_to_add = 0.5
                    if (Diluted_Factor_Needed > Epitube_Vol):
                        Diluted_Factor_Needed = Epitube_Vol + Cell_Volumes[(line_count-1)*len(Levels)+i]
                        cereal_Run = True
                        cereal_Dilutions +=1
                    #How much liquid needs to be added to top up to the correct concentration
                    Top_up_Volume = Diluted_Factor_Needed - Volume_to_add - Cell_Volumes[(line_count-1)*len(Levels)+i]
                    Dilution_Liquid_Needed += Top_up_Volume
                    if not cereal_Run:
                        Commands.extend(Fill_Up( Source[line_count-1][1], Top_up_Volume, Volume_to_add, Well_Location, Destination, Rack))
                    else:
                        Cereal_Commands.extend(Fill_Up( Source[line_count-1][1], Top_up_Volume, Volume_to_add, Well_Location, Destination, Rack))

                    while(cereal_Run):
                        Index = (line_count-1)*len(Levels)+i+cereal_Dilutions
                        cereal_Run = False
                        Cereal_Location = Well_Location
                        Rack = 3
                        if Index > 24:
                            Rack = 1
                        if (Index < 24):
                            Well_Location = Layout[Index]
                            Destination = 2
                            Source_Location[(line_count-1)*len(Levels)+i] = 3
                        else:
                            Well_Location = Layout[Index - 24 + len(Factors)]
                            Destination = 1
                            Source_Location[(line_count-1)*len(Levels)+i] = 1
                        Diluted_Factor_Needed = User_Vol[(line_count-1)*len(Levels)+i]
                        Volume_to_add = (float(row[i+2])*len(Factors))/(Manual_Concentrations[line_count-1]*Volume_to_add/Epitube_Vol)*Diluted_Factor_Needed
                        if (Volume_to_add < Min_Dilution):
                            Volume_to_add = Volume_to_add * Epitube_Vol/User_Vol[(line_count-1)*len(Levels)+i]
                            Diluted_Factor_Needed = Epitube_Vol + Cell_Volumes[(line_count-1)*len(Levels)+i]
                        if (Volume_to_add < Min_Dilution):
                            Volume_to_add = 0.5
                            cereal_Run = True
                            cereal_Dilutions +=1
                        Top_up_Volume = Diluted_Factor_Needed - Volume_to_add - Cell_Volumes[(line_count-1)*len(Levels)+i]
                        if not cereal_Run:
                            Commands.extend(Fill_Up(Cereal_Location, Top_up_Volume, Volume_to_add, Well_Location, Destination, Rack))
                        else:
                            Cereal_Commands.extend(Fill_Up(Cereal_Location, Top_up_Volume, Volume_to_add, Well_Location, Destination, Rack))
                    Dilutions.append(Well_Location)
                line_count += 1

    return Dilutions, Commands, Needed_Vol, ceil(Dilution_Liquid_Needed/1000)*1100, Cereal_Commands, Source_Location

def Fill_Up(Source_Rack, Source, Top_up_Volume, Volume_to_add, Destination_Rack, Well_Location):
    Commands = []
    if (Top_up_Volume%10 >= 0.5) and (Top_up_Volume < 10):
        Commands.append([2,1,Destination_Rack,Well_Location,Top_up_Volume%10, "TS_10"])
        Top_up_Volume = Top_up_Volume - Top_up_Volume%10
    if (Top_up_Volume % 50 >= 0.5) and (Top_up_Volume < 50):
        Commands.append([2,1,Destination_Rack,Well_Location,Top_up_Volume%50, "TS_50"])
        Top_up_Volume = Top_up_Volume - Top_up_Volume%50
    if (Volume_to_add % 10 >= 0.5) and (Volume_to_add < 10):
        Commands.append([Source_Rack,Source,Destination_Rack,Well_Location,Volume_to_add%10, "TS_10"])
        Volume_to_add = Volume_to_add - Volume_to_add%10
    if (Volume_to_add % 50 >= 0.5) and (Volume_to_add < 50):
        Commands.append([Source_Rack,Source,Destination_Rack,Well_Location,Volume_to_add%50, "TS_50"])
        Volume_to_add = Volume_to_add - Volume_to_add%50
    if (Volume_to_add % 1000 >= 0.5):
        Commands.append([Source_Rack,Source,Destination_Rack,Well_Location,Volume_to_add%1000, "TS_1000"])
        Volume_to_add =  Volume_to_add - Volume_to_add%1000
    if (Top_up_Volume % 1000 >= 0.5):
        Commands.append([2,1,Destination_Rack,Well_Location,Top_up_Volume%1000, "TS_1000"])
        Top_up_Volume = Top_up_Volume - Top_up_Volume%1000
    while (Volume_to_add >= 1000):
        Commands.append([Source_Rack,Source,Destination_Rack,Well_Location,1000, "TS_1000"])
        Volume_to_add =  Volume_to_add - 1000
    while(Top_up_Volume >= 1000):
        Commands.append([2,1,Destination_Rack,Well_Location,1000, "TS_1000"])
        Top_up_Volume = Top_up_Volume - 1000

    return Commands

class Plate():

    def __init__(self, Plate_Type,Edge):
        if(Plate_Type == "96 Well"):
            self.Rows = 8
            self.Cols = 12
            self.Vol = 100
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
                Well_Value = list(string.ascii_uppercase)[j+self.num_Edgewells]+str(i+1+self.num_Edgewells)
                self.Wells.append(Well_Value)
