#A program to take in the JMP output and convert it into an input for Epmotion machine

#!/usr/bin/python
import csv
import numpy as np
import string
from tkinter import messagebox
import pandas as pd
from pathlib import Path

#Takes the information from the JMP file and places everything into a format readable by Epmotion
def Rearrangment(Experiment_Matrix, Handler_Bing):

    #Set Variables for readability
    Y_Vars = Experiment_Matrix.columns[Experiment_Matrix.isna().all()].tolist()
    X_Vars = Experiment_Matrix.drop(Y_Vars, axis =1)
    Factor_Vol = Handler_Bing.Well_Vol/len(X_Vars.columns)  #Volume of factors the EpMotion
    Cell_Vol = Handler_Bing.Cell_Volume/len(X_Vars.columns)
    Factors = X_Vars.columns
    Levels = set()

    #Used to take into account fractional factorial, i.e. uses all of the levels created by all the factors
    [Levels.update(X_Vars[Factor].unique().tolist()) for Factor in Factors]
    Levels = sorted(Levels)

    #Find the volume for each Level which will be tested
    Summary = pd.concat([X_Vars[Factor].value_counts() for Factor in X_Vars.columns], axis =1).T
    Factor_Volume_Frame = Summary*Factor_Vol
    Cell_Volume_Frame = Summary*Cell_Vol

    #Checks to ensure that each volume needed isn't bigger than the Epitube
    if (((Factor_Volume_Frame) > Handler_Bing.Epitube_Vol).any(axis=None)):
        messagebox.showinfo("Error", "The Volume of a single tube at a specific level is too high for the Epitube, please block your experiment. (Modifications can be made to the program to adjust this by splitting the Levels into two tubes if needed)")
        quit()
    #Let the user know the experimient has too much runs and not enough space on the epmotion.
    if len(Levels)*len(Factors) > 48-len(Factors):
        messagebox.showinfo("Error", "Your Experiment has a large amount of Factors and Runs, which is beyond the space avalible in the EpMotion. I recommended blocking your experiments to avoid nuisance factors")
        quit()

    #Get a file containing the dilution concentrations and make it run until the user fixes it
    FileName, Screwup = Get_Concentrations(Factor_Volume_Frame)
    if (Screwup):
        while (Screwup):
            messagebox.showerror("Error", "A factor source value is lower than it's level dilution value. Please fill it again")
            FileName, Screwup = Get_Concentrations(Factor_Volume_Frame, True,name)

    #Calculations for diluting manually diluting the factor and  the factors so to each level
    Dilution_Conc_Frame, Needed_Vol = Man_Dilution_Calc(Factor_Volume_Frame, Handler_Bing, FileName)
    Dilution_Information, Cereal_Dilution_Commands = Dilution_Prep(Dilution_Conc_Frame, Factor_Volume_Frame, Handler_Bing)
    Dilution_Information.append(Cell_Volume_Frame)
    Factor_Commands = Factor_Dilution_Commands(Dilution_Information, Handler_Bing)

    #Create Commands for Factorial design.
    Plate_wells = len(Plate(Handler_Bing.Plate, Handler_Bing.EdgeNum).Wells)
    Num_Plates = len(X_Vars.index)
    Plates = []
    Experiments = 1
    for row_index in range(len(X_Vars.index)):
        Well = row_index % Plate_wells
        if Well == 0:
            Plates.append(Plate(Handler_Bing.Plate, Handler_Bing.EdgeNum))
        for column_index in range(len(X_Vars.columns)):
            Factor = X_Vars.columns[column_index]
            Level = X_Vars.iloc[row_index, column_index]
            Source, Source_Rack = Handler_Bing.Factor_Locations.loc[Factor,Level]
            Well_Location = Plates[-1].Wells[Well]
            Plates[-1].Commands.append([Source_Rack, Source, 2, Well_Location, Factor_Vol-Cell_Vol, "TS_50"])

    return Plates, Needed_Vol, Factor_Commands, Cereal_Dilution_Commands

#Create a CSV which contains the user specified dilutions, and asks the user to input these concentrations
def Get_Concentrations(Factor_Volume_Frame,Screwup = False, name = ""):

    if not Screwup:
        Factors = Factor_Volume_Frame.index
        Levels = Factor_Volume_Frame.columns.tolist()

        Header = [["Factors", "Source"] + Levels]
        name = "Dilution_Concentrations_SR.csv"
        with open(Path.cwd() / name, "w") as csvFile:
            writer = csv.writer(csvFile, lineterminator = "\n")
            writer.writerows(Header)
            for Factor in Factors:
                writer.writerow([Factor])
        csvFile.close()
        messagebox.showinfo("Dilutions", "A File called " + name + " has been created for you to populate with the concentrations of your Factors and Levels, please fill it now")

    input("\nPress Enter to continue once completed...")

    Concentrations = pd.read_csv(Path.cwd() / name, header = 0)
    if not (Concentrations.loc[:, "Source":].idxmax(1) == "Source").all():
        Screwup = True
    else:
        Screwup = False

    return name, Screwup

#Using the user input the function finds out how to manually dilute for the epmotion
def Man_Dilution_Calc(Factor_Volume_Frame, Handler_Bing, FileName):

    #Get some variables
    Levels = Factor_Volume_Frame.columns
    Factors = Factor_Volume_Frame.index

    #Read the user's file
    Dilution_Conc = pd.read_csv(Path.cwd() / FileName, index_col = 0)

    #So since the liquid is going to be diluted into the plate, it must first be more concentrated this is the check
    if ((1/Dilution_Conc.iloc[:,1:].div(Dilution_Conc.Source, axis=0))/len(Factors) < 1 ).any(axis=None):
        messagebox.showinfo("Error", Factor + "'s Source concentration is too low for the concentration need to enter into the plate")
        quit()

    #mass/moles of Factors used  Sum(Vi*Ci = mass/mols)
    Volume_Times_Conc = pd.DataFrame(Dilution_Conc.drop(["Source"], axis =1)
                                    .values*(Factor_Volume_Frame+Handler_Bing.Dead_Vol),
                                    columns = Factor_Volume_Frame.columns,
                                    index = Factor_Volume_Frame.index)

    #Volumes of Manual Concentration which should be made
    Manual_Volumes = Volume_Times_Conc.sum(axis = 1)/(Dilution_Conc.iloc[:,len(Levels)]*1.1)
    if (Manual_Volumes > Handler_Bing.Epitube_Vol).any():
        Below_Cutoff_Values = Manual_Volumes[Manual_Volumes > Handler_Bing.Epitube_Vol]
        Manual_Volumes = Manual_Volumes.where(Below_Cutoff_Values.isna(), Handler_Bing.Epitube_Vol/(1.1**2))
    Manual_Concentrations = Volume_Times_Conc.sum(axis = 1)*len(Factors)/Manual_Volumes

    #Find the amounts the user needs to add to dilute the factors
    Manual_Volumes = Manual_Volumes*1.1**2
    Dilution_Amount = Manual_Volumes*Manual_Concentrations/Dilution_Conc.Source
    Needed_Vol = [list(a) for a in zip(Dilution_Amount.tolist(), Manual_Volumes.tolist())]

    #Secondary check
    if (Manual_Concentrations > Dilution_Conc.Source).any():
        messagebox.showinfo("Error", "The program's recommended manual dilution is higher than the source concentration. This is generally because you have too many runs and a low source concentration.")
        quit()

    #Add the manual concentrations to the file so the user has a copy of them
    Dilution_Conc["Manually Diluted Concentration"] = Manual_Concentrations.tolist()
    Dilution_Conc.to_csv(Path.cwd() / FileName)

    return Dilution_Conc, Needed_Vol

#Volume calculator for the EpMotion to create the factor dilutions and also Experiment_Matrix to create cereal dilutions tubes in case they are needed
def Dilution_Prep(Dilution_Conc, Factor_Volume_Frame, Handler_Bing):

    #Set the space used for the factors
    Handler_Bing.Factor_Space_Used(len(Factors))

    Factors = Factor_Volume_Frame.index
    Levels = Factor_Volume_Frame.columns
    Cereal_Commands = []

    Total_Volume = (Factor_Volume_Frame+Handler_Bing.Dead_Vol)*1.1

    #Multiply all of the Concentrations by the volume including the dead vol and a 10% buffer
    Vol_Times_Conc = pd.DataFrame(Dilution_Conc.drop(["Source", "Manually Diluted Concentration"], axis =1)
                                    .values*Total_Volume*len(Factors),
                                    columns = Levels,
                                    index = Factors)

    #Find the volume to add from the manual source concentrations get the correct concentrations
    Vol_To_Add = Vol_Times_Conc.div(Dilution_Conc["Manually Diluted Concentration"], axis=0)
    Vol_To_Add = Vol_To_Add.replace(0,np.nan)

    #Scale up the amount of factor and media used to avoid cereal dilutions
    Below_Cutoff_Values = Vol_To_Add[Vol_To_Add < Handler_Bing.Min_Dilution_Vol]
    Scaled_Total_Vol = Total_Volume.div(Below_Cutoff_Values)*Handler_Bing.Min_Dilution_Vol
    Scaled_Total_Vol = Scaled_Total_Vol[Scaled_Total_Vol < Handler_Bing.Epitube_Vol]
    Vol_To_Add = Vol_To_Add.where(Scaled_Total_Vol.isna(), 0.5)
    Total_Volume = Total_Volume.where(Scaled_Total_Vol.isna(),Scaled_Total_Vol.values)

    #Scale the rest of the concentrations below the cutoff to have a total volume of a Epitube
    Below_Cutoff_Values = Vol_To_Add[Vol_To_Add < Handler_Bing.Min_Dilution_Vol]
    Scaled_Vol_Add = Below_Cutoff_Values.div(Total_Volume)*Handler_Bing.Epitube_Vol
    Vol_To_Add = Vol_To_Add.where(Scaled_Vol_Add.isna(), Scaled_Vol_Add)
    Total_Volume = Total_Volume.where(Scaled_Vol_Add.isna(),Handler_Bing.Epitube_Vol)

    #Find the number of cereal dilutions required for each factor and make sure we can fit them on the EpMotion
    Min_Dilution_Series = pd.Series([Handler_Bing.Min_Dilution_Vol]*len(Factors), index=Factors)
    Number_of_Dilutions = pd.Series(np.ceil(
                                    np.log(Min_Dilution_Series.div(Vol_To_Add.min(axis=1)))/
                                    np.log(Handler_Bing.Epitube_Vol/Handler_Bing.Min_Dilution_Vol)))
    if len(Handler_Bing.SpaceLeft) < Number_of_Dilutions.sum() + len(Factors)*len(Levels):
        messagebox.showinfo("Error", "There are too many tubes and too little space in the EpMotion required for cereally diluting the samples")
        quit()

    #Create Dataframes which contains the concentrations and locations of each factor
    Handler_Bing.Source_Locations = pd.DataFrame(Handler_Bing.Assign_Space(len(Factors)),
                             index= Factors,
                             columns = ["Manual Dilution"])
    Handler_Bing.Source_Locations = pd.concat([Handler_Bing.Source_Locations,
                                pd.DataFrame(columns =
                                             range(1,int(Number_of_Dilutions.max())+1))])

    #Add the Dilutions to the above Dataframes and fill out cereal Experiment_Matrix
    Source_Rack = 1 #An assumption being made is that the source always has to be on the first rack
    for j,Factor in enumerate(Number_of_Dilutions.index):
        Source = Handler_Bing.Source_Locations.loc[Factor]["Manual Dilution"]
        for Number in range(int(Number_of_Dilutions.loc[Factor])): #Different Top Up Volume for last dilution
            if np.count_nonzero(Handler_Bing.SpaceLeft == 24) == 2:
                Destination_Rack = 1
            else:
                Destination_Rack = 2
            Well_Location = int(Handler_Bing.Assign_Space(1))
            Top_up_Vol = Handler_Bing.Epitube_Vol - Handler_Bing.Min_Dilution_Vol
            Cereal_Commands.extend(Fill_Up(Source_Rack,
                                           Source,
                                           Top_up_Vol,
                                           Handler_Bing.Min_Dilution_Vol,
                                           Destination_Rack,
                                           Well_Location,
                                           Number))
            Handler_Bing.Media_Used(Top_up_Vol)
            Source = Well_Location
            Handler_Bing.Source_Locations.at[Factor,Number+1] = Well_Location

    Dilution_Information = [Vol_To_Add, Total_Volume]

    return Dilution_Information, Cereal_Commands

# Takes the information and dilutes the factors
def Factor_Dilution_Commands(Dilution_Information, Handler_Bing):

    #extract information from list
    Vol_To_Add = Dilution_Information[0]
    Total_Volume = Dilution_Information[1]
    Cell_Vol_Frame = Dilution_Information[2]

    #define some variables for readability
    Factors = Vol_To_Add.index
    Levels = Vol_To_Add.columns
    Factor_Commands = []

    #Create a dataframe containing the locations of every factor at each level
    Handler_Bing.Factor_Locations = pd.DataFrame(index = Factors, columns = Levels)

    #Find the volume needed from the cereal dilutions
    Below_Cutoff_Values = Vol_To_Add[Vol_To_Add < Handler_Bing.Min_Dilution_Vol] #Need to recalculate after the scale-up
    Min_Dilution_Frame = pd.DataFrame(Handler_Bing.Min_Dilution_Vol, index = Vol_To_Add.index, columns = Vol_To_Add.columns)
    Dilution_Tubes = pd.DataFrame(np.ceil( #Find which tube needs to be used
                                  np.log(Min_Dilution_Frame.div(Below_Cutoff_Values))/
                                  np.log(Handler_Bing.Epitube_Vol/Handler_Bing.Min_Dilution_Vol)))
    Dilution_Frame = Min_Dilution_Frame*(Handler_Bing.Epitube_Vol/(Handler_Bing.Min_Dilution_Vol*Handler_Bing.Min_Dilution_Vol))
    Diluted_Vol_To_Add =  Below_Cutoff_Values*(Dilution_Frame.pow(Dilution_Tubes))
    Vol_To_Add = Vol_To_Add.where(Diluted_Vol_To_Add.isna(), Diluted_Vol_To_Add)
    Dilution_Tubes.fillna(0, inplace = True)

    #Check to make sure that the factors and dilutions haven't gone into the second rack
    if np.count_nonzero(Handler_Bing.SpaceLeft == 24) == 2:
                Source_Rack = 1
    #Should not be an issue, but a check is always good, if this error appears, modify the code to allow Source_Rack to be = 2
    else:
        messagebox.showinfo("Error", "There is not enough space as source tubes would have to be placed in the second rack")
        quit()

    #Replace the Nan values with 0 to allow the program to work
    Vol_To_Add = Vol_To_Add.replace(np.nan, 0)

    for i,Factor in enumerate(Factors):
        for j,Level in enumerate(Levels):
            if np.count_nonzero(Handler_Bing.SpaceLeft == 24) == 2:
                Destination_Rack = 1
            else:
                Destination_Rack = 2
            Source = Handler_Bing.Source_Locations.iloc[i, int(Dilution_Tubes.loc[Factor,Level])]
            Destination = int(Handler_Bing.Assign_Space(1))
            Top_up_Vol = Total_Volume.loc[Factor,Level]-Vol_To_Add.loc[Factor,Level]-Cell_Vol_Frame.loc[Factor,Level]
            Factor_Commands.extend(Fill_Up(Source_Rack,
                                           Source,
                                           Top_up_Vol,
                                           Vol_To_Add.loc[Factor,Level],
                                           Destination_Rack,
                                           Destination))
            Handler_Bing.Media_Used(Top_up_Vol)
            Handler_Bing.Factor_Locations.at[Factor,Level] = [Destination, Destination_Rack]

    return Factor_Commands

#Function to take information and fill out a list to easily fill out a CSV
def Fill_Up(Source_Rack, Source, Top_up_Volume, Volume_to_add, Destination_Rack, Well_Location, Round = None):
    Experiment_Matrix = []
    if (Top_up_Volume%10 >= 0.5) and (Top_up_Volume < 10):
        Experiment_Matrix.append([2,1,Destination_Rack,Well_Location,Top_up_Volume%10, "TS_10", Round])
        Top_up_Volume = Top_up_Volume - Top_up_Volume%10
    if (Top_up_Volume % 50 >= 0.5) and (Top_up_Volume < 50):
        Experiment_Matrix.append([2,1,Destination_Rack,Well_Location,Top_up_Volume%50, "TS_50", Round])
        Top_up_Volume = Top_up_Volume - Top_up_Volume%50
    if (Volume_to_add % 10 >= 0.5) and (Volume_to_add < 10):
        Experiment_Matrix.append([Source_Rack,Source,Destination_Rack,Well_Location,Volume_to_add%10, "TS_10", Round])
        Volume_to_add = Volume_to_add - Volume_to_add%10
    if (Volume_to_add % 50 >= 0.5) and (Volume_to_add < 50):
        Experiment_Matrix.append([Source_Rack,Source,Destination_Rack,Well_Location,Volume_to_add%50, "TS_50", Round])
        Volume_to_add = Volume_to_add - Volume_to_add%50
    if (Volume_to_add % 1000 >= 0.5):
        Experiment_Matrix.append([Source_Rack,Source,Destination_Rack,Well_Location,Volume_to_add%1000, "TS_1000", Round])
        Volume_to_add =  Volume_to_add - Volume_to_add%1000
    if (Top_up_Volume % 1000 >= 0.5):
        Experiment_Matrix.append([2,1,Destination_Rack,Well_Location,Top_up_Volume%1000, "TS_1000", Round])
        Top_up_Volume = Top_up_Volume - Top_up_Volume%1000
    while (Volume_to_add >= 1000):
        Experiment_Matrix.append([Source_Rack,Source,Destination_Rack,Well_Location,1000, "TS_1000", Round])
        Volume_to_add =  Volume_to_add - 1000
    while(Top_up_Volume >= 1000):
        Experiment_Matrix.append([2,1,Destination_Rack,Well_Location,1000, "TS_1000", Round])
        Top_up_Volume = Top_up_Volume - 1000

    return Experiment_Matrix

#class for storing information about the plate
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
        for j in range(self.Rows - (self.num_Edgewells * 2)):
            for i in range(self.Cols - (self.num_Edgewells * 2)):
                Well_Value = list(string.ascii_uppercase)[j+self.num_Edgewells]+str(i+1+self.num_Edgewells)
                self.Wells.append(Well_Value)
