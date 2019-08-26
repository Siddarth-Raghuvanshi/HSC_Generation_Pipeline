from datetime import datetime
import csv
import sys
import shutil
import pandas as pd
import numpy as np
import os
import xlrd
import xlwt
from xlutils.copy import copy
from pathlib import Path

#Produces a collection of output folders and files which are required
def Produce_Output_Folder(Blocks):
    Folder_Name = Path.cwd() / ("Experiment_Files_" + "_" + str(datetime.now().date()) + "_" + str(datetime.now().hour) + "_" + str(datetime.now().minute) + "_" + str(datetime.now().second) + "_SR")
    os.mkdir(Folder_Name)
    Block_Names = [Folder_Name]
    for Num_Block in range(Blocks):
        os.mkdir(Folder_Name / ("Block" + str(Num_Block + 1)))
        os.mkdir(Folder_Name / ("Block" + str(Num_Block + 1)+ "/EpMotion"))
        os.mkdir(Folder_Name / ("Block" + str(Num_Block + 1) + "/Cytoflex"))
        Block_Names.append(Folder_Name / ("Block" + str(Num_Block + 1)))

    return Block_Names

def Experiment_Summary(Folder_Name, Experiment):
    Experiment.to_excel(Folder_Name / "Summary.xls")

#Outputs a CSV that is usable by Epmotion
def Epmotion_Output(Info,Purpose, Folder_Name):
    #Header
    Header_Data = [["Rack","Source","Rack","Destination","Volume","Tool"]]

    if (Purpose == "PLATE"):
        for i in range(len(Info)):
            #Creates a CSV file and feeds in the new data for Epmotion
            name = Folder_Name / ("EpMotion/Epmotion_Plate_"+ str(i+1) + "_SR.csv")
            with open(name, "w") as csvFile:
                writer = csv.writer(csvFile, lineterminator = "\n")
                writer.writerows(Header_Data)
                #writer.writerows(Info[i].EdgeData) No longer needed as it is incorporated into our EpMotion Templates, but its good to keep for testing and possible future use
                writer.writerows(Info[i].Commands)
            csvFile.close()
    elif (Purpose == "FACTOR"):
        Info = Script_Optimizer(Info, Purpose)
        Names = [Folder_Name / "EpMotion/Factor_Prep_SR_Base.csv", Folder_Name / "EpMotion/Factor_Prep_SR.csv", Folder_Name / "EpMotion/Factor_Prep_SR_Mixing.csv" ]
        for i,Name in enumerate(Names):
            with open(Name , "w") as csvFile:
                writer = csv.writer(csvFile, lineterminator = "\n")
                writer.writerows(Header_Data)
                writer.writerows(Info[i])
            csvFile.close()
    else:
        Info = Script_Optimizer(Info, Purpose)
        for i,Round in enumerate(Info):
            Names = [(Folder_Name / ("EpMotion/Dilution_Prep_Base_Block" + str(i+1) + ".csv")), (Folder_Name / ("EpMotion/Dilution_Prep_Mixing_Block" + str(i+1) + ".csv"))]
            for j,Name in enumerate(Names):
                with open(Name , "w") as csvFile:
                    writer = csv.writer(csvFile, lineterminator = "\n")
                    writer.writerows(Header_Data)
                    writer.writerows(Round[j])
                csvFile.close()

#Optimizes the layout, so that the EpMotion would use the least amount of times and time
def Script_Optimizer(Info, Purpose):
    if Purpose == "FACTOR":
        Data = pd.DataFrame(Info, columns = ["S_Rack","Source","D_Rack","Destination","Volume","Tool", "Round"]).drop("Round", axis = 1)

        Base_Values = ((Data.sort_values(by = "Volume", ascending = False)
                    .drop_duplicates(subset = ['D_Rack', "Destination"]))
                    .sort_values(by = ["Tool","S_Rack", "Source"])
                    .reset_index(drop = True))

        Factor_Top_Up_Commands = (Data.append(Base_Values)
                    .drop_duplicates(keep = False)
                    .sort_values(by = ["Tool", "S_Rack", "Source"])
                    .reset_index(drop = True))

        Mixing_Commands = Base_Values.copy()
        Mixing_Commands.Volume = Mixing_Commands.Volume/4
        #Also change the values of the tool if there is a volume change
        Mixing_Commands.loc[Mixing_Commands.Volume < 40, 'Tool'] = "TS_50"
        Base_Values.Volume = Base_Values.Volume*3/4
        Base_Values.loc[Base_Values.Volume < 40, 'Tool'] = "TS_50"

        Info = [Base_Values.values.tolist(),Factor_Top_Up_Commands.values.tolist(), Mixing_Commands.values.tolist()]

    else:
        Data = pd.DataFrame(Info, columns = ["S_Rack","Source","D_Rack","Destination","Volume","Tool", "Round"])

        Rounds = len(Data.Round.unique())
        Info = []

        for Round_Num in range(Rounds):

            Mixing_Commands = ((Data[Data.Round == Round_Num].drop("Round", axis = 1)
                        .sort_values(by = "Volume", ascending = False)
                        .drop_duplicates(subset = ['D_Rack', "Destination"]))
                        .sort_values(by = ["Tool","S_Rack", "Source"])
                        .reset_index(drop = True))

            Factor_Top_Up_Commands = (Data[Data.Round == Round_Num].drop("Round", axis = 1)
                        .append(Mixing_Commands)
                        .drop_duplicates(keep = False)
                        .sort_values(by = "Volume", ascending = False))
                        #.sort_values(by = ["Tool", "S_Rack", "Source"])
                        #.reset_index(drop = True))

            Info.append([Mixing_Commands.values.tolist(), Factor_Top_Up_Commands.values.tolist()])

    return Info

#Outputs a written protocol for original rack placement
def Protcol_Output(Folder_Name, Needed_Vol, Handler_Bing):
    name = Folder_Name / "Protocol_SR.txt"
    File =  open(name,"w")

    File.write("Epmotion Protocol" + str(datetime.now().hour) + "_" + str(datetime.now().minute) + "_" + str(datetime.now().second) + "\n\n")

    File.write("MEDIA AND FACTOR PREPARTION \n\n")
    File.write("1. Create " + str(Handler_Bing.Media_Needed*1.1/1000) + "mls of media\n" )
    for i,Factor in enumerate(Handler_Bing.Source_Locations.index):
        Initial_Vol = Needed_Vol[i][0]
        Media_Volume = Needed_Vol[i][1]
        if Needed_Vol[i][0] < 1:
            Media_Volume = 1/(Needed_Vol[i][0]/Needed_Vol[i][1])
            Initial_Vol =  1
            if Media_Volume > 1600:
                Initial_Vol = Needed_Vol[i][0]*1600
                File.write("%d a). Dilute Stock %s by adding %.2f ul into %.2f ul of media\n" % (i + 2, Factor, Initial_Vol, 1599))
                File.write("%d b). Dilute the created %s dilution by adding %.2f ul into %.2f ul of media\n" % (i + 2, Factor, Initial_Vol, Needed_Vol[i][1] - Initial_Vol))
                File.write("Test the dilution out real quick, I haven't tested this before")
                continue
        File.write("%d. Dilute Stock %s by adding %.2f ul into %.2f ul of Media\n" % (i + 2, Factor, Initial_Vol, Media_Volume - Initial_Vol))
    File.write("\n\n")

    File.write("DILUTION RACK PLACEMENT \n\n")
    File.write("1. Place a 7 slot reservoir rack into the EpMotion in C2\n")
    File.write("2. Place 2 24-well racks into the EpMotion in A2 and B2\n")
    File.write("3. Place a box of 10 and 50 ul tips into the EpMotion in B1 and C1\n")
    File.write("4. Ensure that there is a space for a plate in the front of the EpMotion in TMX\n")
    File.write("5. Place a TS_10 and a TS_50 into the EpMotion\n\n")

    File.write("LIQUID LAYOUT \n\n")
    File.write("1. Place a 25 ml reservoir containing remaining amount of dilution liquid or media into the 1st slot in the reservoir\n")
    File.write("2. Place a 25 ml reservoir containing edge liquid into the 2nd slot in the reservoir\n")
    for i,Factor in enumerate(Handler_Bing.Source_Locations.index):
        File.write("%d. Place the Diluted %s into the %s well in the first rack\n" % ( i+3, Factor, Handler_Bing.Space[i]))
    File.write("%d. Place %d sterile Epitubes into the first rack starting from slot %d and move into the second rack if needed\n" %(i+4, len(Handler_Bing.Space)-len(Handler_Bing.SpaceLeft)-len(Handler_Bing.Source_Locations.index), len(Handler_Bing.Source_Locations.index) ))
    File.write("\n")

    File.write("FACTOR DILUTION PROTOCOL (Only do this if you have Dilution_Prep CSVs)\n\n")
    File.write("1. Open the Factor_Prep Template from the sidr97 user.\n")
    File.write("2. Repeat the steps below for each dilution block\n")
    File.write("3. Fill Out the Dilution_Base.csv into the dilution base command\n")
    File.write("4. Ensure that the large volumes are above the 10 ul in order.\n")
    File.write("5. Ensure that the P10 volumes are water contact.\n")
    File.write("6. Fill out the Mixing.csv into the mixing Command..\n\n")

    File.write("FACTOR PREPARTION PROTOCOL\n\n")
    File.write("1. Open the Factor_Prep Template from the sidr97 user or use the one created above.\n")
    File.write("2. Fill Out the Factor_Base.csv into the Factor Base Command.\n")
    File.write("3. Ensure that the large volumes are above the P10s in order.\n\n")

    File.write("EXPERIMENT PROTOCOL \n")
    File.write("3. Under the Parameter Option, place the Stock 96-well as Source 1\n")
    File.write("4. Under the Parameter Option, place the reservoir as Source 2\n")
    File.write("5. Under the Parameter Option, place the Dilution 24-well as Source 3\n")
    File.write("6. Under the Parameter Option, place the Dilution 24-well as Destination 1\n")
    File.write("7. Under the Parameter Option, place the Stock 24-well as Destination 2\n")
    File.write("8. Ensure all other setting in the transfer command are satisfactory\n")
    File.write("9. Click the Sample transfer command.\n")
    File.write("10. Click on the CSV Command in top bar.\n")
    File.write("11. Select the Dilution_Command CSV and hit open.\n")
    File.write("12. Select the Pipette option and hit OK\n")
    File.write("13. Once finished repeat step 11 with the CSV required for each plate. \n\n")


    File.close()
