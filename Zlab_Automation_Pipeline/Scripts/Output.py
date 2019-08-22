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

def Experiment_Summary(Folder_Name, JMP_Sheet,IDs):
    Summary_Template = xlrd.open_workbook(Path.cwd() / "Templates/Automation Summary Template.xlsx")
    Summary_File = copy(Summary_Template)
    Output_Sheet = Summary_File.add_sheet("Well Map")
    Output_Sheet.write(0,0, "Run Name")
    for i,name in enumerate(IDs):
        Output_Sheet.write(i+1,0, name)
    for i in range(JMP_Sheet.ncols):
        for j, Cell in enumerate(JMP_Sheet.col_values(i)):
            Output_Sheet.write(j, i + 2, Cell)
    Summary_File.save(Folder_Name / "Summary.xls")

#Outputs a CSV that is usable by Epmotion
def Epmotion_Output(Info,Purpose, Folder_Name):
    #Header
    Header_Data = [["Rack","Source","Rack","Destination","Volume","Tool"]]


    if (Purpose == "PLATE"):
        for i in range(len(Info)):
            #Creates a CSV file and feeds in the new data for Epmotion
            Optimized_Commands = Script_Optimizer(Info[i].Commands, Purpose)
            name = Folder_Name / ("EpMotion/Epmotion_Plate_"+ str(i+1) + "_SR.csv")
            with open(name, "w") as csvFile:
                writer = csv.writer(csvFile, lineterminator = "\n")
                writer.writerows(Header_Data)
                #writer.writerows(Info[i].EdgeData) No longer needed as it is incorporated into our EpMotion Templates, but its good to keep for testing and possible future use
                writer.writerows(Optimized_Commands)
            csvFile.close()
    elif (Purpose == "FACTOR"):
        Optimized_Commands, Tools, Mixing_Info = Script_Optimizer(Info, Purpose)
        Names = [Folder_Name / "EpMotion/Factor_Prep_SR_Base.csv", Folder_Name / "EpMotion/Factor_Prep_SR_P10.csv", Folder_Name / "EpMotion/Factor_Prep_SR_Final.csv", Folder_Name / "EpMotion/Factor_Prep_SR_Mixing.csv" ]
        for i,Name in enumerate(Names):
            if not Optimized_Commands[i]:
                continue
            with open(Name , "w") as csvFile:
                writer = csv.writer(csvFile, lineterminator = "\n")
                writer.writerows(Header_Data)
                writer.writerows(Optimized_Commands[i])
            csvFile.close()
        return Mixing_Info
    else:
        Optimized_Commands, Tools, _  = Script_Optimizer(Info, Purpose)
        Names = [Folder_Name / "EpMotion/Dilution_Prep_Base.csv", Folder_Name / "EpMotion/Dilution_Prep_SR_P10.csv", Folder_Name / "EpMotion/Dilution_Prep_SR_Mixing.csv"]
        for i,Name in enumerate(Names):
            if not Optimized_Commands[i]:
                continue
            with open(Name , "w") as csvFile:
                writer = csv.writer(csvFile, lineterminator = "\n")
                writer.writerows(Header_Data)
                writer.writerows(Optimized_Commands[i])
            csvFile.close()

#Optimizes the layout, so that the EpMotion would use the least amount of times and time
def Script_Optimizer(Info, Purpose):
    Data = pd.DataFrame(Info, columns = ["S_Rack","Source","D_Rack","Destination","Volume","Tool"])

    if(Purpose == "PLATE"):
        Plate_Commands = Data.sort_values(["S_Rack", "Source"])
        return Plate_Commands.values.tolist()
    else:
        Base_Values = ((Data.sort_values(by = "Volume", ascending = False)
                    .drop_duplicates(subset = ['D_Rack', "Destination"]))
                    .sort_values(by = ["Tool","S_Rack", "Source"])
                    .reset_index(drop = True))

        Tools_Used = Base_Values["Tool"].unique()

        Factor_Top_Up_Values = (Data.append(Base_Values)
                    .drop_duplicates(keep = False)
                    .sort_values(by = ["Tool", "S_Rack", "Source"])
                    .reset_index(drop = True))

        if not Purpose == "CEREAL" :
            Mixing_Commands = Base_Values.copy()
            Mixing_Commands.Volume = Mixing_Commands.Volume/4
            #Also change the values of the tool if there is a volume change
            Base_Values.Volume = Base_Values.Volume*3/4
            Mixing_Info = [Mixing_Commands.Volume.min(), Base_Values.Volume.max()/Mixing_Commands.Volume.min()]
        else:
            Mixing_Commands = pd.DataFrame()
            Mixing_Info = []

        P_10_Commands = Factor_Top_Up_Values[Factor_Top_Up_Values["Tool"] == "TS_10"]
        Large_P_Commands = Factor_Top_Up_Values[Factor_Top_Up_Values["Tool"] != "TS_10"]

        Tools_Used = np.append(Tools_Used, (Factor_Top_Up_Values["Tool"].unique()))

        return [Base_Values.values.tolist(), P_10_Commands.values.tolist(), Large_P_Commands.values.tolist(), Mixing_Commands.values.tolist()], np.unique(Tools_Used).tolist(), Mixing_Info

#Outputs a written protocol for original rack placement
def Protcol_Output(Dilutions_Num, Source, Rack_Layout, Folder_Name, Needed_Vol, Media_Vol_Needed):
    name = Folder_Name / "Protocol_SR.txt"
    File =  open(name,"w")

    File.write("Epmotion Protocol" + str(datetime.now().hour) + "_" + str(datetime.now().minute) + "_" + str(datetime.now().second) ""\n\n")
    File.write("DILUTION RACK PLACEMENT \n\n")
    File.write("1. Place a 7 slot reservoir rack into the EpMotion in C2\n")
    if Rack_Layout.all():
        File.write("2. Place 2 24-well racks into the EpMotion in A2 and B2\n")
    else:
        File.write("2. Place a 24-well rack and a 96-well plate into the EpMotion in A2 and B2\n")
    File.write("3. Place a box of 10 and 50 ul tips into the EpMotion in B1 and C1\n")
    File.write("4. Ensure that there is a space for a plate in the front of the EpMotion in TMX\n")
    File.write("5. Place a TS_10 and a TS_50 into the EpMotion\n\n")

    File.write("MANUAL DILUTIONS \n\n")
    for i,Factor in enumerate(Source):
        Initial_Vol = Needed_Vol[i][0]
        Media_Volume = Needed_Vol[i][1]
        if Needed_Vol[i][0] < 1:
            Media_Volume = 1/(Needed_Vol[i][0]/Needed_Vol[i][1])
            Initial_Vol =  1
            if Media_Volume > 1600:
                Initial_Vol = Needed_Vol[i][0]*1600
                File.write("%d a). Dilute Stock %s by adding %.2f ul into %.2f ul of Media\n" % (i + 1, Factor[0], Initial_Vol, 1599))
                File.write("%d b). Dilute the created %s dilution by adding %.2f ul into %.2f ul of Media\n" % (i + 1, Factor[0], Initial_Vol, Needed_Vol[i][1] - Initial_Vol))
                File.write("Test the dilution out real quick, I haven't tested this before")
                continue
        File.write("%d. Dilute Stock %s by adding %.2f ul into %.2f ul of Media\n" % (i + 1, Factor[0], Initial_Vol, Media_Volume - Initial_Vol))
    File.write("\n\n")

    File.write("LIQUID LAYOUT \n\n")
    File.write("1. Place a 25 ml reservoir containing " + str(Media_Vol_Needed/1000) + " ml of dilution liquid into the 1st slot in the reservoir\n")
    File.write("2. Place a 25 ml reservoir containing edge liquid into the 2nd slot in the reservoir\n")
    for i in range(len(Source)):
        File.write("%d. Place the Diluted %s into the %s well in the first rack\n" % ( i+3, Source[i][0], Source[i][1]))
    if Rack_Layout.all():
        File.write("%d. Place 24 sterile Epitubes into the second rack from %s to %s\n" %( i+3, Rack_Layout[0], Rack_Layout[23]))
        File.write("%d. Place %d sterile Epitubes into the second rack from %s to %s\n\n" %(len(Source) + 4 , Dilutions_Num - 24, Rack_Layout[len(Source)], Rack_Layout[Dilutions_Num-24]))
    else:
        File.write("\n")


    File.write("EPBLUE PROTOCOL \n\n")
    File.write("1. Create a new application.\n")
    File.write("2. Insert three Sample transfer command.\n\n")
    File.write("3. Insert three STOP (Us) transfer command.\n\n")
    File.write("DILUTION PROTOCOL \n")
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

    File.write("EXPERIMENT TIP PLACEMENT \n\n")
    File.write("1. Place a TS_50 and a TS_1000 into the EpMotion\n")
    File.write("2. Place a box of 50 and 1000 ul tips into the EpMotion\n" )

    File.close()
