from Epmotion_GUI import Get_Data
from Convert import *
from Output import *
import pandas as pd
import xlrd
import os
from pathlib import Path
import numpy

#Opens the Workbook and first sheet using the xlrd
def JMP_Input(Input_File):

    Input_workbook = xlrd.open_workbook(Input_File)
    Input_Sheet = Input_workbook.sheet_by_index(0)

    return Input_Sheet

if __name__ == '__main__':

    CREATE AN EPMOTION CLASS WITH THE INFORMATION THE USER ENTERS

    #Layout of 24 well racks in the EpMotion
    Rack_Layout = numpy.arange(1,25,1)

    Input, Plate, Well_Volume, Edge_Num, Dead_Vol, Added_Cell_Vol  = Get_Data()
    #Check if the experiment is blocked, and if so divide them into blocks
    Temp_Files =  [Input]
    num_Blocks = 0
    Experiment = pd.read_excel(Input)
    if "Block" in Experiment.columns:
        Temp_Files = []
        Blocks = Experiment.groupby("Block")
        num_Blocks = len(Experiment.groupby("Block"))
        for Block_num, Block in Blocks:
            Block.drop("Block", axis = 1).to_excel("Temp_File" + Block_num + ".xlsx")
            Temp_Files.append("Temp_File" + Block_num + ".xlsx")

    Folders = Produce_Output_Folder(num_Blocks + 1)

    for i, Folder in enumerate(Folders[1:]):
        JMP_Sheet = JMP_Input(Temp_Files[i])
        Output_Plates, Dil_Num, Dil_Commands, Sources, Needed_Vol, Rack, Media_Vol_Needed, Cereal_Commands  = Rearrangment(JMP_Sheet, Rack_Layout,Plate, Well_Volume, Edge_Num, Dead_Vol,Added_Cell_Vol)

        if Rack == 96:
            Rack_Layout = False #False because the rack layout is no longer needed, perhaps change it to actual 96 well layout in future
        if len(Cereal_Commands) != 0: #Only run cereal commands if there are some
            Epmotion_Output(Cereal_Commands,"CEREAL", Folder)
        Mixing_Info = Epmotion_Output(Dil_Commands, "FACTOR", Folder)
        print(Mixing_Info)
        Epmotion_Output(Output_Plates,"PLATE", Folder)
        Protcol_Output(Dil_Num, Sources, Rack_Layout, Folder, Needed_Vol, Media_Vol_Needed)
    IDs = ["Name1", "Name2", "NameN"]
    Experiment_Summary(Folders[0], JMP_Sheet, IDs)
    os.rename(Path.cwd() / "Dilution_Concentrations_SR.csv", Folders[0] / "Dilution_Concentrations_SR.csv")
