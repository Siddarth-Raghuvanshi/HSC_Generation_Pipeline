from Epmotion_GUI import Get_Data
from Convert import *
from Output import *
import pandas as pd
import xlrd
import os
from pathlib import Path

#Opens the Workbook and first sheet using the xlrd
def JMP_Input(Input_File):

    Input_workbook = xlrd.open_workbook(Input_File)
    Input_Sheet = Input_workbook.sheet_by_index(0)

    return Input_Sheet

if __name__ == '__main__':

    #Layout of 24 well racks in the EpMotion
    Rack_Layout = ["A1","A2","A3","A4","A5","A6","B1","B2","B3","B4","B5","B6","C1","C2","C3","C4","C5","C6","D1","D2","D3","D4","D5","D6"]

    Input, Plate, Well_Volume, Edge_Num, Dead_Vol  = Get_Data()
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
        Output_Plates, Dil_Num, Dil_Commands, Sources, Needed_Vol, Rack  = Rearrangment(JMP_Sheet, Rack_Layout,Plate, Well_Volume, Edge_Num, Dead_Vol)

        if Rack == 96:
            Rack_Layout = False #False because the rack layout is no longer needed, perhaps change it to actual 96 well layout in future

        Epmotion_Output(Dil_Commands, "DILUTION", Folder)
        Epmotion_Output(Output_Plates,"PLATE", Folder)
        Protcol_Output(Dil_Num, Sources, Rack_Layout, Folder, Needed_Vol)
    IDs = ["Name1", "Name2", "NameN"]
    Experiment_Summary(Folders[0], JMP_Sheet, IDs)
    os.rename(Path.cwd() / "Dilution_Concentrations_SR.csv", Folders[0] / "Dilution_Concentrations_SR.csv")
