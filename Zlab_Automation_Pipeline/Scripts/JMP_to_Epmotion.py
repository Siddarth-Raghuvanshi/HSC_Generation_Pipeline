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

    #Layout of 24 well racks in the EpMotion
    Rack_Layout = numpy.arange(1,25,1)
    #Get Data from the GUI
    Input, Plate, Well_Volume, Edge_Num, Dead_Vol, Added_Cell_Vol  = Get_Data()

    #Create a new EpMotion to store information from the user and preset values
    Handler_Bing = EpMotion()
    Handler_Bing.Set_UserSpecs(Plate, Well_Volume, Edge_Num, Dead_Vol, Added_Cell_Vol, Rack_Layout)

    #Check if the experiment is blocked, and if so divide them into blocks
    Temp_Files =  [Input]
    num_Blocks = 0
    Experiment = pd.read_excel(Input)
    if "Block" in Experiment.columns:
        Temp_Files = []
        Blocks = Experiment.groupby("Block")
        num_Blocks = len(Experiment.groupby("Block"))
        for Block_num, Block in Blocks:
            Temp_Files.append(Block.drop("Block", axis = 1))

    Folders = Produce_Output_Folder(num_Blocks + 1)

    for i, Folder in enumerate(Folders[1:]):
        DOE_Table = Temp_Files[i]
        Output_Plates, Needed_Vol, Factor_Commands, Cereal_Commands = Rearrangment(DOE_Table, Handler_Bing)

        if len(Cereal_Commands) != 0: #Only run cereal commands if there are some
            Epmotion_Output(Cereal_Commands,"CEREAL", Folder)

        Mixing_Info = Epmotion_Output(Factor_Commands, "FACTOR", Folder)
        print(Mixing_Info)
        Epmotion_Output(Output_Plates,"PLATE", Folder)
        Protcol_Output(Folder, Needed_Vol, Handler_Bing)

    IDs = ["Name1", "Name2", "NameN"]
    Experiment_Summary(Folders[0], Experiment, IDs)
    os.rename(Path.cwd() / "Dilution_Concentrations_SR.csv", Folders[0] / "Dilution_Concentrations_SR.csv")

#EpMotion Class with information about the system
class EpMotion():

    def __init__(self):
        self.Min_Dilution_Vol = 0.5
        self.Epitube_Vol = 1600
        self.media_used = 0
        self.Source_Locations = None
        self.Factor_Locations = None

    def Set_UserSpecs(self,Plate, Well_Volume, Edge_Num, Dead_Volume, Volume_of_Cells, Rack_Layout):
        self.Well_Vol = Well_Volume
        self.Dead_Vol = Dead_Vol
        self.Cell_Volume = Volume_of_Cells
        self.EdgeNum =  Edge_Num
        self.Rack_Layout = Rack_Layout
        self.Plate = Plate

     #Assumes just a single Factor rack and one source rack in the Epmotion
    def Factor_Space_Used(self, Num_Factors):
        self.Space = np.append(Layout, Layout)
        self.SpaceLeft = np.append(Layout, Layout[Num_Factors:])

    def Assign_Space(self, Number):
        Output = self.SpaceLeft[:Number]
        self.SpaceLeft = self.SpaceLeft[Number:]
        return Output

    def Media_Used(Volume):
        self.media_used += Volume
