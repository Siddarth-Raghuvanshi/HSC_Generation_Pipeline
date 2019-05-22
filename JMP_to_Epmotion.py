from Epmotion_GUI import Get_Data
from Convert import *
from Output import *
import xlrd

Rack_Layout = ["A1","A2","A3","A4","A5","A6","B1","B2","B3","B4","B5","B6","C1","C2","C3","C4","C5","C6","D1","D2","D3","D4","D5","D6"]


#Opens the Workbook and first sheet using the xlrd
def JMP_Input(Input_File):

    Input_workbook = xlrd.open_workbook(Input_File)
    Input_Sheet = Input_workbook.sheet_by_index(0)

    return Input_Sheet

if __name__ == '__main__':

    Input, Plate, Dil_Volume, Add_Volume, Edge_Volume  = Get_Data()
    JMP_Sheet = JMP_Input(Input)
    Output_Plates, Dil_Num, Dil_Commands, Sources = Rearrangment(JMP_Sheet, Rack_Layout, Dil_Volume,Plate, Add_Volume, Edge_Volume)
    Folder = Produce_Output_Folder()
    Epmotion_Output(Dil_Commands, "DILUTION", Folder)
    Epmotion_Output(Output_Plates,"PLATE", Folder)
    Protcol_Output(Dil_Num, Sources, Rack_Layout, Folder)
