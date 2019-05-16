from Epmotion_GUI import Get_Data
from Convert import *
import xlrd

Rack_Layout = ["A1","A2","A3","A4","A5","A6","B1","B2","B3","B4","B5","B6","C1","C2","C3","C4","C5","C6","D1","D2","D3","D4","D5","D6"]


#Opens the Workbook and first sheet using the xlrd
def JMP_Input(Input_File):

    Input_workbook = xlrd.open_workbook(Input_File)
    Input_Sheet = Input_workbook.sheet_by_index(0)

    return Input_Sheet

def Run(Input, Plate, Dilution_Vol, Factor_Vol):

    Output_Plates, Conditions = Rearrangment(Input, Rack_Layout, Dilution_Vol,Plate, Factor_Vol)
    Epmotion_Output(Output_Plates,"PLATE")
    Protcol_Output(Conditions)

if __name__ == '__main__':

    Input, Plate, Dil_Volume, Add_Volume  = Get_Data()
    JMP_Sheet = JMP_Input(Input)
    Run(JMP_Sheet, Plate, Dil_Volume, Add_Volume)
