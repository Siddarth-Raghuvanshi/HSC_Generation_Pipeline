from Epmotion_GUI import Get_Files
from Convert import *
import xlrd


#Opens the Workbook and first sheet using the xlrd
def JMP_Input(Input_File):

    Input_workbook = xlrd.open_workbook(Input_File)
    Input_Sheet = Input_workbook.sheet_by_index(0)

    return Input_Sheet

def Run(Input, Dilution_Vol):

    Output_Plates = Rearrangment(Input, Rack_Layout, Dilution_Vol)
    Epmotion_Output(Output_Plates,"PLATE")
    Protcol_Output()

if __name__ == '__main__':
    Input, Plate,Volume = Get_Files()
    JMP_Sheet = JMP_Input(Input)
    Run(JMP_Sheet, Plate, Volume)
