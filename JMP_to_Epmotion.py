from Epmotion_GUI import Get_Files

#Opens the Workbook and first sheet using the xlrd
def JMP_Input(Input_File):

    Input_workbook = xlrd.open_workbook(Input_File)
    Input_Sheet = Input_workbook.sheet_by_index(0)

    return Input_Sheet

#Create a CSV that dilutes the stock concentrations as per the user input
def Dilution():
    Factor_Conc = []

if __name__ == '__main__':
    Input, Plate = Get_Files()
    JMP_Sheet = JMP_Input(Input)
    Output_Plates = Rearrangment(JMP_Sheet, Rack_Layout)
    Epmotion_Output(Output_Plates)
    Protcol_Output()
