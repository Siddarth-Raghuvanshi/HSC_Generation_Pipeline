from datetime import datetime
import csv
import sys
import os
import shutil
import xlrd
import xlwt
from xlutils.copy import copy

#Produces a collection of output folders and files which are required
def Produce_Output_Folder():
    Folder_Name = "../"+"/Experiment_Files_" + str(datetime.now()) + "_SR"
    os.mkdir(Folder_Name)
    os.mkdir(Folder_Name+"/EpMotion")
    os.mkdir(Folder_Name+"/Cytoflex")
    os.rename("Dilution_Concentrations_SR.csv", Folder_Name+"/EpMotion/Dilution_Concentrations_SR.csv")

    return Folder_Name

def Experiment_Summary(Folder_Name, JMP_Sheet,IDs):
    Summary_Template = xlrd.open_workbook("../Templates/Automation Summary Template.xlsx")
    Summary_File = copy(Summary_Template)
    Output_Sheet = Summary_File.add_sheet("Well Map")
    Output_Sheet.write(0,0, "Run Name")
    for i,name in enumerate(IDs):
        Output_Sheet.write(i+1,0, name)
    for i in range(JMP_Sheet.ncols):
        for j, Cell in enumerate(JMP_Sheet.col_values(i)):
            Output_Sheet.write(j, i + 2, Cell)
    Summary_File.save(Folder_Name+"/Summary.xls")


#Outputs a CSV that is usable by Epmotion
def Epmotion_Output(Info,Purpose, Folder_Name):
    #Header
    Header_Data = [["Rack","Source","Rack","Destination","Volume","Tool"]]

    if (Purpose == "PLATE"):
        for i in range(len(Info)):
            #Creates a CSV file and feeds in the new data for Epmotion
            name = Folder_Name + "/EpMotion/Epmotion_Plate_"+ str(i+1) + "_SR.csv"
            with open(name, "w") as csvFile:
                writer = csv.writer(csvFile)
                writer.writerows(Header_Data)
                writer.writerows(Info[i].EdgeData)
                writer.writerows(Info[i].Commands)
            csvFile.close()
    else:
        name = Folder_Name + "/EpMotion/Dilution_Commands_SR.csv"
        with open(name, "w") as csvFile:
            writer = csv.writer(csvFile)
            writer.writerows(Header_Data)
            writer.writerows(Info)
        csvFile.close()

#Outputs a written protocol for original rack placement
def Protcol_Output(Dilutions_Num, Source, Rack_Layout, Folder_Name, Needed_Vol):
    name = Folder_Name + "/Protocol_SR.txt"
    File =  open(name,"w")

    File.write("Epmotion Protocol\n\n")
    File.write("DILUTION RACK PLACEMENT \n\n")
    File.write("1. Place a reservoir into the EpMotion\n")
    if Rack_Layout:
        File.write("2. Place 2 24-well racks into the EpMotion\n")
    else:
        File.write("2. Place a 24-well rack and a 96-well plate into the EpMotion\n")
    File.write("3. Place a box of 10 and 50 ul tips into the EpMotion\n")
    File.write("4. Ensure that there is a space for a plate in the EpMotion\n")
    File.write("5. Place a TS_10 and a TS_50 into the EpMotion\n\n")

    File.write("MANUAL DILUTIONS \n\n")
    for i in range(len(Source)):
        File.write("%d. Dilute Stock %s by adding %.2f ul into %.2f ul of Media\n" % (i + 1, Source[i][0], Needed_Vol[i], 1500 - Needed_Vol[i]))
    print("\n\n")

    File.write("LIQUID LAYOUT \n\n")
    File.write("1. Place a boat containing dilution liquid into the 1st slot in the reservoir\n")
    File.write("2. Place a boat containing edge liquid into the 2nd slot in the reservoir\n")
    for i in range(len(Source)):
        File.write("%d. Place the Diluted %s into the %s well in the first rack\n" % ( i+3, Source[i][0], Source[i][1]))
    if Rack_Layout:
        File.write("%d. Place %d sterile Epitubes into the second rack from %s to %s\n\n" %(len(Source) + 3, Dilutions_Num, Rack_Layout[0], Rack_Layout[Dilutions_Num-1]))
    else:
        File.write("\n")

    File.write("EPBLUE PROTOCOL \n\n")
    File.write("1. Create a new application.\n")
    File.write("2. Insert the Sample transfer command.\n")
    File.write("3. Under the Parameter Option, place the Stock 24-well as Source 1\n")
    File.write("4. Under the Parameter Option, place the reservoir as Source 2\n")
    if Rack_Layout:
        File.write("5. Under the Parameter Option, place the Dilution 24-well as Source 3\n")
        File.write("6. Under the Parameter Option, place the Dilution 24-well as Destination 1\n")
    else:
        File.write("5. Under the Parameter Option, place the Dilution 96-well plate as Source 3\n")
        File.write("6. Under the Parameter Option, place the Dilution 96-well plate as Destination 1\n")
    File.write("7. Under the Parameter Option, place the Plate as Destination 2\n")
    File.write("8. Ensure all other setting in the transfer command are satisfactory\n")
    File.write("9. Click the Sample transfer command.\n")
    File.write("10. Click on the CSV Command in top bar.\n")
    File.write("11. Select the Dilution_Command CSV and hit open.\n")
    File.write("12. Select the Pipette option and hit OK\n")
    File.write("13. Once finished repeat step 11 with the CSV required for each plate. \n\n")

    File.write("EXPERIMENT TIP PLACEMENT \n\n")
    File.write("1. Place a TS_50 and a TS_300 into the EpMotion\n")
    File.write("2. Place a box of 50 and 300 ul tips into the EpMotion\n" )

    File.close()
