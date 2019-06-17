from datetime import datetime
import csv
import sys
import shutil
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
            name = Folder_Name / ("EpMotion/Epmotion_Plate_"+ str(i+1) + "_SR.csv")
            with open(name, "w") as csvFile:
                writer = csv.writer(csvFile)
                writer.writerows(Header_Data)
                writer.writerows(Info[i].EdgeData)
                writer.writerows(Info[i].Commands)
            csvFile.close()
    else:
        name = Folder_Name / "EpMotion/Dilution_Commands_SR.csv"
        with open(name, "w") as csvFile:
            writer = csv.writer(csvFile)
            writer.writerows(Header_Data)
            writer.writerows(Info)
        csvFile.close()

#Outputs a written protocol for original rack placement
def Protcol_Output(Dilutions_Num, Source, Rack_Layout, Folder_Name, Needed_Vol, Media_Vol_Needed):
    name = Folder_Name / "Protocol_SR.txt"
    File =  open(name,"w")

    File.write("Epmotion Protocol\n\n")
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
        if Needed_Vol[i][0] < 1:
            Initial_Vol = Needed_Vol[i][0]
            File.write("%d a). Dilute Stock %s by adding %.2f ul into %.2f ul of Media\n" % (i + 1, Factor[0], Needed_Vol[i][0], Needed_Vol[i][1] - Needed_Vol[i][0]))
            File.write("%d b). Dilute Stock %s by adding %.2f ul into %.2f ul of Media\n" % (i + 1, Factor[0], Needed_Vol[i][0], Needed_Vol[i][1] - Needed_Vol[i][0]))
        else:
            File.write("%d. Dilute Stock %s by adding %.2f ul into %.2f ul of Media\n" % (i + 1, Factor[0], Needed_Vol[i][0], Needed_Vol[i][1] - Needed_Vol[i][0]))
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
    if Rack_Layout.all():
        File.write("5. Under the Parameter Option, place the Dilution 24-well as Source 3\n")
        File.write("6. Under the Parameter Option, place the Dilution 24-well as Destination 1\n")
    else:
        File.write("5. Under the Parameter Option, place the Dilution 96-well plate as Source 3\n")
        File.write("6. Under the Parameter Option, place the Dilution 96-well plate as Destination 1\n")
    File.write("7. Under the Parameter Option, place the Plate as Destination 2\n")  #Ensure that thIS IS ONLY FOR PLATE TRANSFER
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
