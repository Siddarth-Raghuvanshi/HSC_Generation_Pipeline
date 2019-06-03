from datetime import datetime
import csv
import sys
import os
import shutil
from openpyxl import load_workbook
import xlrd

#Produces a collection of output folders and files which are required
def Produce_Output_Folder():
    Folder_Name = os.getcwd()+"/Experiment_Files_" + str(datetime.now()) + "_SR"
    os.mkdir(Folder_Name)
    os.mkdir(Folder_Name+"/EpMotion")
    os.mkdir(Folder_Name+"/Cytoflex")
    os.rename("Dilution_Concentrations_SR.csv", Folder_Name+"/EpMotion/Dilution_Concentrations_SR.csv")

    return Folder_Name

def Experiment_Summary(Folder_Name):
    shutil.copy("Automation Summary Template.xlsx", Folder_Name+"/Summary.xlsx")

    Summary = load_workbook(Folder_Name+"/Summary.xlsx")

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
def Protcol_Output(Dilutions_Num, Source, Rack_Layout, Folder_Name):
    name = Folder_Name + "/Protocol_SR.txt"
    File =  open(name,"w")

    File.write("Epmotion Protocol\n\n")
    File.write("RACK PLACEMENT \n\n")
    File.write("1. Place a reservoir into the EpMotion\n")
    File.write("2. Place 2 24-well racks into the EpMotion\n")
    File.write("3. Place a box of 50 and 300 ul tips into the EpMotion\n")
    File.write("4. Ensure that there is a space for a plate in the EpMotion\n")
    File.write("5. Place a TS_50 and a TS_300 into the EpMotion\n\n")

    File.write("LIQUID LAYOUT \n\n")
    File.write("1. Place a boat containing dilution liquid into the 1st slot in the reservoir\n")
    File.write("2. Place a boat containing edge liquid into the 2nd slot in the reservoir\n")
    for i in range(len(Source)):
        File.write("%d. Place the Stock %s into the %s well in the first rack\n" % ( i+3, Source[i][0], Source[i][1]))
    File.write("%d. Place %d sterile Epitubes into the second rack from %s to %s\n\n" %(len(Source) + 3, Dilutions_Num, Rack_Layout[0], Rack_Layout[Dilutions_Num-1]))

    File.write("EPBLUE PROTOCOL \n\n")
    File.write("1. Create a new application.\n")
    File.write("2. Insert the Sample transfer command.\n")
    File.write("3. Under the Parameter Option, place the Stock 24-well as Source 1\n")
    File.write("4. Under the Parameter Option, place the reservoir as Source 2\n")
    File.write("5. Under the Parameter Option, place the Dilution 24-well as Source 3\n")
    File.write("6. Under the Parameter Option, place the Dilution 24-well as Destination 1\n")
    File.write("7. Under the Parameter Option, place the Plate as Destination 2\n")
    File.write("8. Ensure all other setting in the transfer command are satisfactory\n")
    File.write("9. Click the Sample transfer command.\n")
    File.write("10. Click on the CSV Command in top bar.\n")
    File.write("11. Select the Dilution_Command CSV and hit open.\n")
    File.write("12. Select the Pipette option and hit OK\n")
    File.write("13. Once finished repeat step 11 with the CSV required for each plate.\n")

    File.close()