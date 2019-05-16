from tkinter import *
import numpy as np
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
import os


class GUI(object):

    def __init__(self, Master):
        self.row = 0
        self.Files = [[]]
        self.Files.append([])
        self.PlateType = []
        self.Volume = []
        self.Win = Master

    #Button Functions
    def End(self):
        self.Win.quit()

    def File(self,FileName, Type):
        File = askopenfilename()
        FileName["text"] = os.path.basename(File)
        self.Files[0].append(File)

    #Function to create headings and buttons
    def Organizer(self, Heading, Type):

        #Create a Prompt and a Frame
        Prompt = Label(self.Win, text = Heading)
        F = Frame(self.Win)

        #Display the prompt
        Prompt.grid(row = self.row, column = 0,pady = 5, padx = 100)
        self.row += 1

        #IF this is an FILE button, follow this code to display it
        if (Type == "FILE"):
            FileName = Button(F,text = "File Name", command = lambda: self.File(FileName,Type))
            FileName.grid(row = 0, column = 2)

        #IF this is a plate, run this code to display the menu
        elif (Type == "PLATE"):
            Formats = [ "96 Well", "384 Well", "24 Well"]
            Format_Var = StringVar(self.Win)
            Format_Var.set("96 Well") # default value
            self.PlateType.append(Format_Var)
            Plate_Type = OptionMenu(F, Format_Var, *Formats)
            Plate_Type.grid(row = 0, column = 2)

        #If this is a entry box for volumes, run this code.
        else:
            Volume_Var = Entry(F)
            Volume_Var.grid(row = 0, column = 2)
            self.Volume.append(Volume_Var)

        return F

    def run(self):

        self.Win.title("JMP to Epmotion")

        self.Input = self.Organizer("Please select the JMP Excel Output", "FILE")
        self.Input.grid(row = self.row, column = 0)
        self.row += 1

        self.Plate = self.Organizer("Please select the type of Plate", "PLATE")
        self.Plate.grid(row = self.row, column = 0)
        self.row += 1

        self.Vol = self.Organizer("What is the Volume of Diluted Factor Needed (uL) ?", "VOL")
        self.Vol.grid(row = self.row, column = 0)
        self.row += 1

        self.Vol = self.Organizer("What is the Volume of each Factor added to the plate (uL) ?", "VOL")
        self.Vol.grid(row = self.row, column = 0)
        self.row += 1

        #Done Button
        self.Done = Button(self.Win, text = "Done", command = self.End)
        self.Done.grid(row = self.row, column = 0, pady = 50)
        self.row += 1


def Get_Data():

    Root = Tk()
    Program = GUI(Root)
    Program.run()
    Root.mainloop()

    JMP_Excel = Program.Files[0][0]
    Plate_Format = Program.PlateType[0].get()
    Dil_Volume = Program.Volume[0].get()
    Add_Volume = Program.Volume[1].get()

    #Add code submiting for additional commands

    Root.destroy()

    return( JMP_Excel, Plate_Format, Dil_Volume, Add_Volume)
